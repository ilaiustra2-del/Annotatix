using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace PluginsManager.Commands
{
    public class ClashResolveSession
    {
        // ----------------------------------------------------------------
        // Win32 P/Invoke
        // ----------------------------------------------------------------
        [DllImport("user32.dll")] private static extern short GetKeyState(int nVirtKey);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern int  GetWindowThreadProcessId(IntPtr hWnd, out int pid);
        [DllImport("user32.dll")] [return:MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc fn, IntPtr hMod, uint threadId);
        [DllImport("user32.dll")] [return:MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet=CharSet.Auto)] private static extern IntPtr GetModuleHandle(string name);

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN     = 0x0100;
        private const int WM_SYSKEYDOWN  = 0x0104;
        private const int VK_ESCAPE      = 0x1B;
        private const int VK_RETURN      = 0x0D;
        private const int VK_CONTROL     = 0x11;
        private static bool IsCtrlDown() => (GetKeyState(VK_CONTROL) & 0x8000) != 0;

        // Returns true when the foreground window belongs to the Revit process.
        // Uses IsChild so docked/undocked viewports all pass.
        private static bool IsRevitFocused(IntPtr revitHwnd)
        {
            if (revitHwnd == IntPtr.Zero) return false;
            IntPtr fg = GetForegroundWindow();
            if (fg == IntPtr.Zero) return false;
            if (fg == revitHwnd || IsChild(revitHwnd, fg)) return true;
            // Fallback: same process id (handles tear-off dialogs)
            GetWindowThreadProcessId(fg, out int fgPid);
            return fgPid == Process.GetCurrentProcess().Id;
        }

        // ----------------------------------------------------------------
        // Win32 low-level keyboard hook — Escape / Enter
        // ----------------------------------------------------------------
        private IntPtr              _llHook;
        private LowLevelKeyboardProc _llProc; // keep delegate alive
        private IntPtr              _revitHwnd;

        private void InstallKeyboardHook()
        {
            try
            {
                _revitHwnd = _uiApp?.MainWindowHandle ?? IntPtr.Zero;
                _llProc    = LowLevelKeyboardCallback;
                using (var mod = Process.GetCurrentProcess().MainModule)
                    _llHook = SetWindowsHookEx(WH_KEYBOARD_LL, _llProc, GetModuleHandle(mod.ModuleName), 0);

                if (_llHook != IntPtr.Zero)
                    DebugLogger.Log("[CLASH-SESSION] Win32 keyboard hook installed");
                else
                    DebugLogger.Log("[CLASH-SESSION] SetWindowsHookEx failed");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-SESSION] InstallKeyboardHook error: {ex.Message}");
            }
        }

        private void UninstallKeyboardHook()
        {
            try
            {
                if (_llHook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_llHook);
                    _llHook = IntPtr.Zero;
                    _llProc = null;
                    DebugLogger.Log("[CLASH-SESSION] Keyboard hook removed");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-SESSION] UninstallKeyboardHook error: {ex.Message}");
            }
        }

        private IntPtr LowLevelKeyboardCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && Current != null
                && ((int)wParam == WM_KEYDOWN || (int)wParam == WM_SYSKEYDOWN)
                && IsRevitFocused(_revitHwnd))
            {
                int vk = Marshal.ReadInt32(lParam);
                if (vk == VK_ESCAPE)
                {
                    DebugLogger.Log("[CLASH-SESSION] Escape pressed — deactivating");
                    var session = Current;
                    Current = null;
                    session._sequentialPicks.Clear();
                    // Must deactivate on UI thread
                    var dispatcher = System.Windows.Application.Current?.Dispatcher
                        ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;
                    dispatcher.BeginInvoke(new Action(() => session.Deactivate()),
                        System.Windows.Threading.DispatcherPriority.Normal);
                    // Block the key so Revit doesn’t also exit its command
                    return (IntPtr)1;
                }
                else if (vk == VK_RETURN)
                {
                    DebugLogger.Log("[CLASH-SESSION] Enter pressed — triggering Resolve");
                    var ctrl = _optionsBarWpfControl;
                    var dispatcher = System.Windows.Application.Current?.Dispatcher
                        ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;
                    dispatcher.BeginInvoke(new Action(() =>
                    {
                        var m = ctrl?.GetType().GetMethod("TriggerResolve",
                            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                        m?.Invoke(ctrl, null);
                    }), System.Windows.Threading.DispatcherPriority.Normal);
                    return (IntPtr)1;
                }
            }
            return CallNextHookEx(_llHook, nCode, wParam, lParam);
        }

        // ----------------------------------------------------------------
        // Singleton
        // ----------------------------------------------------------------
        public static ClashResolveSession Current { get; private set; }

        // ----------------------------------------------------------------
        // Fields
        // ----------------------------------------------------------------
        private UIApplication _uiApp;
        private ExternalEvent  _execEvent;
        private ClashResolveExecuteHandler _execHandler;

        // Sequential Ctrl-picks: index 0 = Pipe A, index 1 = Pipe B
        private readonly List<ElementId> _sequentialPicks = new List<ElementId>();
        private ICollection<ElementId>   _prevSelection   = new HashSet<ElementId>();

        // Options bar
        private object _savedOptionsBarContent;
        private object _optionsBarControl;          // container element (for restore)
        private Action<object> _optionsBarSetContent; // setter for restore
        private object _optionsBarWpfControl;   // our ClashResolveOptionsBar (loaded via reflection)
        private System.Windows.FrameworkElement _dialogBarFe; // DialogBarControl reference
        private System.Windows.Visibility _savedDialogBarVisibility;
        private double _savedDialogBarHeight;

        // ----------------------------------------------------------------
        // Activate / Deactivate
        // ----------------------------------------------------------------

        public static void Activate(UIApplication uiApp, ExternalEvent execEvent,
            ClashResolveExecuteHandler execHandler)
        {
            if (Current != null)
            {
                DebugLogger.Log("[CLASH-SESSION] Already active — re-activating");
                Current.Deactivate();
            }

            var session = new ClashResolveSession();
            Current = session;
            session.DoActivate(uiApp, execEvent, execHandler);
        }

        public static void DeactivateCurrent()
        {
            Current?.Deactivate();
            Current = null;
        }

        private void DoActivate(UIApplication uiApp, ExternalEvent execEvent,
            ClashResolveExecuteHandler execHandler)
        {
            _uiApp       = uiApp;
            _execEvent   = execEvent;
            _execHandler = execHandler;

            _sequentialPicks.Clear();
            _prevSelection = new HashSet<ElementId>();

            InjectOptionsBar();

            DebugLogger.Log("[CLASH-SESSION] Session activated");
        }

        public void Deactivate()
        {
            RestoreOptionsBar();
            DebugLogger.Log("[CLASH-SESSION] Session deactivated");
        }

        // ----------------------------------------------------------------
        // Options Bar injection via AdWindows ComponentManager
        // ----------------------------------------------------------------
        private void InjectOptionsBar()
        {
            try
            {
                // Inject our WPF control into Revit's Options Bar via visual tree.
                // AdWindows ComponentManager does not expose OptionsBar in Revit 2024+;
                // the Options Bar is rendered as a WPF ToolBar below the ribbon.
                System.Windows.Threading.Dispatcher uiDispatcher =
                    System.Windows.Application.Current?.Dispatcher
                    ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;

                uiDispatcher.Invoke(() =>
                {
                    // Load ClashResolveOptionsBar from ClashResolve.Module assembly
                    var moduleAsm = DynamicModuleLoader.GetModuleAssembly("clash_resolve");
                    if (moduleAsm == null)
                    {
                        DebugLogger.Log("[CLASH-SESSION] ClashResolve module assembly not found — cannot inject OptionsBar");
                        return;
                    }

                    var ctrlType = moduleAsm.GetType("ClashResolve.Module.UI.ClashResolveOptionsBar");
                    if (ctrlType == null)
                    {
                        DebugLogger.Log("[CLASH-SESSION] ClashResolveOptionsBar type not found");
                        return;
                    }

                    _optionsBarWpfControl = Activator.CreateInstance(ctrlType);

                    // Subscribe to events
                    var resolveEvent = ctrlType.GetEvent("ResolveRequested");
                    var autoEvent    = ctrlType.GetEvent("AutoRecalcRequested");
                    var exitEvent    = ctrlType.GetEvent("ExitRequested");

                    resolveEvent?.AddEventHandler(_optionsBarWpfControl,
                        Delegate.CreateDelegate(resolveEvent.EventHandlerType, this,
                            typeof(ClashResolveSession).GetMethod(nameof(OnResolveRequested),
                                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)));

                    autoEvent?.AddEventHandler(_optionsBarWpfControl,
                        Delegate.CreateDelegate(autoEvent.EventHandlerType, this,
                            typeof(ClashResolveSession).GetMethod(nameof(OnAutoRecalcRequested),
                                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)));

                    exitEvent?.AddEventHandler(_optionsBarWpfControl,
                        Delegate.CreateDelegate(exitEvent.EventHandlerType, this,
                            typeof(ClashResolveSession).GetMethod(nameof(OnExitRequested),
                                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)));

                    bool injected = false;

                    // Find the Options Bar container via visual tree
                    var (container, setContent, getContent, dialogBarFe) = FindOptionsBarContainer();
                    if (container != null && setContent != null)
                    {
                        _optionsBarControl    = container;
                        _optionsBarSetContent = setContent;
                        _savedOptionsBarContent = getContent?.Invoke();
                        DebugLogger.Log($"[CLASH-SESSION] Container type: {container.GetType().FullName}");
                        DebugLogger.Log($"[CLASH-SESSION] Saved content type: {_savedOptionsBarContent?.GetType().FullName ?? "null"}");
                        DebugLogger.Log($"[CLASH-SESSION] Our control type: {_optionsBarWpfControl?.GetType().FullName ?? "null"}");
                        try
                        {
                            setContent(_optionsBarWpfControl);
                            // Make the DialogBarControl visible — Revit hides it when no native command is active
                            if (dialogBarFe != null)
                            {
                                _savedDialogBarVisibility = dialogBarFe.Visibility;
                                _savedDialogBarHeight = dialogBarFe.Height;
                                dialogBarFe.Visibility = System.Windows.Visibility.Visible;
                                if (double.IsNaN(dialogBarFe.Height) || dialogBarFe.Height < 28)
                                    dialogBarFe.Height = 28;
                                _dialogBarFe = dialogBarFe;
                                DebugLogger.Log($"[CLASH-SESSION] DialogBarControl made visible, H={dialogBarFe.Height}");
                            }
                            var verifyContent = getContent?.Invoke();
                            DebugLogger.Log($"[CLASH-SESSION] After inject — content is: {verifyContent?.GetType().FullName ?? "null"}");
                            injected = true;
                        }
                        catch (Exception setEx)
                        {
                            DebugLogger.Log($"[CLASH-SESSION] setContent threw: {setEx.GetType().Name}: {setEx.Message}");
                        }
                        if (injected)
                        {
                            DebugLogger.Log("[CLASH-SESSION] OptionsBar injected successfully");
                            // Install MainWindow PreviewKeyDown handler for Escape / Enter
                            InstallKeyboardHook();
                        }
                    }

                    if (!injected)
                        DebugLogger.Log("[CLASH-SESSION] WARNING: Could not inject OptionsBar — all paths failed");
                });
            }
            catch (Exception ex)
            {
                var inner = ex.InnerException ?? ex;
                DebugLogger.Log($"[CLASH-SESSION] InjectOptionsBar error: {inner.GetType().Name}: {inner.Message}");
                if (inner.InnerException != null)
                    DebugLogger.Log($"[CLASH-SESSION]   inner: {inner.InnerException.Message}");
                DebugLogger.Log($"[CLASH-SESSION]   stack: {inner.StackTrace?.Split('\n')[0]}");
            }
        }

        private void RestoreOptionsBarCore()
        {
            // Execute the restore logic — must be called on the WPF UI thread.
            if (_optionsBarSetContent == null) return;
            _optionsBarSetContent(_savedOptionsBarContent);
            if (_dialogBarFe != null)
            {
                _dialogBarFe.Visibility = _savedDialogBarVisibility;
                _dialogBarFe.Height = _savedDialogBarHeight;
                _dialogBarFe = null;
            }
            UninstallKeyboardHook();
            DebugLogger.Log("[CLASH-SESSION] OptionsBar restored");
            _optionsBarSetContent = null;
        }

        private void RestoreOptionsBar()
        {
            try
            {
                if (_optionsBarSetContent == null) return;

                var uiDispatcher =
                    System.Windows.Application.Current?.Dispatcher
                    ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;

                if (uiDispatcher.CheckAccess())
                {
                    // Already on UI thread (e.g. button click) — call directly
                    RestoreOptionsBarCore();
                }
                else
                {
                    // Called from background/Revit thread — marshal to UI thread
                    uiDispatcher.Invoke(() => RestoreOptionsBarCore());
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-SESSION] RestoreOptionsBar error: {ex.Message}");
            }
        }

        /// <summary>
        /// Walks the Revit main window visual tree to find the WPF ToolBar
        /// that Revit uses to render the Options Bar (the thin strip below the ribbon).
        /// </summary>
        private static System.Windows.Window GetRevitMainWindow()
        {
            // Try WPF Application.Current.MainWindow first
            var wpfMain = System.Windows.Application.Current?.MainWindow;
            if (wpfMain != null) return wpfMain;

            // Fallback: find Revit main window via HwndSource
            try
            {
                var handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                if (handle != IntPtr.Zero)
                {
                    var src = System.Windows.Interop.HwndSource.FromHwnd(handle);
                    if (src?.RootVisual is System.Windows.Window w) return w;
                }
            }
            catch { }

            // Last resort: find any visible Window
            if (System.Windows.Application.Current != null)
            {
                foreach (System.Windows.Window w in System.Windows.Application.Current.Windows)
                {
                    if (w.IsVisible) return w;
                }
            }
            return null;
        }

        // Returns the OptionsBar container to inject our control into,
        // along with a setter action to set its content.
        private static (System.Windows.DependencyObject container, Action<object> setContent, Func<object> getContent, System.Windows.FrameworkElement dialogBarFe) FindOptionsBarContainer()
        {
            try
            {
                var mainWin = GetRevitMainWindow();
                if (mainWin == null) return (null, null, null, null);

                var allElements = FindAllVisualChildren<System.Windows.FrameworkElement>(mainWin);

                // Step 1: Find the correct DialogBarControl — the one below the ribbon (not the status bar)
                // Revit has two: one under the ribbon (~Y 130-200), one in the status bar at the bottom.
                // We want the one with the smallest Y coordinate (closest to the top of the window).
                var allDialogBars = new List<(System.Windows.FrameworkElement el, double y)>();
                foreach (var el in allElements)
                {
                    if (el.GetType().Name == "DialogBarControl")
                    {
                        try
                        {
                            var pt = el.TranslatePoint(new System.Windows.Point(0, 0), mainWin);
                            allDialogBars.Add((el, pt.Y));
                            DebugLogger.Log($"[CLASH-SESSION] DialogBarControl found: W={el.ActualWidth:F0} H={el.ActualHeight:F0} Y={pt.Y:F0} Visibility={el.Visibility}");
                        }
                        catch
                        {
                            allDialogBars.Add((el, double.MaxValue));
                        }
                    }
                }

                if (allDialogBars.Count == 0)
                {
                    DebugLogger.Log("[CLASH-SESSION] DialogBarControl not found in visual tree");
                    return (null, null, null, null);
                }

                // Pick the topmost one (below ribbon, not the status bar at the bottom)
                allDialogBars.Sort((a, b) => a.y.CompareTo(b.y));
                System.Windows.FrameworkElement dialogBar = allDialogBars[0].el;
                DebugLogger.Log($"[CLASH-SESSION] Using DialogBarControl at Y={allDialogBars[0].y:F0}");

                // Step 2: Try to overlay our WPF control directly inside DialogBarControl
                // ControlHost is WinForms-based and cannot host WPF UserControl directly.
                // Instead, wrap DialogBarControl content in a Grid and add our control on top.
                var dialogBarType = dialogBar.GetType();
                DebugLogger.Log($"[CLASH-SESSION] DialogBarControl base types: {dialogBarType.BaseType?.FullName}");

                // Check if DialogBarControl is a Panel (can use Children)
                var childrenProp = dialogBarType.GetProperty("Children",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (childrenProp != null)
                {
                    DebugLogger.Log("[CLASH-SESSION] DialogBarControl has Children property (Panel)");
                    var children = childrenProp.GetValue(dialogBar) as System.Windows.Controls.UIElementCollection;
                    if (children != null)
                        return (
                            dialogBar,
                            val =>
                            {
                                if (val is System.Windows.UIElement uie) children.Add(uie);
                            },
                            () => null,
                            dialogBar);
                }

                // Try ContentControl.Content on DialogBarControl itself
                var contentProp = dialogBarType.GetProperty("Content",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (contentProp != null && contentProp.CanWrite)
                {
                    DebugLogger.Log("[CLASH-SESSION] DialogBarControl.Content is writable — using directly");
                    return (
                        dialogBar,
                        val => contentProp.SetValue(dialogBar, val),
                        () => contentProp.GetValue(dialogBar),
                        dialogBar);
                }

                // Find ControlHost inside DialogBarControl — it holds the actual content
                var controlHost = FindVisualChildByTypeName(dialogBar, "ControlHost");
                if (controlHost != null)
                {
                    DebugLogger.Log($"[CLASH-SESSION] Found ControlHost: {controlHost.GetType().FullName}");
                    int n = System.Windows.Media.VisualTreeHelper.GetChildrenCount(controlHost);
                    for (int i = 0; i < n; i++)
                    {
                        var ch = System.Windows.Media.VisualTreeHelper.GetChild(controlHost, i);
                        var chn = ch.GetType().Name;
                        var chname = (ch as System.Windows.FrameworkElement)?.Name ?? "";
                        DebugLogger.Log($"[CLASH-SESSION]   ControlHost child[{i}]: {chn} Name='{chname}'");
                    }
                    // Log all public properties that could hold content
                    foreach (var propName in new[] { "Content", "Child", "Control", "HostedControl", "Element" })
                    {
                        var p = controlHost.GetType().GetProperty(propName,
                            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                        if (p != null)
                            DebugLogger.Log($"[CLASH-SESSION]   ControlHost.{propName} exists, CanWrite={p.CanWrite}, value={p.GetValue(controlHost)?.GetType().Name ?? "null"}");
                    }
                    var setter = TryGetContentSetter(controlHost);
                    if (setter.setContent != null)
                        return (controlHost, setter.setContent, setter.getContent, dialogBar);
                }

                // Step 3: Fallback — ContentPresenter directly inside DialogBarControl
                var cp = FindVisualChildByTypeName(dialogBar, "ContentPresenter");
                if (cp != null)
                {
                    DebugLogger.Log("[CLASH-SESSION] Using ContentPresenter inside DialogBarControl");
                    var setter = TryGetContentSetter(cp);
                    if (setter.setContent != null)
                        return (cp, setter.setContent, setter.getContent, dialogBar);
                }

                // Step 4: Set content directly on DialogBarControl itself
                {
                    DebugLogger.Log("[CLASH-SESSION] Using DialogBarControl itself as container");
                    var setter = TryGetContentSetter(dialogBar);
                    if (setter.setContent != null)
                        return (dialogBar, setter.setContent, setter.getContent, dialogBar);
                }

                DebugLogger.Log("[CLASH-SESSION] Could not find settable container inside DialogBarControl");
                return (null, null, null, null);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-SESSION] FindOptionsBarContainer error: {ex.Message}");
                return (null, null, null, null);
            }
        }

        private static (Action<object> setContent, Func<object> getContent) TryGetContentSetter(System.Windows.DependencyObject obj)
        {
            var type = obj.GetType();
            foreach (var propName in new[] { "Content", "Child", "DataContext" })
            {
                var prop = type.GetProperty(propName,
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (prop != null && prop.CanWrite)
                {
                    return (
                        val => prop.SetValue(obj, val),
                        () => prop.GetValue(obj)
                    );
                }
            }
            // For ItemsControl (ToolBar, ToolBarTray, etc.) — use Items collection
            var itemsProp = type.GetProperty("Items",
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            if (itemsProp != null)
            {
                var items = itemsProp.GetValue(obj) as System.Collections.IList;
                if (items != null)
                    return (
                        val => { items.Clear(); if (val != null) items.Add(val); },
                        () => items.Count > 0 ? items[0] : null
                    );
            }
            return (null, null);
        }

        private static System.Windows.DependencyObject FindVisualChildByName(
            System.Windows.DependencyObject parent, string name)
        {
            if (parent == null) return null;
            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is System.Windows.FrameworkElement fe && fe.Name == name)
                    return child;
                var result = FindVisualChildByName(child, name);
                if (result != null) return result;
            }
            return null;
        }

        private static System.Windows.DependencyObject FindVisualChildByTypeName(
            System.Windows.DependencyObject parent, string typeName)
        {
            if (parent == null) return null;
            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child.GetType().Name == typeName) return child;
                var result = FindVisualChildByTypeName(child, typeName);
                if (result != null) return result;
            }
            return null;
        }

        // Legacy: keep FindOptionsBarToolBar for RestoreOptionsBar compatibility
        private static System.Windows.Controls.ToolBar FindOptionsBarToolBar()
        {
            return null; // replaced by FindOptionsBarContainer
        }

        private static System.Collections.Generic.List<T> FindAllVisualChildren<T>(System.Windows.DependencyObject parent)
            where T : System.Windows.DependencyObject
        {
            var result = new System.Collections.Generic.List<T>();
            if (parent == null) return result;
            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T typed) result.Add(typed);
                result.AddRange(FindAllVisualChildren<T>(child));
            }
            return result;
        }

        // ----------------------------------------------------------------
        // Selection tracking (called by ClashResolveSelectionHandler)
        // ----------------------------------------------------------------
        public void OnSelectionChanged(UIApplication uiApp)
        {
            try
            {
                var uidoc = uiApp.ActiveUIDocument;
                if (uidoc == null) return;

                var currentSelection = new HashSet<ElementId>(uidoc.Selection.GetElementIds());

                // New selection logic:
                // - If user has exactly 2 elements selected (any combination with/without Ctrl),
                //   we treat the one that was already in _sequentialPicks[0] as Pipe A,
                //   and the new one as Pipe B.
                // - If 1 element selected: record as first pick.
                // - If 2 elements selected: the one added last becomes Pipe B.

                var added   = currentSelection.Where(id => !_prevSelection.Contains(id)).ToList();
                var removed = _prevSelection.Where(id => !currentSelection.Contains(id)).ToList();

                // Remove deselected from sequential list
                foreach (var id in removed)
                    _sequentialPicks.Remove(id);

                // Only process single-element additions (ignore box/multi-select)
                if (added.Count == 1)
                {
                    var id = added[0];
                    if (!_sequentialPicks.Contains(id))
                    {
                        if (_sequentialPicks.Count >= 2)
                            _sequentialPicks.Clear();
                        _sequentialPicks.Add(id);
                        DebugLogger.Log($"[CLASH-SESSION] Sequential pick #{_sequentialPicks.Count}: ID={id.Value}");
                    }
                }
                else if (added.Count > 1)
                {
                    // Box/multi-select: clear picks so user must re-select manually
                    _sequentialPicks.Clear();
                    DebugLogger.Log($"[CLASH-SESSION] Multi-select ({added.Count} elements) — picks cleared");
                }

                // If selection was cleared, reset
                if (currentSelection.Count == 0)
                    _sequentialPicks.Clear();

                _prevSelection = currentSelection;

                UpdateHint(uiApp.ActiveUIDocument.Document);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-SESSION] OnSelectionChanged error: {ex.Message}");
            }
        }

        // ----------------------------------------------------------------
        // Options Bar callbacks
        // ----------------------------------------------------------------
        private void OnResolveRequested(object paramsObj)
        {
            // paramsObj is ClashResolveOptionsBarParams from ClashResolve.Module
            // Validate sequential picks
            if (_sequentialPicks.Count != 2)
            {
                System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(() =>
                    MessageBox.Show(
                        "Выберите последовательно два элемента (по одному кликом). " +
                        "Первый элемент будет обходить второй.",
                        "Ошибка выбора элементов",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning));
                return;
            }

            var pipeAId = _sequentialPicks[0];
            var pipeBId = _sequentialPicks[1];

            if (pipeAId == pipeBId)
            {
                MessageBox.Show("Первый и второй выделенные элементы совпадают. Выделите два разных элемента.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Read params via reflection (paramsObj is from ClashResolve.Module assembly)
            var pt = paramsObj.GetType();
            bool useAngle45    = (bool)pt.GetProperty("UseAngle45").GetValue(paramsObj);
            double clearanceMm = (double)pt.GetProperty("ClearanceMm").GetValue(paramsObj);
            bool autoClear     = (bool)pt.GetProperty("AutoClearance").GetValue(paramsObj);
            double halfLength  = (double)pt.GetProperty("HalfLengthMm").GetValue(paramsObj);
            bool autoHalf      = (bool)pt.GetProperty("AutoHalfLength").GetValue(paramsObj);

            // Schedule execution via ExternalEvent
            _execHandler.PendingAction = (app) =>
            {
                var doc = app.ActiveUIDocument?.Document;
                if (doc == null) return;

                // Load ClashResolver via reflection
                var moduleAsm = DynamicModuleLoader.GetModuleAssembly("clash_resolve");
                if (moduleAsm == null)
                {
                    DebugLogger.Log("[CLASH-SESSION] ClashResolve module not loaded");
                    return;
                }

                var resolverType  = moduleAsm.GetType("ClashResolve.Module.Core.ClashResolver");
                var pairType      = moduleAsm.GetType("ClashResolve.Module.Core.ClashPair");
                var resolver      = Activator.CreateInstance(resolverType);
                var pair          = Activator.CreateInstance(pairType);

                pairType.GetProperty("PipeAId").SetValue(pair, pipeAId);
                pairType.GetProperty("PipeBId").SetValue(pair, pipeBId);
                pairType.GetProperty("ClearanceMm").SetValue(pair, clearanceMm);
                pairType.GetProperty("HalfLengthMm").SetValue(pair, halfLength);
                pairType.GetProperty("AutoClearance").SetValue(pair, autoClear);
                pairType.GetProperty("AutoHalfLength").SetValue(pair, autoHalf);
                pairType.GetProperty("UseAngle45").SetValue(pair, useAngle45);

                var resolveMethod = resolverType.GetMethod("ResolveClash");
                var result = resolveMethod.Invoke(resolver, new object[] { doc, pair });

                bool success = (bool)result.GetType().GetProperty("Success").GetValue(result);
                string msg   = (string)result.GetType().GetProperty("Message").GetValue(result);
                double usedClear = (double)result.GetType().GetProperty("UsedClearanceMm").GetValue(result);
                double usedHalf  = (double)result.GetType().GetProperty("UsedHalfLengthMm").GetValue(result);

                DebugLogger.Log($"[CLASH-SESSION] Resolve: Success={success}, Msg={msg}");

                // Clear sequential picks immediately after resolve — prevents stale IDs in next iteration
                _sequentialPicks.Clear();
                _prevSelection = new System.Collections.Generic.HashSet<ElementId>();

                // Update OptionsBar with used values on UI thread
                System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                {
                    try
                    {
                        var ctrlType = _optionsBarWpfControl?.GetType();
                        if (ctrlType == null) return;

                        if (autoClear && usedClear > 0)
                        {
                            var txtC = ctrlType.GetField("txtClearance",
                                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                                ?? (object)ctrlType.GetProperty("txtClearance");
                            // Access via FindName
                            var frameworkEl = _optionsBarWpfControl as System.Windows.FrameworkElement;
                            var tc = frameworkEl?.FindName("txtClearance") as System.Windows.Controls.TextBox;
                            if (tc != null) tc.Text = ((int)usedClear).ToString();
                        }
                        if (autoHalf && usedHalf > 0)
                        {
                            var frameworkEl = _optionsBarWpfControl as System.Windows.FrameworkElement;
                            var ts = frameworkEl?.FindName("txtSegmentLength") as System.Windows.Controls.TextBox;
                            if (ts != null) ts.Text = ((int)(usedHalf * 2)).ToString();
                        }
                    }
                    catch { }

                    // Show result
                    string icon = success ? "✓" : "✗";
                    MessageBox.Show($"{icon} {msg}",
                        success ? "Обход выполнен" : "Ошибка обхода",
                        MessageBoxButton.OK,
                        success ? MessageBoxImage.Information : MessageBoxImage.Warning);
                    // Reset hint after dialog closes
                    var fe2 = _optionsBarWpfControl as System.Windows.FrameworkElement;
                    var hint = fe2?.FindName("txtSelectionHint") as System.Windows.Controls.TextBlock;
                    if (hint != null) hint.Text = "Выделите трубу A";
                });
            };

            _execEvent.Raise();
        }

        private void OnAutoRecalcRequested()
        {
            TryUpdateAutoDisplay();
        }

        private void OnExitRequested()
        {
            DebugLogger.Log("[CLASH-SESSION] Exit button clicked — deactivating");
            // Do NOT use Dispatcher.BeginInvoke — that would deadlock inside RestoreOptionsBar's own Dispatcher.Invoke.
            try
            {
                RestoreOptionsBar();
                DebugLogger.Log("[CLASH-SESSION] Session deactivated via Exit button");
            }
            finally
            {
                Current = null;
            }
        }

        // ----------------------------------------------------------------
        // Auto-display recalculation
        // ----------------------------------------------------------------
        private void TryUpdateAutoDisplay()
        {
            try
            {
                if (_optionsBarWpfControl == null) return;
                if (_sequentialPicks.Count < 2) return;

                var doc = _uiApp?.ActiveUIDocument?.Document;
                if (doc == null) return;

                var elemA = doc.GetElement(_sequentialPicks[0]) as Autodesk.Revit.DB.MEPCurve;
                var elemB = doc.GetElement(_sequentialPicks[1]) as Autodesk.Revit.DB.MEPCurve;
                if (elemA == null || elemB == null) return;

                double rA = GetOuterRadiusMm(elemA);
                double rB = GetOuterRadiusMm(elemB);
                if (rA <= 0 || rB <= 0) return;

                double rMaxMm = Math.Max(rA, rB);

                // Call UpdateAutoDisplay on the WPF control via reflection
                _optionsBarWpfControl.GetType()
                    .GetMethod("UpdateAutoDisplay")
                    ?.Invoke(_optionsBarWpfControl, new object[] { rMaxMm });
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-SESSION] TryUpdateAutoDisplay: {ex.Message}");
            }
        }

        private static double GetOuterRadiusMm(Autodesk.Revit.DB.MEPCurve mep)
        {
            if (mep == null) return 0;
            var p = mep.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                 ?? mep.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
                 ?? mep.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            return (p != null && p.HasValue) ? p.AsDouble() / 2.0 * 304.8 : 0;
        }

        // ----------------------------------------------------------------
        // Hint label update
        // ----------------------------------------------------------------
        private void UpdateHint(Document doc)
        {
            try
            {
                string hint;
                switch (_sequentialPicks.Count)
                {
                    case 0:
                        hint = "Выделите трубу A — обходящую (Ctrl+клик)";
                        break;
                    case 1:
                        var nameA = GetElementName(doc, _sequentialPicks[0]);
                        hint = $"Труба A: {nameA} — теперь выделите трубу B (Ctrl+клик)";
                        break;
                    default:
                        var nA = GetElementName(doc, _sequentialPicks[0]);
                        var nB = GetElementName(doc, _sequentialPicks[1]);
                        hint = $"A: {nA}  |  B: {nB}  — нажмите «Перестроить»";
                        break;
                }

                System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                {
                    var frameworkEl = _optionsBarWpfControl as System.Windows.FrameworkElement;
                    var lbl = frameworkEl?.FindName("txtSelectionHint") as System.Windows.Controls.TextBlock;
                    if (lbl != null) lbl.Text = hint;
                });

                if (_sequentialPicks.Count == 2)
                    TryUpdateAutoDisplay();
            }
            catch { }
        }

        private static string GetElementName(Document doc, ElementId id)
        {
            var elem = doc?.GetElement(id);
            if (elem == null) return $"ID {id.Value}";
            string cat = elem.Category?.Name ?? elem.GetType().Name;
            return $"ID {id.Value} ({cat})";
        }
    }
}
