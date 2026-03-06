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
    /// <summary>
    /// Session for "Множественное исправление коллизий".
    /// Three-step workflow: collect pipe A ids → collect pipe B ids → configure params → batch resolve.
    /// Architecture mirrors ClashResolveSession (same Win32 hook, same OptionsBar injection).
    /// </summary>
    public class MultiClashSession
    {
        // ----------------------------------------------------------------
        // Win32 P/Invoke (identical to ClashResolveSession)
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

        private static bool IsRevitFocused(IntPtr revitHwnd)
        {
            if (revitHwnd == IntPtr.Zero) return false;
            IntPtr fg = GetForegroundWindow();
            if (fg == IntPtr.Zero) return false;
            if (fg == revitHwnd || IsChild(revitHwnd, fg)) return true;
            GetWindowThreadProcessId(fg, out int fgPid);
            return fgPid == Process.GetCurrentProcess().Id;
        }

        // ----------------------------------------------------------------
        // Keyboard hook fields
        // ----------------------------------------------------------------
        private IntPtr               _llHook;
        private LowLevelKeyboardProc _llProc;
        private IntPtr               _revitHwnd;

        private void InstallKeyboardHook()
        {
            try
            {
                _revitHwnd = _uiApp?.MainWindowHandle ?? IntPtr.Zero;
                _llProc    = LowLevelKeyboardCallback;
                using (var mod = Process.GetCurrentProcess().MainModule)
                    _llHook = SetWindowsHookEx(WH_KEYBOARD_LL, _llProc, GetModuleHandle(mod.ModuleName), 0);

                DebugLogger.Log(_llHook != IntPtr.Zero
                    ? "[MULTI-SESSION] Keyboard hook installed"
                    : "[MULTI-SESSION] SetWindowsHookEx failed");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MULTI-SESSION] InstallKeyboardHook error: {ex.Message}");
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
                    DebugLogger.Log("[MULTI-SESSION] Keyboard hook removed");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MULTI-SESSION] UninstallKeyboardHook error: {ex.Message}");
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
                    DebugLogger.Log("[MULTI-SESSION] Escape pressed — deactivating");
                    var session = Current;
                    Current = null;
                    var dispatcher = System.Windows.Application.Current?.Dispatcher
                        ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;
                    dispatcher.BeginInvoke(new Action(() => session.Deactivate()),
                        System.Windows.Threading.DispatcherPriority.Normal);
                    return (IntPtr)1;
                }
                else if (vk == VK_RETURN)
                {
                    DebugLogger.Log("[MULTI-SESSION] Enter pressed — triggering Done");
                    var ctrl = _optionsBarWpfControl;
                    var dispatcher = System.Windows.Application.Current?.Dispatcher
                        ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;
                    dispatcher.BeginInvoke(new Action(() =>
                    {
                        var m = ctrl?.GetType().GetMethod("TriggerDone",
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
        public static MultiClashSession Current { get; private set; }

        // ----------------------------------------------------------------
        // Fields
        // ----------------------------------------------------------------
        private UIApplication              _uiApp;
        private ExternalEvent              _execEvent;
        private ClashResolveExecuteHandler _execHandler;

        // Collected ids
        private List<ElementId> _pipeAIds = new List<ElementId>();
        private List<ElementId> _pipeBIds = new List<ElementId>();

        // Session-side step tracker — authoritative source (avoids UI-thread race in GetCurrentStep)
        private int _currentStep = 1;

        // Current Revit selection (for steps 1 & 2)
        private ICollection<ElementId> _currentSelection = new HashSet<ElementId>();

        // Options bar
        private object _savedOptionsBarContent;
        private object _optionsBarControl;
        private Action<object> _optionsBarSetContent;
        private Func<object>   _optionsBarGetContent;
        private object _optionsBarWpfControl;
        private System.Windows.FrameworkElement _dialogBarFe;
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
                DebugLogger.Log("[MULTI-SESSION] Already active — re-activating");
                Current.Deactivate();
            }

            var session = new MultiClashSession();
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

            _pipeAIds.Clear();
            _pipeBIds.Clear();
            _currentStep = 1;
            _currentSelection = new HashSet<ElementId>();

            InjectOptionsBar();
            DebugLogger.Log("[MULTI-SESSION] Session activated");
        }

        public void Deactivate()
        {
            RestoreOptionsBar();
            DebugLogger.Log("[MULTI-SESSION] Session deactivated");
        }

        // ----------------------------------------------------------------
        // Options Bar injection (identical pattern to ClashResolveSession)
        // ----------------------------------------------------------------
        private void InjectOptionsBar()
        {
            try
            {
                var uiDispatcher = System.Windows.Application.Current?.Dispatcher
                    ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;

                uiDispatcher.Invoke(() =>
                {
                    var moduleAsm = DynamicModuleLoader.GetModuleAssembly("clash_resolve");
                    if (moduleAsm == null)
                    {
                        DebugLogger.Log("[MULTI-SESSION] ClashResolve module assembly not found");
                        return;
                    }

                    var ctrlType = moduleAsm.GetType("ClashResolve.Module.UI.MultiClashOptionsBar");
                    if (ctrlType == null)
                    {
                        DebugLogger.Log("[MULTI-SESSION] MultiClashOptionsBar type not found");
                        return;
                    }

                    _optionsBarWpfControl = Activator.CreateInstance(ctrlType);

                    // Subscribe to events via reflection
                    SubscribeBarEvent("Step1Done",  nameof(OnStep1Done));
                    SubscribeBarEvent("Step2Done",  nameof(OnStep2Done));
                    SubscribeBarEvent("Step3Done",  nameof(OnStep3Done));
                    SubscribeBarEvent("CancelRequested", nameof(OnCancelRequested));

                    bool injected = false;
                    var (container, setContent, getContent, dialogBarFe) = FindOptionsBarContainer();
                    if (container != null && setContent != null)
                    {
                        _optionsBarControl    = container;
                        _optionsBarSetContent = setContent;
                        _optionsBarGetContent = getContent;
                        _savedOptionsBarContent = getContent?.Invoke();
                        try
                        {
                            setContent(_optionsBarWpfControl);
                            if (dialogBarFe != null)
                            {
                                _savedDialogBarVisibility = dialogBarFe.Visibility;
                                _savedDialogBarHeight     = dialogBarFe.Height;
                                dialogBarFe.Visibility    = System.Windows.Visibility.Visible;
                                if (double.IsNaN(dialogBarFe.Height) || dialogBarFe.Height < 28)
                                    dialogBarFe.Height = 28;
                                _dialogBarFe = dialogBarFe;
                            }
                            injected = true;
                        }
                        catch (Exception setEx)
                        {
                            DebugLogger.Log($"[MULTI-SESSION] setContent threw: {setEx.Message}");
                        }

                        if (injected)
                        {
                            DebugLogger.Log("[MULTI-SESSION] OptionsBar injected successfully");
                            InstallKeyboardHook();
                        }
                    }

                    if (!injected)
                        DebugLogger.Log("[MULTI-SESSION] WARNING: Could not inject OptionsBar");
                });
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MULTI-SESSION] InjectOptionsBar error: {ex.Message}");
            }
        }

        private void SubscribeBarEvent(string eventName, string handlerName)
        {
            try
            {
                var ctrlType = _optionsBarWpfControl?.GetType();
                if (ctrlType == null) return;

                var ev = ctrlType.GetEvent(eventName);
                if (ev == null)
                {
                    DebugLogger.Log($"[MULTI-SESSION] Event '{eventName}' not found on bar");
                    return;
                }

                var method = typeof(MultiClashSession).GetMethod(handlerName,
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                if (method == null)
                {
                    DebugLogger.Log($"[MULTI-SESSION] Handler '{handlerName}' not found in session");
                    return;
                }

                ev.AddEventHandler(_optionsBarWpfControl,
                    Delegate.CreateDelegate(ev.EventHandlerType, this, method));
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MULTI-SESSION] SubscribeBarEvent({eventName}) error: {ex.Message}");
            }
        }

        private void RestoreOptionsBarCore()
        {
            if (_optionsBarSetContent == null) return;
            _optionsBarSetContent(_savedOptionsBarContent);
            if (_dialogBarFe != null)
            {
                _dialogBarFe.Visibility = _savedDialogBarVisibility;
                _dialogBarFe.Height     = _savedDialogBarHeight;
                _dialogBarFe = null;
            }
            UninstallKeyboardHook();
            DebugLogger.Log("[MULTI-SESSION] OptionsBar restored");
            _optionsBarSetContent = null;
        }

        private void RestoreOptionsBar()
        {
            try
            {
                if (_optionsBarSetContent == null) return;
                var uiDispatcher = System.Windows.Application.Current?.Dispatcher
                    ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;

                if (uiDispatcher.CheckAccess())
                    RestoreOptionsBarCore();
                else
                    uiDispatcher.Invoke(() => RestoreOptionsBarCore());
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MULTI-SESSION] RestoreOptionsBar error: {ex.Message}");
            }
        }

        // ----------------------------------------------------------------
        // OptionsBar container finder (copy from ClashResolveSession)
        // ----------------------------------------------------------------
        private static System.Windows.Window GetRevitMainWindow()
        {
            var wpfMain = System.Windows.Application.Current?.MainWindow;
            if (wpfMain != null) return wpfMain;
            try
            {
                var handle = Process.GetCurrentProcess().MainWindowHandle;
                if (handle != IntPtr.Zero)
                {
                    var src = HwndSource.FromHwnd(handle);
                    if (src?.RootVisual is System.Windows.Window w) return w;
                }
            }
            catch { }
            if (System.Windows.Application.Current != null)
            {
                foreach (System.Windows.Window w in System.Windows.Application.Current.Windows)
                    if (w.IsVisible) return w;
            }
            return null;
        }

        private static (System.Windows.DependencyObject container, Action<object> setContent,
            Func<object> getContent, System.Windows.FrameworkElement dialogBarFe)
            FindOptionsBarContainer()
        {
            try
            {
                var mainWin = GetRevitMainWindow();
                if (mainWin == null) return (null, null, null, null);

                var allElements = FindAllVisualChildren<System.Windows.FrameworkElement>(mainWin);

                var allDialogBars = new List<(System.Windows.FrameworkElement el, double y)>();
                foreach (var el in allElements)
                {
                    if (el.GetType().Name == "DialogBarControl")
                    {
                        try
                        {
                            var pt = el.TranslatePoint(new System.Windows.Point(0, 0), mainWin);
                            allDialogBars.Add((el, pt.Y));
                        }
                        catch
                        {
                            allDialogBars.Add((el, double.MaxValue));
                        }
                    }
                }

                if (allDialogBars.Count == 0) return (null, null, null, null);

                allDialogBars.Sort((a, b) => a.y.CompareTo(b.y));
                var dialogBar     = allDialogBars[0].el;
                var dialogBarType = dialogBar.GetType();

                var contentProp = dialogBarType.GetProperty("Content",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (contentProp != null && contentProp.CanWrite)
                {
                    return (dialogBar,
                        val => contentProp.SetValue(dialogBar, val),
                        () => contentProp.GetValue(dialogBar),
                        dialogBar);
                }

                var setter = TryGetContentSetter(dialogBar);
                if (setter.setContent != null)
                    return (dialogBar, setter.setContent, setter.getContent, dialogBar);

                return (null, null, null, null);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MULTI-SESSION] FindOptionsBarContainer error: {ex.Message}");
                return (null, null, null, null);
            }
        }

        private static (Action<object> setContent, Func<object> getContent) TryGetContentSetter(
            System.Windows.DependencyObject obj)
        {
            var type = obj.GetType();
            foreach (var propName in new[] { "Content", "Child", "DataContext" })
            {
                var prop = type.GetProperty(propName,
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (prop != null && prop.CanWrite)
                    return (val => prop.SetValue(obj, val), () => prop.GetValue(obj));
            }
            var itemsProp = type.GetProperty("Items",
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            if (itemsProp != null)
            {
                var items = itemsProp.GetValue(obj) as System.Collections.IList;
                if (items != null)
                    return (
                        val => { items.Clear(); if (val != null) items.Add(val); },
                        () => items.Count > 0 ? items[0] : null);
            }
            return (null, null);
        }

        private static List<T> FindAllVisualChildren<T>(System.Windows.DependencyObject parent)
            where T : System.Windows.DependencyObject
        {
            var result = new List<T>();
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
                // Only track during steps 1 and 2
                if (_currentStep == 3) return;

                var uidoc = uiApp.ActiveUIDocument;
                if (uidoc == null) return;

                var sel = new HashSet<ElementId>(uidoc.Selection.GetElementIds());
                _currentSelection = sel;

                // Push ids into the bar control so FireStepNDone can read them
                var bar = _optionsBarWpfControl;
                if (bar != null)
                {
                    var setIds = bar.GetType().GetMethod("SetCurrentIds",
                        System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                    setIds?.Invoke(bar, new object[] { sel.ToList() });

                    var updateCount = bar.GetType().GetMethod("UpdateSelectionCount",
                        System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                    updateCount?.Invoke(bar, new object[] { sel.Count });
                }

                DebugLogger.Log($"[MULTI-SESSION] Step {_currentStep} selection: {sel.Count} element(s)");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MULTI-SESSION] OnSelectionChanged error: {ex.Message}");
            }
        }

        // ----------------------------------------------------------------
        // Step event handlers
        // ----------------------------------------------------------------
        private void OnStep1Done(List<ElementId> pipeAIds)
        {
            if (pipeAIds == null || pipeAIds.Count == 0)
            {
                MessageBox.Show("Выберите хотя бы одну трубу A.",
                    "Ошибка выбора", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _pipeAIds = new List<ElementId>(pipeAIds);
            // Advance step tracker BEFORE clearing selection (prevents race in OnSelectionChanged)
            _currentStep = 2;
            DebugLogger.Log($"[MULTI-SESSION] Step1Done: {_pipeAIds.Count} pipe(s) A — advancing to step 2");

            // AdvanceToStep is already on the UI thread (called from button click chain)
            // Call directly without Dispatcher.Invoke to avoid nested re-entrant deadlock
            var advance = _optionsBarWpfControl?.GetType()
                .GetMethod("AdvanceToStep",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            advance?.Invoke(_optionsBarWpfControl, new object[] { 2 });

            TryClearRevitSelection();
        }

        private void OnStep2Done(List<ElementId> pipeBIds)
        {
            _pipeBIds = new List<ElementId>(pipeBIds);
            // Advance step tracker BEFORE clearing selection
            _currentStep = 3;
            DebugLogger.Log($"[MULTI-SESSION] Step2Done: {_pipeBIds.Count} pipe(s) B — advancing to step 3");

            var advance = _optionsBarWpfControl?.GetType()
                .GetMethod("AdvanceToStep",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            advance?.Invoke(_optionsBarWpfControl, new object[] { 3 });

            TryClearRevitSelection();
        }

        private void OnStep3Done(object paramsObj)
        {
            // Read params via reflection (paramsObj is MultiClashOptionsBarParams from ClashResolve.Module)
            var pt = paramsObj.GetType();
            double angleDeg    = (double)pt.GetProperty("AngleDegrees").GetValue(paramsObj);
            double clearanceMm = (double)pt.GetProperty("ClearanceMm").GetValue(paramsObj);
            bool autoClear     = (bool)pt.GetProperty("AutoClearance").GetValue(paramsObj);
            double halfLength  = (double)pt.GetProperty("HalfLengthMm").GetValue(paramsObj);
            bool autoHalf      = (bool)pt.GetProperty("AutoHalfLength").GetValue(paramsObj);
            bool bypassUp      = (bool)pt.GetProperty("BypassUp").GetValue(paramsObj);

            var pipeAIds = new List<ElementId>(_pipeAIds);
            var pipeBIds = new List<ElementId>(_pipeBIds);

            _execHandler.PendingAction = (app) =>
            {
                var doc = app.ActiveUIDocument?.Document;
                if (doc == null) return;

                var moduleAsm = DynamicModuleLoader.GetModuleAssembly("clash_resolve");
                if (moduleAsm == null)
                {
                    DebugLogger.Log("[MULTI-SESSION] ClashResolve module not loaded");
                    return;
                }

                var resolverType    = moduleAsm.GetType("ClashResolve.Module.Core.ClashResolver");
                var multiPairType   = moduleAsm.GetType("ClashResolve.Module.Core.ClashMultiPair");

                int successCount = 0;
                int skipCount    = 0;
                int failCount    = 0;
                var messages     = new System.Text.StringBuilder();

                foreach (var pipeAId in pipeAIds)
                {
                    try
                    {
                        var resolver  = Activator.CreateInstance(resolverType);
                        var multiPair = Activator.CreateInstance(multiPairType);

                        multiPairType.GetProperty("PipeAId").SetValue(multiPair, pipeAId);
                        multiPairType.GetProperty("PipeBIds").SetValue(multiPair, pipeBIds);
                        multiPairType.GetProperty("ClearanceMm").SetValue(multiPair, clearanceMm);
                        multiPairType.GetProperty("HalfLengthMm").SetValue(multiPair, halfLength);
                        multiPairType.GetProperty("AutoClearance").SetValue(multiPair, autoClear);
                        multiPairType.GetProperty("AutoHalfLength").SetValue(multiPair, autoHalf);
                        multiPairType.GetProperty("AngleDegrees").SetValue(multiPair, angleDeg);
                        multiPairType.GetProperty("BypassUp").SetValue(multiPair, bypassUp);

                        var resolveMethod = resolverType.GetMethod("ResolveClashMultiB");
                        var result = resolveMethod.Invoke(resolver, new object[] { doc, multiPair });

                        bool   success = (bool)result.GetType().GetProperty("Success").GetValue(result);
                        string msg     = (string)result.GetType().GetProperty("Message").GetValue(result);

                        DebugLogger.Log($"[MULTI-SESSION] Pipe A ID={pipeAId.Value}: Success={success} Msg={msg}");

                        if (success)
                            successCount++;
                        else if (string.IsNullOrEmpty(msg))
                            skipCount++; // silent skip — no intersections
                        else
                        {
                            failCount++;
                            messages.AppendLine($"• ID {pipeAId.Value}: {msg}");
                        }
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        messages.AppendLine($"• ID {pipeAId.Value}: {ex.Message}");
                        DebugLogger.Log($"[MULTI-SESSION] Pipe A ID={pipeAId.Value} exception: {ex.Message}");
                    }
                }

                // Show summary on UI thread
                System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                {
                    string summary = $"Выполнено: {successCount} из {pipeAIds.Count}";
                    if (skipCount > 0)
                        summary += $"\nПропущено (нет пересечений): {skipCount}";
                    if (failCount > 0)
                        summary += $"\n\nОшибки ({failCount}):\n{messages}";

                    MessageBox.Show(summary,
                        failCount == 0 ? "Готово" : "Готово с ошибками",
                        MessageBoxButton.OK,
                        failCount == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
                });
            };

            _execEvent.Raise();
        }

        private void OnCancelRequested()
        {
            DebugLogger.Log("[MULTI-SESSION] Cancel clicked — deactivating");
            try
            {
                RestoreOptionsBar();
            }
            finally
            {
                Current = null;
            }
        }

        // ----------------------------------------------------------------
        // Helpers
        // ----------------------------------------------------------------
        private void TryClearRevitSelection()
        {
            try
            {
                var uidoc = _uiApp?.ActiveUIDocument;
                uidoc?.Selection.SetElementIds(new List<ElementId>());
                _currentSelection = new HashSet<ElementId>();
            }
            catch { }
        }
    }
}
