using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Tracer ribbon session — handles 3-stage workflow:
    /// 1. Select main line
    /// 2. Select risers (multiple)
    /// 3. Configure slope
    /// </summary>
    public class TracerSession
    {
        // ----------------------------------------------------------------
        // Win32 P/Invoke for keyboard hook
        // ----------------------------------------------------------------
        [DllImport("user32.dll")] private static extern short GetKeyState(int nVirtKey);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int pid);
        [DllImport("user32.dll")] [return:MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc fn, IntPtr hMod, uint threadId);
        [DllImport("user32.dll")] [return:MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet=CharSet.Auto)] private static extern IntPtr GetModuleHandle(string name);

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int VK_ESCAPE = 0x1B;
        private const int VK_RETURN = 0x0D;

        private IntPtr _llHook;
        private LowLevelKeyboardProc _llProc;
        private IntPtr _revitHwnd;

        private bool IsRevitFocused(IntPtr revitHwnd)
        {
            if (revitHwnd == IntPtr.Zero) return false;
            IntPtr fg = GetForegroundWindow();
            if (fg == IntPtr.Zero) return false;
            if (fg == revitHwnd || IsChild(revitHwnd, fg)) return true;
            GetWindowThreadProcessId(fg, out int fgPid);
            return fgPid == Process.GetCurrentProcess().Id;
        }

        private void InstallKeyboardHook()
        {
            try
            {
                _revitHwnd = _uiApp?.MainWindowHandle ?? IntPtr.Zero;
                _llProc = LowLevelKeyboardCallback;

                IntPtr hMod = IntPtr.Zero;
                try
                {
                    using (var mod = Process.GetCurrentProcess().MainModule)
                        hMod = GetModuleHandle(mod?.ModuleName);
                }
                catch (Exception modEx)
                {
                    DebugLogger.Log($"[TRACER-SESSION] MainModule access failed: {modEx.Message}");
                }

                _llHook = SetWindowsHookEx(WH_KEYBOARD_LL, _llProc, hMod, 0);

                if (_llHook != IntPtr.Zero)
                    DebugLogger.Log("[TRACER-SESSION] Win32 keyboard hook installed");
                else
                    DebugLogger.Log("[TRACER-SESSION] SetWindowsHookEx failed");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-SESSION] InstallKeyboardHook error: {ex.Message}");
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
                    DebugLogger.Log("[TRACER-SESSION] Keyboard hook removed");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-SESSION] UninstallKeyboardHook error: {ex.Message}");
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
                    DebugLogger.Log("[TRACER-SESSION] Escape pressed — deactivating");
                    var session = Current;
                    Current = null;
                    var dispatcher = Application.Current?.Dispatcher
                        ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;
                    dispatcher.BeginInvoke(new Action(() => session.Deactivate()),
                        System.Windows.Threading.DispatcherPriority.Normal);
                    return (IntPtr)1;
                }
                else if (vk == VK_RETURN)
                {
                    DebugLogger.Log("[TRACER-SESSION] Enter pressed — triggering Done");
                    var ctrl = _optionsBarWpfControl;
                    var dispatcher = Application.Current?.Dispatcher
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
        public static TracerSession Current { get; private set; }

        // ----------------------------------------------------------------
        // Fields
        // ----------------------------------------------------------------
        private UIApplication _uiApp;
        private ExternalEvent _execEvent;
        private TracerExecuteHandler _execHandler;
        private object _optionsBarWpfControl;

        // Selection state
        private ElementId _selectedMainLineId;
        private List<ElementId> _selectedRiserIds = new List<ElementId>();
        private TracerConnectionType _connectionType;
        private double _slopeValue = 2.0;
        private bool _useMainLineSlope = true;

        // Stage: 1 = select main, 2 = select risers, 3 = configure slope
        private int _currentStage = 1;

        public enum TracerConnectionType
        {
            Angle45,
            LShaped,
            Bottom
        }

        // ----------------------------------------------------------------
        // Activation / Deactivation
        // ----------------------------------------------------------------
        public static void Activate(UIApplication uiApp, TracerConnectionType connectionType)
        {
            DeactivateCurrent();

            DebugLogger.Log($"[TRACER-SESSION] Activating for {connectionType}");
            Current = new TracerSession();
            Current._uiApp = uiApp;
            Current._connectionType = connectionType;
            Current._currentStage = 1;
            Current._selectedRiserIds.Clear();
            Current._selectedMainLineId = null;

            // Create ExternalEvent
            Current._execHandler = new TracerExecuteHandler();
            Current._execEvent = ExternalEvent.Create(Current._execHandler);

            // Install keyboard hook
            Current.InstallKeyboardHook();

            // Show OptionsBar and start stage 1
            Current.ShowOptionsBar();
            Current.StartStage1();
        }

        public static void DeactivateCurrent()
        {
            if (Current != null)
            {
                Current.Deactivate();
                Current = null;
            }
        }

        private void Deactivate()
        {
            DebugLogger.Log("[TRACER-SESSION] Deactivating");
            UninstallKeyboardHook();
            HideOptionsBar();
            _uiApp = null;
            Current = null;
        }

        // ----------------------------------------------------------------
        // OptionsBar injection (same pattern as ClashResolveSession)
        // ----------------------------------------------------------------
        private object _savedOptionsBarContent;
        private object _optionsBarControl;
        private Action<object> _optionsBarSetContent;
        private System.Windows.FrameworkElement _dialogBarFe;
        private System.Windows.Visibility _savedDialogBarVisibility;
        private double _savedDialogBarHeight;

        private void ShowOptionsBar()
        {
            try
            {
                System.Windows.Threading.Dispatcher uiDispatcher =
                    System.Windows.Application.Current?.Dispatcher
                    ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;

                uiDispatcher.Invoke(() =>
                {
                    // Load TracerOptionsBar from Tracer.Module assembly
                    var moduleAsm = System.Reflection.Assembly.Load("Tracer.Module");
                    if (moduleAsm == null)
                    {
                        DebugLogger.Log("[TRACER-SESSION] Tracer.Module assembly not found — cannot inject OptionsBar");
                        return;
                    }

                    var ctrlType = moduleAsm.GetType("Tracer.Module.UI.TracerOptionsBar");
                    if (ctrlType == null)
                    {
                        DebugLogger.Log("[TRACER-SESSION] TracerOptionsBar type not found");
                        return;
                    }

                    _optionsBarWpfControl = Activator.CreateInstance(ctrlType);

                    // Subscribe to events
                    var doneEvent = ctrlType.GetEvent("DoneRequested");
                    var cancelEvent = ctrlType.GetEvent("CancelRequested");
                    var slopeChangedEvent = ctrlType.GetEvent("SlopeChanged");
                    var useMainSlopeChangedEvent = ctrlType.GetEvent("UseMainLineSlopeChanged");

                    if (doneEvent != null)
                        doneEvent.AddEventHandler(_optionsBarWpfControl,
                            Delegate.CreateDelegate(doneEvent.EventHandlerType, this,
                                typeof(TracerSession).GetMethod(nameof(OnDoneRequested),
                                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)));
                    if (cancelEvent != null)
                        cancelEvent.AddEventHandler(_optionsBarWpfControl,
                            Delegate.CreateDelegate(cancelEvent.EventHandlerType, this,
                                typeof(TracerSession).GetMethod(nameof(OnCancelRequested),
                                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)));
                    if (slopeChangedEvent != null)
                        slopeChangedEvent.AddEventHandler(_optionsBarWpfControl,
                            Delegate.CreateDelegate(slopeChangedEvent.EventHandlerType, this,
                                typeof(TracerSession).GetMethod(nameof(OnSlopeChanged),
                                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)));
                    if (useMainSlopeChangedEvent != null)
                        useMainSlopeChangedEvent.AddEventHandler(_optionsBarWpfControl,
                            Delegate.CreateDelegate(useMainSlopeChangedEvent.EventHandlerType, this,
                                typeof(TracerSession).GetMethod(nameof(OnUseMainLineSlopeChanged),
                                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)));

                    // Set initial state
                    UpdateOptionsBarState();

                    bool injected = false;

                    // Find the Options Bar container via visual tree
                    var (container, setContent, getContent, dialogBarFe) = FindOptionsBarContainer(_uiApp?.MainWindowHandle ?? IntPtr.Zero);
                    if (container != null && setContent != null)
                    {
                        _optionsBarControl = container;
                        _optionsBarSetContent = setContent;
                        _savedOptionsBarContent = getContent?.Invoke();
                        DebugLogger.Log($"[TRACER-SESSION] Container type: {container.GetType().FullName}");
                        DebugLogger.Log($"[TRACER-SESSION] Saved content type: {_savedOptionsBarContent?.GetType().FullName ?? "null"}");
                        DebugLogger.Log($"[TRACER-SESSION] Our control type: {_optionsBarWpfControl?.GetType().FullName ?? "null"}");
                        try
                        {
                            setContent(_optionsBarWpfControl);
                            // Make the DialogBarControl visible
                            if (dialogBarFe != null)
                            {
                                _savedDialogBarVisibility = dialogBarFe.Visibility;
                                _savedDialogBarHeight = dialogBarFe.Height;
                                dialogBarFe.Visibility = System.Windows.Visibility.Visible;
                                if (double.IsNaN(dialogBarFe.Height) || dialogBarFe.Height < 28)
                                    dialogBarFe.Height = 28;
                                _dialogBarFe = dialogBarFe;
                                DebugLogger.Log($"[TRACER-SESSION] DialogBarControl made visible, H={dialogBarFe.Height}");
                            }
                            var verifyContent = getContent?.Invoke();
                            DebugLogger.Log($"[TRACER-SESSION] After inject — content is: {verifyContent?.GetType().FullName ?? "null"}");
                            injected = true;
                        }
                        catch (Exception setEx)
                        {
                            DebugLogger.Log($"[TRACER-SESSION] setContent threw: {setEx.GetType().Name}: {setEx.Message}");
                        }
                        if (injected)
                        {
                            DebugLogger.Log("[TRACER-SESSION] OptionsBar injected successfully");
                        }
                    }
                    else
                    {
                        DebugLogger.Log("[TRACER-SESSION] OptionsBar container not found — cannot inject");
                    }
                });
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-SESSION] ShowOptionsBar error: {ex.Message}");
            }
        }

        private void HideOptionsBar()
        {
            try
            {
                if (_optionsBarSetContent == null) return;

                var uiDispatcher =
                    System.Windows.Application.Current?.Dispatcher
                    ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;

                if (uiDispatcher.CheckAccess())
                {
                    RestoreOptionsBarCore();
                }
                else
                {
                    uiDispatcher.Invoke(() => RestoreOptionsBarCore());
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-SESSION] HideOptionsBar error: {ex.Message}");
            }
        }

        private void RestoreOptionsBarCore()
        {
            try
            {
                if (_optionsBarSetContent != null)
                {
                    _optionsBarSetContent(_savedOptionsBarContent);
                }
                if (_dialogBarFe != null)
                {
                    _dialogBarFe.Visibility = _savedDialogBarVisibility;
                    _dialogBarFe.Height = _savedDialogBarHeight;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-SESSION] RestoreOptionsBarCore error: {ex.Message}");
            }
            _optionsBarWpfControl = null;
            _optionsBarSetContent = null;
            DebugLogger.Log("[TRACER-SESSION] OptionsBar restored");
        }

        // ----------------------------------------------------------------
        // OptionsBar container finder (copied from ClashResolveSession)
        // ----------------------------------------------------------------
        private static System.Windows.Window GetRevitMainWindow(IntPtr preferredHwnd = default)
        {
            if (preferredHwnd != IntPtr.Zero)
            {
                try
                {
                    var src = System.Windows.Interop.HwndSource.FromHwnd(preferredHwnd);
                    if (src?.RootVisual is System.Windows.Window wHwnd)
                    {
                        DebugLogger.Log($"[TRACER-SESSION] GetRevitMainWindow: found via UIApp HWND, type={wHwnd.GetType().Name}");
                        return wHwnd;
                    }
                }
                catch { }
            }

            var wpfMain = System.Windows.Application.Current?.MainWindow;
            if (wpfMain != null)
            {
                var children = FindAllVisualChildren<System.Windows.FrameworkElement>(wpfMain);
                DebugLogger.Log($"[TRACER-SESSION] GetRevitMainWindow: Application.MainWindow type={wpfMain.GetType().Name} children={children.Count}");
                if (children.Count > 50)
                    return wpfMain;
                DebugLogger.Log("[TRACER-SESSION] GetRevitMainWindow: Application.MainWindow has too few children, skipping");
            }

            System.Windows.Window bestWindow = null;
            int bestCount = 0;
            if (System.Windows.Application.Current != null)
            {
                foreach (System.Windows.Window w in System.Windows.Application.Current.Windows)
                {
                    if (!w.IsVisible) continue;
                    var cnt = FindAllVisualChildren<System.Windows.FrameworkElement>(w).Count;
                    DebugLogger.Log($"[TRACER-SESSION] GetRevitMainWindow: Window {w.GetType().Name} children={cnt}");
                    if (cnt > bestCount) { bestCount = cnt; bestWindow = w; }
                }
            }
            if (bestWindow != null)
            {
                DebugLogger.Log($"[TRACER-SESSION] GetRevitMainWindow: using largest window {bestWindow.GetType().Name} children={bestCount}");
                return bestWindow;
            }

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

            return null;
        }

        private static (System.Windows.DependencyObject container, Action<object> setContent, Func<object> getContent, System.Windows.FrameworkElement dialogBarFe) FindOptionsBarContainer(IntPtr revitHwnd = default)
        {
            try
            {
                var mainWin = GetRevitMainWindow(revitHwnd);
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
                            DebugLogger.Log($"[TRACER-SESSION] DialogBarControl found: W={el.ActualWidth:F0} H={el.ActualHeight:F0} Y={pt.Y:F0} Visibility={el.Visibility}");
                        }
                        catch
                        {
                            allDialogBars.Add((el, double.MaxValue));
                        }
                    }
                }

                if (allDialogBars.Count == 0)
                {
                    DebugLogger.Log("[TRACER-SESSION] DialogBarControl not found in visual tree");
                    return (null, null, null, null);
                }

                allDialogBars.Sort((a, b) => a.y.CompareTo(b.y));
                System.Windows.FrameworkElement dialogBar = allDialogBars[0].el;
                DebugLogger.Log($"[TRACER-SESSION] Using DialogBarControl at Y={allDialogBars[0].y:F0}");

                var dialogBarType = dialogBar.GetType();
                DebugLogger.Log($"[TRACER-SESSION] DialogBarControl base types: {dialogBarType.BaseType?.FullName}");

                var childrenProp = dialogBarType.GetProperty("Children",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (childrenProp != null)
                {
                    DebugLogger.Log("[TRACER-SESSION] DialogBarControl has Children property (Panel)");
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

                var contentProp = dialogBarType.GetProperty("Content",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (contentProp != null && contentProp.CanWrite)
                {
                    DebugLogger.Log("[TRACER-SESSION] DialogBarControl.Content is writable — using directly");
                    return (
                        dialogBar,
                        val => contentProp.SetValue(dialogBar, val),
                        () => contentProp.GetValue(dialogBar),
                        dialogBar);
                }

                var controlHost = FindVisualChildByTypeName(dialogBar, "ControlHost");
                if (controlHost != null)
                {
                    DebugLogger.Log($"[TRACER-SESSION] Found ControlHost: {controlHost.GetType().FullName}");
                    int n = System.Windows.Media.VisualTreeHelper.GetChildrenCount(controlHost);
                    for (int i = 0; i < n; i++)
                    {
                        var ch = System.Windows.Media.VisualTreeHelper.GetChild(controlHost, i);
                        DebugLogger.Log($"[TRACER-SESSION]   ControlHost child[{i}]: {ch.GetType().Name}");
                    }
                }

                return (null, null, null, null);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-SESSION] FindOptionsBarContainer error: {ex.Message}");
                return (null, null, null, null);
            }
        }

        private static System.Windows.DependencyObject FindVisualChildByTypeName(System.Windows.DependencyObject parent, string typeName)
        {
            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child.GetType().Name == typeName)
                    return child;
                var result = FindVisualChildByTypeName(child, typeName);
                if (result != null) return result;
            }
            return null;
        }

        private static List<T> FindAllVisualChildren<T>(System.Windows.DependencyObject depObj) where T : System.Windows.DependencyObject
        {
            var list = new List<T>();
            if (depObj == null) return list;

            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(depObj);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(depObj, i);
                if (child is T t) list.Add(t);
                list.AddRange(FindAllVisualChildren<T>(child));
            }
            return list;
        }

        private void UpdateOptionsBarState()
        {
            if (_optionsBarWpfControl == null) return;

            try
            {
                var type = _optionsBarWpfControl.GetType();
                var setStageMethod = type.GetMethod("SetStage");
                var setConnectionTypeMethod = type.GetMethod("SetConnectionType");
                var setRiserCountMethod = type.GetMethod("SetRiserCount");
                var setSlopeMethod = type.GetMethod("SetSlope");
                var setMainLineSelectedMethod = type.GetMethod("SetMainLineSelected");

                setStageMethod?.Invoke(_optionsBarWpfControl, new object[] { _currentStage });
                setConnectionTypeMethod?.Invoke(_optionsBarWpfControl, new object[] { (int)_connectionType });
                setRiserCountMethod?.Invoke(_optionsBarWpfControl, new object[] { _selectedRiserIds.Count });
                setSlopeMethod?.Invoke(_optionsBarWpfControl, new object[] { _slopeValue, _useMainLineSlope });
                setMainLineSelectedMethod?.Invoke(_optionsBarWpfControl, new object[] { _selectedMainLineId != null });
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-SESSION] UpdateOptionsBarState error: {ex.Message}");
            }
        }

        // ----------------------------------------------------------------
        // Selection Changed Handler (called from TracerSelectionHandler)
        // ----------------------------------------------------------------
        public void OnSelectionChanged(UIApplication uiApp)
        {
            try
            {
                var doc = uiApp.ActiveUIDocument.Document;
                var selection = uiApp.ActiveUIDocument.Selection;
                var selectedIds = selection.GetElementIds();

                if (_currentStage == 1)
                {
                    // Stage 1: Track single main line selection
                    var pipeId = selectedIds.FirstOrDefault(id => doc.GetElement(id) is Pipe);
                    if (pipeId != _selectedMainLineId)
                    {
                        _selectedMainLineId = pipeId;
                        DebugLogger.Log($"[TRACER-SESSION] Main line selected: {_selectedMainLineId}");
                        
                        // Get real slope from main line
                        if (_selectedMainLineId != null)
                        {
                            _slopeValue = GetMainLineSlope(doc, _selectedMainLineId);
                            DebugLogger.Log($"[TRACER-SESSION] Main line slope: {_slopeValue}%");
                        }
                        
                        UpdateOptionsBarState();
                    }
                }
                else if (_currentStage == 2)
                {
                    // Stage 2: Track multiple riser selections
                    var riserIds = selectedIds.Where(id => doc.GetElement(id) is Pipe).ToList();
                    if (!riserIds.SequenceEqual(_selectedRiserIds))
                    {
                        _selectedRiserIds = riserIds;
                        DebugLogger.Log($"[TRACER-SESSION] Selected {_selectedRiserIds.Count} risers");
                        UpdateOptionsBarState();
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-SESSION] OnSelectionChanged error: {ex.Message}");
            }
        }

        // ----------------------------------------------------------------
        // Slope calculation from main line
        // ----------------------------------------------------------------
        private double GetMainLineSlope(Document doc, ElementId mainLineId)
        {
            try
            {
                var pipe = doc.GetElement(mainLineId) as Pipe;
                if (pipe == null) return 2.0;

                // Get pipe location curve
                var location = pipe.Location as LocationCurve;
                if (location == null) return 2.0;

                var curve = location.Curve;
                var start = curve.GetEndPoint(0);
                var end = curve.GetEndPoint(1);

                // Calculate slope percentage
                double dx = end.X - start.X;
                double dy = end.Y - start.Y;
                double dz = end.Z - start.Z;
                double horizontalLength = Math.Sqrt(dx * dx + dy * dy);
                
                if (horizontalLength < 0.001) // Vertical pipe
                    return 0.0;

                double slope = Math.Abs(dz) / horizontalLength * 100.0;
                return slope;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-SESSION] GetMainLineSlope error: {ex.Message}");
                return 2.0;
            }
        }

        // ----------------------------------------------------------------
        // Stage Handlers
        // ----------------------------------------------------------------
        private void StartStage1()
        {
            DebugLogger.Log("[TRACER-SESSION] Stage 1: Select main line");
            _currentStage = 1;
            _selectedMainLineId = null;
            UpdateOptionsBarState();
            // Selection is tracked via OnSelectionChanged event
        }

        private void StartStage2()
        {
            DebugLogger.Log("[TRACER-SESSION] Stage 2: Select risers");
            _currentStage = 2;
            _selectedRiserIds.Clear();
            UpdateOptionsBarState();
            // Selection is tracked via OnSelectionChanged event
        }

        private void StartStage3()
        {
            DebugLogger.Log("[TRACER-SESSION] Stage 3: Configure slope");
            _currentStage = 3;
            UpdateOptionsBarState();
            // Stage 3 waits for user input on the OptionsBar
        }

        // ----------------------------------------------------------------
        // Event Handlers from OptionsBar
        // ----------------------------------------------------------------
        private void OnDoneRequested()
        {
            DebugLogger.Log($"[TRACER-SESSION] Done requested in stage {_currentStage}");
            switch (_currentStage)
            {
                case 1:
                    if (_selectedMainLineId != null)
                    {
                        StartStage2();
                    }
                    break;
                case 2:
                    if (_selectedRiserIds.Count > 0)
                    {
                        StartStage3();
                    }
                    break;
                case 3:
                    ExecuteConnections();
                    break;
            }
        }

        private void OnCancelRequested()
        {
            DebugLogger.Log("[TRACER-SESSION] Cancel requested");
            Deactivate();
        }

        private void OnSlopeChanged(double slope)
        {
            _slopeValue = slope;
            DebugLogger.Log($"[TRACER-SESSION] Slope changed to {slope}%");
        }

        private void OnUseMainLineSlopeChanged(bool useMainLineSlope)
        {
            _useMainLineSlope = useMainLineSlope;
            DebugLogger.Log($"[TRACER-SESSION] Use main line slope: {useMainLineSlope}");
        }

        // ----------------------------------------------------------------
        // Execution
        // ----------------------------------------------------------------
        private void ExecuteConnections()
        {
            DebugLogger.Log($"[TRACER-SESSION] Executing connections for {_selectedRiserIds.Count} risers");

            if (_selectedMainLineId == null || _selectedRiserIds.Count == 0)
            {
                DebugLogger.Log("[TRACER-SESSION] ERROR: Missing main line or risers");
                return;
            }

            // Set data for the execute handler
            _execHandler.SetData(
                _selectedMainLineId,
                _selectedRiserIds,
                _connectionType,
                _slopeValue,
                _useMainLineSlope);

            _execEvent.Raise();
            Deactivate();
        }

        // ----------------------------------------------------------------
        // Selection Filters - Not used with SelectionChanged pattern
        // ----------------------------------------------------------------
    }

    /// <summary>
    /// ExternalEvent handler for executing Tracer connections
    /// </summary>
    public class TracerExecuteHandler : IExternalEventHandler
    {
        private ElementId _mainLineId;
        private List<ElementId> _riserIds;
        private TracerSession.TracerConnectionType _connectionType;
        private double _slopeValue;
        private bool _useMainLineSlope;

        public void SetData(ElementId mainLineId, List<ElementId> riserIds,
            TracerSession.TracerConnectionType connectionType,
            double slopeValue, bool useMainLineSlope)
        {
            _mainLineId = mainLineId;
            _riserIds = riserIds;
            _connectionType = connectionType;
            _slopeValue = slopeValue;
            _useMainLineSlope = useMainLineSlope;
        }

        public void Execute(UIApplication app)
        {
            DebugLogger.Log($"[TRACER-EXEC] Executing {_riserIds?.Count ?? 0} connections");

            if (_mainLineId == null || _riserIds == null || _riserIds.Count == 0)
            {
                DebugLogger.Log("[TRACER-EXEC] ERROR: Missing data");
                return;
            }

            var doc = app.ActiveUIDocument.Document;

            // Note: Each connection handler manages its own transaction,
            // so we don't wrap them in a transaction here
            try
            {
                foreach (var riserId in _riserIds)
                {
                    // Call the appropriate connection method based on type
                    switch (_connectionType)
                    {
                        case TracerSession.TracerConnectionType.Angle45:
                            Create45DegreeConnection(app, doc, _mainLineId, riserId);
                            break;
                        case TracerSession.TracerConnectionType.LShaped:
                            CreateLShapedConnection(app, doc, _mainLineId, riserId);
                            break;
                        case TracerSession.TracerConnectionType.Bottom:
                            CreateBottomConnection(app, doc, _mainLineId, riserId);
                            break;
                    }
                }

                DebugLogger.Log("[TRACER-EXEC] All connections created successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-EXEC] ERROR: {ex.Message}");
            }
        }

        private void Create45DegreeConnection(UIApplication app, Document doc, ElementId mainLineId, ElementId riserId)
        {
            try
            {
                // Get main line and riser data using Tracer.Module.Core utilities via reflection
                var tracerAssembly = System.Reflection.Assembly.Load("Tracer.Module");
                var utilsType = tracerAssembly?.GetType("Tracer.Module.Core.RevitPipeUtils");
                var calcType = tracerAssembly?.GetType("Tracer.Module.Core.ConnectionCalculator");

                if (utilsType == null || calcType == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not load Tracer.Module types");
                    return;
                }

                // Get MainLineData
                var getMainLineMethod = utilsType.GetMethod("GetMainLineData",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var mainLine = getMainLineMethod?.Invoke(null, new object[] { doc, mainLineId });

                // Get RiserData
                var getRiserMethod = utilsType.GetMethod("GetRiserData",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var riser = getRiserMethod?.Invoke(null, new object[] { doc, riserId });

                if (mainLine == null || riser == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not get main line or riser data");
                    return;
                }

                // Calculate connection points
                var calcMethod = calcType.GetMethod("CalculateConnectionPoints",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var result = calcMethod?.Invoke(null, new object[] { mainLine, riser });

                if (result == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not calculate connection points");
                    return;
                }

                // Extract connectionPoint and endPoint from ValueTuple using reflection
                var resultType = result.GetType();
                var connectionPoint = resultType.GetField("Item1")?.GetValue(result) as XYZ;
                var endPoint = resultType.GetField("Item2")?.GetValue(result) as XYZ;

                if (connectionPoint == null || endPoint == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Connection points are null");
                    return;
                }

                // Get pipe diameter from riser
                var riserDiameter = riser.GetType().GetProperty("Diameter")?.GetValue(riser) as double? ?? 0.1;

                // Determine slope to use
                double slope = _useMainLineSlope
                    ? (mainLine.GetType().GetProperty("Slope")?.GetValue(mainLine) as double? ?? 2.0)
                    : _slopeValue;

                DebugLogger.Log($"[TRACER-EXEC] Creating 45° connection: slope={slope:F2}%, dia={riserDiameter*304.8:F0}mm");

                // Call existing TracerCreateConnectionHandler via reflection
                var assembly = System.Reflection.Assembly.Load("PluginsManager");
                var handlerType = assembly?.GetType("PluginsManager.Commands.TracerCreateConnectionHandler");
                if (handlerType != null)
                {
                    var setDataMethod = handlerType.GetMethod("SetConnectionData",
                        System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    setDataMethod?.Invoke(null, new object[] {
                        mainLineId, riserId,
                        connectionPoint, endPoint, riserDiameter,
                        slope
                    });

                    // Create and execute the handler directly
                    var handler = System.Activator.CreateInstance(handlerType) as IExternalEventHandler;
                    handler?.Execute(app);
                }
                else
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not find TracerCreateConnectionHandler");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-EXEC] ERROR in Create45DegreeConnection: {ex.Message}");
            }
        }

        private void CreateLShapedConnection(UIApplication app, Document doc, ElementId mainLineId, ElementId riserId)
        {
            try
            {
                // Get main line and riser data using Tracer.Module.Core utilities via reflection
                var tracerAssembly = System.Reflection.Assembly.Load("Tracer.Module");
                var utilsType = tracerAssembly?.GetType("Tracer.Module.Core.RevitPipeUtils");
                var calcType = tracerAssembly?.GetType("Tracer.Module.Core.ConnectionCalculator");

                if (utilsType == null || calcType == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not load Tracer.Module types");
                    return;
                }

                // Get MainLineData
                var getMainLineMethod = utilsType.GetMethod("GetMainLineData",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var mainLine = getMainLineMethod?.Invoke(null, new object[] { doc, mainLineId });

                // Get RiserData
                var getRiserMethod = utilsType.GetMethod("GetRiserData",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var riser = getRiserMethod?.Invoke(null, new object[] { doc, riserId });

                if (mainLine == null || riser == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not get main line or riser data");
                    return;
                }

                // Calculate connection points (same as 45° for the starting point)
                var calcMethod = calcType.GetMethod("CalculateConnectionPoints",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var result = calcMethod?.Invoke(null, new object[] { mainLine, riser });

                if (result == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not calculate connection points");
                    return;
                }

                // Extract connectionPoint and endPoint from ValueTuple using reflection
                var resultType = result.GetType();
                var connectionPoint = resultType.GetField("Item1")?.GetValue(result) as XYZ;
                var endPoint = resultType.GetField("Item2")?.GetValue(result) as XYZ;

                if (connectionPoint == null || endPoint == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Connection points are null");
                    return;
                }

                // Get pipe diameter from riser
                var riserDiameter = riser.GetType().GetProperty("Diameter")?.GetValue(riser) as double? ?? 0.1;

                // Determine slope to use
                double slope = _useMainLineSlope
                    ? (mainLine.GetType().GetProperty("Slope")?.GetValue(mainLine) as double? ?? 2.0)
                    : _slopeValue;

                // Get main line start/end points for L-connection
                var mainStartPoint = mainLine.GetType().GetProperty("StartPoint")?.GetValue(mainLine) as XYZ;
                var mainEndPoint = mainLine.GetType().GetProperty("EndPoint")?.GetValue(mainLine) as XYZ;

                DebugLogger.Log($"[TRACER-EXEC] Creating L-shaped connection: slope={slope:F2}%, dia={riserDiameter*304.8:F0}mm");

                // Call existing TracerCreateLConnectionHandler via reflection
                var assembly = System.Reflection.Assembly.Load("PluginsManager");
                var handlerType = assembly?.GetType("PluginsManager.Commands.TracerCreateLConnectionHandler");
                if (handlerType != null)
                {
                    var setDataMethod = handlerType.GetMethod("SetConnectionData",
                        System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    setDataMethod?.Invoke(null, new object[] {
                        mainLineId, riserId,
                        connectionPoint, endPoint, riserDiameter,
                        slope, mainStartPoint, mainEndPoint
                    });

                    // Create and execute the handler directly
                    var handler = System.Activator.CreateInstance(handlerType) as IExternalEventHandler;
                    handler?.Execute(app);
                }
                else
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not find TracerCreateLConnectionHandler");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-EXEC] ERROR in CreateLShapedConnection: {ex.Message}");
            }
        }

        private void CreateBottomConnection(UIApplication app, Document doc, ElementId mainLineId, ElementId riserId)
        {
            try
            {
                // Get main line and riser data using Tracer.Module.Core utilities via reflection
                var tracerAssembly = System.Reflection.Assembly.Load("Tracer.Module");
                var utilsType = tracerAssembly?.GetType("Tracer.Module.Core.RevitPipeUtils");
                var calcType = tracerAssembly?.GetType("Tracer.Module.Core.ConnectionCalculator");

                if (utilsType == null || calcType == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not load Tracer.Module types");
                    return;
                }

                // Get MainLineData
                var getMainLineMethod = utilsType.GetMethod("GetMainLineData",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var mainLine = getMainLineMethod?.Invoke(null, new object[] { doc, mainLineId });

                // Get RiserData
                var getRiserMethod = utilsType.GetMethod("GetRiserData",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var riser = getRiserMethod?.Invoke(null, new object[] { doc, riserId });

                if (mainLine == null || riser == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not get main line or riser data");
                    return;
                }

                // Calculate connection points (same as 45° for the starting point)
                var calcMethod = calcType.GetMethod("CalculateConnectionPoints",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var result = calcMethod?.Invoke(null, new object[] { mainLine, riser });

                if (result == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not calculate connection points");
                    return;
                }

                // Extract connectionPoint and endPoint from ValueTuple using reflection
                var resultType = result.GetType();
                var connectionPoint = resultType.GetField("Item1")?.GetValue(result) as XYZ;
                var endPoint = resultType.GetField("Item2")?.GetValue(result) as XYZ;

                if (connectionPoint == null || endPoint == null)
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Connection points are null");
                    return;
                }

                // Get pipe diameter from riser
                var riserDiameter = riser.GetType().GetProperty("Diameter")?.GetValue(riser) as double? ?? 0.1;

                // Determine slope to use
                double slope = _useMainLineSlope
                    ? (mainLine.GetType().GetProperty("Slope")?.GetValue(mainLine) as double? ?? 2.0)
                    : _slopeValue;

                // Get main line start/end points for bottom connection
                var mainStartPoint = mainLine.GetType().GetProperty("StartPoint")?.GetValue(mainLine) as XYZ;
                var mainEndPoint = mainLine.GetType().GetProperty("EndPoint")?.GetValue(mainLine) as XYZ;

                DebugLogger.Log($"[TRACER-EXEC] Creating bottom connection: slope={slope:F2}%, dia={riserDiameter*304.8:F0}mm");

                // Call existing TracerCreateBottomConnectionHandler via reflection
                var assembly = System.Reflection.Assembly.Load("PluginsManager");
                var handlerType = assembly?.GetType("PluginsManager.Commands.TracerCreateBottomConnectionHandler");
                if (handlerType != null)
                {
                    var setDataMethod = handlerType.GetMethod("SetConnectionData",
                        System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    setDataMethod?.Invoke(null, new object[] {
                        mainLineId, riserId,
                        connectionPoint, endPoint, riserDiameter,
                        slope, mainStartPoint, mainEndPoint
                    });

                    // Create and execute the handler directly
                    var handler = System.Activator.CreateInstance(handlerType) as IExternalEventHandler;
                    handler?.Execute(app);
                }
                else
                {
                    DebugLogger.Log("[TRACER-EXEC] ERROR: Could not find TracerCreateBottomConnectionHandler");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-EXEC] ERROR in CreateBottomConnection: {ex.Message}");
            }
        }

        public string GetName()
        {
            return "Tracer Execute Handler";
        }
    }
}
