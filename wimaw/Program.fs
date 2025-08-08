open System
open System.Runtime.InteropServices
open System.Windows
open System.Windows.Shapes
open System.Windows.Media
open System.Windows.Forms
open System.Windows.Interop
open System.Windows.Media.Effects
open System.Windows.Controls
open System.IO
open System.Text
open System.Text.RegularExpressions
open System.Reflection
open System.Collections.Generic
open System.Threading


type WinEventDelegate = delegate of IntPtr * uint32 * IntPtr * int32 * int32 * uint32 * uint32 -> unit

[<Struct>]
type RECT =
    val Left : int
    val Top : int
    val Right : int
    val Bottom : int

    new ( argLeft : int, argTop : int, argRight : int, argBottom : int ) =
        {
            Left = argLeft;
            Top = argTop;
            Right = argRight;
            Bottom = argBottom;
        }
    member this.Width = this.Right - this.Left
    member this.Height = this.Bottom - this.Top
    member this.Scaling ( ssf : float ) =
        RECT(
            int ( float this.Left / ssf ),
            int ( float this.Top / ssf ),
            int ( float this.Right / ssf ),
            int ( float this.Bottom / ssf )
        )

[<Struct; NoComparison>]
type WINDOWPLACEMENT =
    val length : uint32
    val flags : uint32
    val showCmd : uint32
    val ptMinPosition : System.Drawing.Point
    val ptMaxPosition : System.Drawing.Point
    val rcNormalPosition : RECT

    new ( argLength : uint32 ) =
        {
            length = argLength;
            flags = 0u;
            showCmd = 0u;
            ptMinPosition = System.Drawing.Point()
            ptMaxPosition = System.Drawing.Point()
            rcNormalPosition = RECT()
        }

    static member Init() =
        WINDOWPLACEMENT( uint32 ( Marshal.SizeOf( typeof<WINDOWPLACEMENT> ) ) )


[<Struct; StructLayout(LayoutKind.Sequential)>]
type OSVERSIONINFOEX =
    val dwOSVersionInfoSize : uint32
    val dwMajorVersion : uint32
    val dwMinorVersion : uint32
    val dwBuildNumber : uint32
    val dwPlatformId : uint32
    [<MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)>]
    val szCSDVersion : string
    val wServicePackMajor : uint16
    val wServicePackMinor : uint16
    val wSuiteMask : uint16
    val wProductType : byte
    val wReserved : byte

    new ( argOSVersionInfoSize : uint32 ) =
        {
            dwOSVersionInfoSize = argOSVersionInfoSize;
            dwMajorVersion = 0u;
            dwMinorVersion = 0u;
            dwBuildNumber = 0u;
            dwPlatformId = 0u;
            szCSDVersion = "";
            wServicePackMajor = 0us;
            wServicePackMinor = 0us;
            wSuiteMask = 0us;
            wProductType = 0uy;
            wReserved = 0uy;
        }
    static member Init() =
        OSVERSIONINFOEX( uint32 ( Marshal.SizeOf( typeof<OSVERSIONINFOEX> ) ) )


[<DllImport("user32.dll")>]
extern IntPtr GetForegroundWindow()

[<DllImport("user32.dll")>]
extern bool SetForegroundWindow( IntPtr hWnd )

[<DllImport("user32.dll")>]
extern IntPtr GetWindow( IntPtr hWnd, uint uCmd )

[<DllImport("user32.dll")>]
extern bool GetWindowRect( IntPtr hWnd, RECT& lpRect )

[<DllImport("user32.dll")>]
extern bool IsWindowVisible( IntPtr hWnd )

[<DllImport("user32.dll")>]
extern bool GetWindowPlacement( IntPtr hWnd, WINDOWPLACEMENT& lpwndpl )

[<DllImport("user32.dll")>]
extern IntPtr SetWinEventHook( uint32 eventMin, uint32 eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint32 idProcess, uint32 idThread, uint32 dwFlags )

[<DllImport("user32.dll")>]
extern bool UnhookWinEvent( IntPtr hWinEventHook )

[<DllImport("user32.dll", SetLastError = true)>]
extern bool SetWindowPos( IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags )

[<DllImport("user32.dll")>]
extern int GetWindowLong( IntPtr hwnd, int index )

[<DllImport("user32.dll")>]
extern int SetWindowLong( IntPtr hwnd, int index, int value )

[<DllImport("user32.dll", SetLastError = true)>]
extern uint32 GetWindowThreadProcessId( IntPtr hWnd, uint32& lpdwProcessId )

[<DllImport("user32.dll", SetLastError = true)>]
extern int GetClassName( IntPtr hWnd, StringBuilder lpClassName, int nMaxCount )

[<DllImport("psapi.dll", SetLastError = true)>]
extern uint32 GetModuleFileNameEx( IntPtr hProcess, IntPtr hModule, StringBuilder lpBaseName, int nSize )

[<DllImport("kernel32.dll", SetLastError = true)>]
extern IntPtr OpenProcess( uint32 dwDesiredAccess, bool bInheritHandle, uint32 dwProcessId )

[<DllImport("kernel32.dll", SetLastError = true)>]
extern bool CloseHandle( IntPtr hObject )

[<DllImport("ntdll.dll")>]
extern int RtlGetVersion( OSVERSIONINFOEX& lpVersionInformation )

[<DllImport("user32.dll", SetLastError = true)>]
extern IntPtr GetProcessWindowStation()

[<DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)>]
extern bool GetUserObjectInformationW( IntPtr hObj, int nIndex, IntPtr pvInfo, uint32 nLength, uint32& lpnLengthNeeded )


let EVENT_SYSTEM_FOREGROUND = 0x0003u
let EVENT_OBJECT_LOCATIONCHANGE = 0x800Bu
let EVENT_SYSTEM_MOVESIZEEND = 0x000Bu
let EVENT_SYSTEM_DESKTOPSWITCH = 0x0020u
let EVENT_SYSTEM_SWITCHEND = 0x0015u
let EVENT_SYSTEM_DIALOGEND = 0x0011u
let EVENT_SYSTEM_DIALOGSTART = 0x0010u
let WINEVENT_OUTOFCONTEXT = 0u
let SWP_NOSIZE = 0x0001u
let SWP_NOACTIVATE = 0x0010u
let SWP_NOMOVE = 0x0002u
let GWL_EXSTYLE = -20
let WS_EX_TRANSPARENT = 0x00000020
let WS_EX_NOACTIVATE = 0x08000000
let WS_EX_TOOLWINDOW = 0x00000080
let PROCESS_QUERY_LIMITED_INFORMATION = 0x1400u
let UOI_NAME = 2


type WindowsVersion =
    | Win11_24H2
    | Unknown


/// <summary>
///  Get windows version number.
/// </summary>
/// <returns>
///  Pair of major minor build numbers. If it failed to get version number, it returns None.
/// </returns>
let getWindowsVersion () : WindowsVersion =
    let mutable versionInfo = OSVERSIONINFOEX.Init()
    let status = RtlGetVersion( &versionInfo )
    if status = 0 then
        match ( versionInfo.dwMajorVersion, versionInfo.dwMinorVersion, versionInfo.dwBuildNumber ) with
        | ( 10u, 0u, b ) when b >= 22000u ->
            WindowsVersion.Win11_24H2
        | _ ->
            WindowsVersion.Unknown
    else
        WindowsVersion.Unknown

/// <summary>
///  Get current WindowStation name.
/// </summary>
/// <returns>
///  Obtained WindowStation name, or "" if it failed to get the name.
/// </returns>
let getWindowStationName() : string =
    let mutable needed = 0u
    let hWinsta = GetProcessWindowStation()
    if hWinsta = IntPtr.Zero then
        ""
    else
        GetUserObjectInformationW( hWinsta, UOI_NAME, IntPtr.Zero, 0u, &needed ) |> ignore
        let buf = Marshal.AllocHGlobal( int needed )
        try
            if GetUserObjectInformationW(hWinsta, UOI_NAME, buf, needed, &needed ) then
                Marshal.PtrToStringUni buf
            else
                ""
        finally
            Marshal.FreeHGlobal buf

/// <summary>
///  Check for double startup.
/// </summary>
/// <returns>
///  Returns the allocated mutex. Release it when the program finishes.
/// </returns>
let CheckDuplicate() : Mutex =
    let winsta = getWindowStationName()
    let mutexName = "3fd6f136-c7cb-4f88-9a6b-d909d37944ba_" + winsta
    let mutable createdNew = false
    let mutex = new Mutex( initiallyOwned = true, name = mutexName, createdNew = &createdNew )
    if not createdNew then exit 0
    mutex

/// <summary>
///  Determines whether the specified window is part of the Windows shell, for Windows 11.
/// </summary>
/// <param name="hWnd">
///  Window to be checked.
/// </param>
/// <returns>
///  Returns True if the specified window is a taskbar, start menu, etc.
/// </returns>
let isExplorerShellWindow_Win11 ( hWnd: IntPtr ) : bool =
    let mutable pid = 0u
    GetWindowThreadProcessId( hWnd, &pid ) |> ignore

    let hProcess = OpenProcess( PROCESS_QUERY_LIMITED_INFORMATION, false, pid )
    if hProcess = IntPtr.Zero then false
    else
        let sb = StringBuilder(1024)
        GetModuleFileNameEx( hProcess, IntPtr.Zero, sb, sb.Capacity ) |> ignore
        CloseHandle(hProcess) |> ignore
        let exeName = Path.GetFileName(sb.ToString()).ToLowerInvariant()

        // The windows that make up the Start menu can include "explorer.exe" and "searchhost.exe".
        let isExproer = String.Compare( exeName, "explorer.exe", StringComparison.OrdinalIgnoreCase ) = 0
        let isSearchhost = String.Compare( exeName, "searchhost.exe", StringComparison.OrdinalIgnoreCase ) = 0
        let isStartMenuExperienceHost = String.Compare( exeName, "startmenuexperiencehost.exe", StringComparison.OrdinalIgnoreCase ) = 0
        if ( not isExproer ) && ( not isSearchhost ) && ( not isStartMenuExperienceHost ) then
            false
        else
            // Further specificity by window class name.
            let classNameSb = StringBuilder(256)
            GetClassName( hWnd, classNameSb, classNameSb.Capacity ) |> ignore
            let className = classNameSb.ToString()

            [
                "Shell_TrayWnd"             // Taskbar
                "Shell_SecondaryTrayWnd"    // Taskbar
                "DV2ControlHost"            // Start menu
                "Progman"                   // Desktop
                "Windows.UI.Core.CoreWindow"
                "NotifyIconOverflowWindow"
                "Button"
            ]
            |> List.exists (fun cls -> className = cls)

/// <summary>
///  Determines whether the specified window is part of the Windows shell, for Windows 11.
/// </summary>
/// <param name="winVer">
///  Windows version number.
/// </param>
/// <param name="hWnd">
///  Window to be checked.
/// </param>
/// <returns>
///  Returns True if the specified window is a taskbar, start menu, etc.
/// </returns>
let isExplorerShellWindow ( winVer : WindowsVersion ) ( hWnd: IntPtr ) : bool =
    match winVer with
    | WindowsVersion.Win11_24H2 ->
        isExplorerShellWindow_Win11 hWnd
    | _ ->
        isExplorerShellWindow_Win11 hWnd

/// <summary>
///  Determine whether or not the target should be highlighted.
/// </summary>
/// <param name="winVer">
///  Windows version number.
/// </param>
/// <param name="hWnd">
///  Window to be checked.
/// </param>
/// <returns>
///  If specified window should be highlighted, true is returned.
/// </returns>
let isValidWindow ( winVer : WindowsVersion ) ( hWnd: IntPtr ) ( realScreenSize : Drawing.Rectangle ) ( windowSize : RECT ) : bool =
    if hWnd = IntPtr.Zero then
        false
    elif not (IsWindowVisible hWnd) then
        false
    else
        let mutable placement = WINDOWPLACEMENT.Init()
        if GetWindowPlacement( hWnd, &placement ) |> not then
            false
        else
            let showCmd = placement.showCmd
            if showCmd = 2u || showCmd = 3u then // 2: minimized, 3: maximized
                false
            elif isExplorerShellWindow winVer hWnd then
                false
            else
                if windowSize.Top = 0 && windowSize.Left = 0 && windowSize.Bottom = realScreenSize.Bottom && windowSize.Right = realScreenSize.Right then
                    false
                else
                    true

/// <summary>
///  Configuration class to read settings from config.ini file.
/// </summary>
/// <remarks>
///  The config.ini file should be placed in the same directory as the executable.
///  The file should contain key-value pairs in the format "KEY = VALUE", with optional comments starting with '#'.
///  The keys are case-insensitive and can be used to configure various aspects of the application.
/// </remarks>
type Configuration() =

    let m_ConfPath = 
        System.Reflection.Assembly.GetEntryAssembly()
        |> _.Location
        |> Path.GetDirectoryName
        |> ( fun v -> v + string Path.DirectorySeparatorChar + "config.ini" )

    let m_ConfValues =
        let r = Regex( "^ *([^#= ][^=]*?) *= *([^ ].*) *$" )
        try
            File.ReadAllLines m_ConfPath
        with
        | _ -> Array.Empty()
        |> Array.map r.Match
        |> Array.filter _.Success
        |> Array.map ( fun itr -> itr.Groups.[1].Value.ToUpper(), itr.Groups.[2].Value )
        |> Array.distinctBy fst
        |> Array.map KeyValuePair
        |> Dictionary

    /// Get the value that is specified by "Type" key.
    member _.HighlightType : string =
        let r, v = m_ConfValues.TryGetValue "TYPE"
        if r then v.ToUpper() else "HALO"

    /// Get the value that is specified by "Topmot" key.
    member this.TopMost : bool =
        this.GetValue( "TOPMOST", false, Boolean.TryParse )

    /// Get the value that is specified by "BorderThickness" key.
    member this.BorderThickness : float =
        this.GetValue( "BORDERTHICKNESS", 5.0, Double.TryParse )
        |> max 0.0 |> min 100.0

    /// Get the value that is specified by "Opacity" key.
    member this.Opacity : float =
        this.GetValue( "OPACITY", 0.8, Double.TryParse )
        |> max 0.0 |> min 1.0

    /// Get the value that is specified by "LeftMargin" key.
    member this.LeftMargin : float =
        this.GetValue( "LEFTMARGIN", 0.0, Double.TryParse )
        |> max -100.0 |> min 100.0

    /// Get the value that is specified by "RightMargin" key.
    member this.RightMargin : float =
        this.GetValue( "RIGHTMARGIN", 0.0, Double.TryParse )
        |> max -100.0 |> min 100.0

    /// Get the value that is specified by "TopMargin" key.
    member this.TopMargin : float =
        this.GetValue( "TOPMARGIN", 0.0, Double.TryParse )
        |> max -100.0 |> min 100.0

    /// Get the value that is specified by "BottomMargin" key.
    member this.BottomMargin : float =
        this.GetValue( "BOTTOMMARGIN", 0.0, Double.TryParse )
        |> max -100.0 |> min 100.0

    /// Get the value that is specified by "Color" key.
    member this.Color : Color =
        let parse ( s : string ) : ( bool * Color ) =
            let r = Regex( "^ *([0-9]+) *, *([0-9]+) *, *([0-9]+) *" )
            let m = r.Match s
            if not m.Success then
                false, Color()
            else
                let r1, v1 = Byte.TryParse m.Groups.[1].Value
                let r2, v2 = Byte.TryParse m.Groups.[2].Value
                let r3, v3 = Byte.TryParse m.Groups.[3].Value
                if not( r1 && r2 && r3 ) then
                    false, Color()
                else
                    true, Color.FromRgb( v1, v2, v3 )
        this.GetValue( "COLOR", Color.FromRgb( 255uy, 0uy, 0uy ), parse )

    /// <summary>
    /// Get the value that is specified by the key name from loaded configurations..
    /// </summary>
    /// <param name="keyName">
    ///  Key name to be searched in the configuration file.
    /// </param>
    /// <param name="defVal">
    ///  Default value to be returned if the key is not found or conversion fails.
    /// </param>
    /// <param name="conv">
    ///  Conversion function to convert the string value to the desired type.
    /// </param>
    /// <returns>
    ///  The value corresponding to the key name, or the default value if not found or conversion fails.
    /// </returns>
    member _.GetValue<'T>( keyName : string, defVal : 'T, conv : string -> ( bool * 'T ) ) : 'T =
        let r, v = m_ConfValues.TryGetValue keyName
        if not r then
            defVal
        else
            let r2, v2 = conv v
            if not r2 then
                defVal
            else
                v2


/// <summary>
///  Implementing common highlighting functionality.
/// </summary>
/// <param name="m_WinVer">
///  Windows version number.
/// </param>
[<AbstractClass>]
type OverlayWindow( m_WinVer : WindowsVersion ) =
    inherit Window(
        Left = 0.0,
        Top = 0.0,
        Width = 0.0,
        Height = 0.0,
        WindowStyle = WindowStyle.None,
        AllowsTransparency = true,
        Background = Brushes.Transparent,
        IsHitTestVisible = false,
        ShowInTaskbar = false
    )

    /// <summary>
    ///  Called when the highlight window is loaded.
    /// </summary>
    /// <param name="appHwnd">
    ///  The window handle of the highlight window.
    /// </param>
    member this.OnLoaded ( appHwnd : Lazy<nativeint> ) : unit =
        appHwnd.Force() |> ignore
        let extendedStyle = GetWindowLong( appHwnd.Value, GWL_EXSTYLE )
        SetWindowLong( appHwnd.Value, GWL_EXSTYLE, extendedStyle ||| WS_EX_TRANSPARENT ||| WS_EX_NOACTIVATE ||| WS_EX_TOOLWINDOW ) |> ignore
        let hWnd = GetForegroundWindow()
        this.OnUpdateInternal appHwnd hWnd
        
    /// <summary>
    ///  Update the highlight window position.
    /// </summary>
    /// <param name="appHwnd">
    ///  The window handle of the highlight window.
    /// </param>
    /// <param name="eventHwnd">
    ///  The window handle of the window where the event notification occurred.
    /// </param>
    member this.OnUpdateInternal ( appHwnd : Lazy<nativeint> ) ( eventHwnd : IntPtr ) : unit =
        let hWnd = GetForegroundWindow()

        if not appHwnd.IsValueCreated then
            this.Show()

        elif eventHwnd = appHwnd.Value then
            // ignore
            ()

        elif hWnd = appHwnd.Value then
            // select next window
            let nextWindow = GetWindow( hWnd, 2u )
            SetForegroundWindow( nextWindow ) |> ignore

        else
            let mutable rect = RECT()
            let realScreenSize = Screen.PrimaryScreen.Bounds
            if GetWindowRect( hWnd, &rect ) |> not then
                this.HideIfVisible()
            elif isValidWindow m_WinVer hWnd realScreenSize rect then
                let ssf =
                    let h1 = float realScreenSize.Height
                    let h2 = SystemParameters.PrimaryScreenHeight
                    ( float h1 ) / h2
                let ssfRest = rect.Scaling ssf
                this.UpdatePosition ssfRest hWnd
            else
                this.HideIfVisible()

    /// <summary>
    ///  Hide the highlight window.
    /// </summary>
    member private this.HideIfVisible() : unit =
        if this.IsVisible then this.Hide()

    // Called when the window position should be updated.
    abstract OnUpdate : IntPtr -> unit

    // Implement a function to change the window position depending on the type of highlight.
    abstract UpdatePosition : RECT -> IntPtr -> unit



/// <summary>
///  Display highlight bar above the active window.
/// </summary>
/// <param name="m_WinVer">
///  Windows version number.
/// </param>
type HaloHighlight( m_WinVer : WindowsVersion, config : Configuration ) as this =
    inherit OverlayWindow( m_WinVer )

    let topMost = config.TopMost
    let highlightColor = config.Color
    let borderThickness = config.BorderThickness
    let opacity = config.Opacity
    let leftMargin = config.LeftMargin
    let rightMargin = config.RightMargin
    let topMargin = config.TopMargin
    let appHwnd = lazy( WindowInteropHelper( this ) |> _.Handle )

    do
        let rect = Rectangle(
            Fill = SolidColorBrush( highlightColor ),
            Effect = new BlurEffect( Radius = 5.0 ),
            IsHitTestVisible  = false,
            Margin = Thickness( 5.0 ),
            Opacity = opacity
        )
        this.Topmost <- topMost
        this.Content <- rect
        this.Loaded.AddHandler ( fun _ _ -> this.OnLoaded appHwnd )

    /// <summary>
    ///  Called when the window position should be updated.
    /// </summary>
    /// <param name="eventHwnd">
    ///  The window handle of the window where the event notification occurred.
    /// </param>
    override this.OnUpdate ( eventHwnd : IntPtr ) : unit = 
        this.OnUpdateInternal appHwnd eventHwnd

    /// <summary>
    ///  Implement a function to change the window position depending on the type of highlight.
    /// </summary>
    /// <param name="r">
    ///  Target window display position.
    /// </param>
    /// <param name="targetWindowHwnd">
    ///  The target window where the highlight should be displayed.
    /// </param>
    override this.UpdatePosition ( r : RECT ) ( targetWindowHwnd : IntPtr ) : unit =
        this.Left <- float r.Left - leftMargin
        this.Top <- float r.Top - 10.0 - borderThickness - topMargin
        this.Width <- float r.Width + leftMargin + rightMargin
        this.Height <- 10.0 + borderThickness
        if not this.IsVisible then
            this.Show()
        let whwnd = if topMost then IntPtr 0 else targetWindowHwnd
        SetWindowPos( appHwnd.Value, whwnd, 0, 0, 0, 0, SWP_NOSIZE ||| SWP_NOACTIVATE ||| SWP_NOMOVE ) |> ignore


/// <summary>
///  Display a highlight around the active window.
/// </summary>
/// <param name="m_WinVer">
///  Windows version number.
/// </param>
type BoxHighlight( m_WinVer : WindowsVersion, config : Configuration ) as this =
    inherit OverlayWindow( m_WinVer )

    let topMost = config.TopMost
    let highlightColor = config.Color
    let borderThickness = config.BorderThickness
    let opacity = config.Opacity
    let leftMargin = config.LeftMargin
    let rightMargin = config.RightMargin
    let topMargin = config.TopMargin
    let bottomMargin = config.BottomMargin
    let appHwnd = lazy( WindowInteropHelper( this ) |> _.Handle )

    do
        let rect = Rectangle(
            Stroke = SolidColorBrush( highlightColor ),
            StrokeThickness = borderThickness,
            Effect = new BlurEffect( Radius = 5.0 ),
            IsHitTestVisible  = false,
            Margin = Thickness( 5.0 ),
            Opacity = opacity
        )
        this.Topmost <- topMost
        this.Content <- rect
        this.Loaded.AddHandler ( fun _ _ -> this.OnLoaded appHwnd )

    /// <summary>
    ///  Called when the window position should be updated.
    /// </summary>
    /// <param name="eventHwnd">
    ///  The window handle of the window where the event notification occurred.
    /// </param>
    override this.OnUpdate ( eventHwnd : IntPtr ) : unit = 
        this.OnUpdateInternal appHwnd eventHwnd

    /// <summary>
    ///  Implement a function to change the window position depending on the type of highlight.
    /// </summary>
    /// <param name="r">
    ///  Target window display position.
    /// </param>
    /// <param name="targetWindowHwnd">
    ///  The target window where the highlight should be displayed.
    /// </param>
    override this.UpdatePosition ( r : RECT ) ( targetWindowHwnd : IntPtr ) : unit =
        this.Left <- float r.Left - 5.0 - borderThickness - leftMargin
        this.Top <- float r.Top - 5.0 - borderThickness - topMargin
        this.Width <- float r.Width + 10.0 + borderThickness * 2.0 + leftMargin + rightMargin
        this.Height <- float r.Height + 10.0 + borderThickness * 2.0 + topMargin + bottomMargin
        if not this.IsVisible then
            this.Show()
        let whwnd = if topMost then IntPtr 0 else targetWindowHwnd
        SetWindowPos( appHwnd.Value, whwnd, 0, 0, 0, 0, SWP_NOSIZE ||| SWP_NOACTIVATE ||| SWP_NOMOVE ) |> ignore


[<EntryPoint; STAThread>]
let main _ =
    use mutex = CheckDuplicate()
    let winVer = getWindowsVersion()
    let app = Application()
    let config = Configuration()

    let overlay =
        match config.HighlightType with
        | "BOX" ->
            BoxHighlight( winVer, config ) :> OverlayWindow
        | _ ->
            HaloHighlight( winVer, config ) :> OverlayWindow

    let contextMenu = new ContextMenuStrip()
    let iconImage = Application.GetResourceStream( new Uri( "/wimaw;component/wimaw.ico", UriKind.Relative ) ).Stream
    contextMenu.Items.Add( "Exit", null, fun _ _ -> app.Shutdown() ) |> ignore
    let _ = new NotifyIcon(
        Icon = new System.Drawing.Icon( iconImage ),
        Visible = true,
        Text = "Which is my active window!",
        ContextMenuStrip = contextMenu
    )

    let callback = WinEventDelegate(fun _ _ hwnd _ _ _ _ ->
        app.Dispatcher.Invoke( Action( fun () ->
            overlay.OnUpdate hwnd
        ) )
        |> ignore
    )

    // Hook multiple events
    let hooks =
        [ EVENT_SYSTEM_FOREGROUND;
          EVENT_OBJECT_LOCATIONCHANGE;
          EVENT_SYSTEM_MOVESIZEEND;
          EVENT_SYSTEM_DESKTOPSWITCH;
          EVENT_SYSTEM_SWITCHEND;
          EVENT_SYSTEM_DIALOGEND;
          EVENT_SYSTEM_DIALOGSTART;
        ]
        |> List.map (fun evt -> SetWinEventHook(evt, evt, IntPtr.Zero, callback, 0u, 0u, WINEVENT_OUTOFCONTEXT))

    let exitCode =
        try
            overlay.Hide()
            app.Run()
        finally
            for hook in hooks do
                if hook <> IntPtr.Zero then
                    UnhookWinEvent(hook) |> ignore

    mutex.Dispose()
    GC.KeepAlive callback
    exitCode
