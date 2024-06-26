using EnvDTE;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
namespace wimvimkeys;

static class Program
{


    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private static LowLevelKeyboardProc _proc = HookCallback;
    private static IntPtr _hookID = IntPtr.Zero;
    private static SynchronizationContext synchronizationContext;

    private static bool VisualMode { get; set; } = false;
    private static int? RepeatKey { get; set; }

    private static bool enableVimMode { get; set; } = false;
    private static bool normalMode { get; set; } = true;




    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.

        //synchronizationContext = SynchronizationContext.Current;


        Task.Run(async () => await WindowMonitoringTask());


        ApplicationConfiguration.Initialize();
        //Application.Run(new Form1());

        _hookID = SetHook(_proc);
        Application.Run();
        UnhookWindowsHookEx(_hookID);



    }

    private static IntPtr SetHook(LowLevelKeyboardProc proc)
    {
        using (System.Diagnostics.Process curProcess = System.Diagnostics.Process.GetCurrentProcess())
        using (ProcessModule curModule = curProcess.MainModule)
        {
            return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
              GetModuleHandle(curModule.ModuleName), 0);
        }
    }


    private static bool Yanked { get; set; }
    private static bool DeletionMode{ get; set; }
    private static bool OperationPending { get; set; }
    private static bool ChangeCommand { get; set; }
    private static bool InnerCommand { get; set; }

    static Queue<char> buffer = new Queue<char>(); // To store recently pressed keys

    private static bool movePanesMode = false;
    private static bool isShiftPressed = false;
    private static bool isControlPressed = false;
    private static bool isCapsLockOn = false;

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    private static IntPtr HookCallback(
      int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN && enableVimMode)
        {
            int vkCode = Marshal.ReadInt32(lParam);



            // Check if Shift key is pressed
            if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
            {
                isShiftPressed = true;
            }
            else
            {
                isShiftPressed = false;
            }
            // Check if Shift key is pressed
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                isControlPressed = true;
            }
            else
            {
                isControlPressed = false;
            }

            // Check if Caps Lock is on
            isCapsLockOn = Control.IsKeyLocked(Keys.CapsLock);

            // Convert virtual key code to character considering Shift and Caps Lock
            char pressedChar = Convert.ToChar(vkCode);

            // Determine if the character is uppercase or lowercase based on Shift and Caps Lock
            bool isLowerCase = (!isShiftPressed && !isCapsLockOn) || (isShiftPressed && isCapsLockOn);

            if (isControlPressed && pressedChar == 'K')
            {
                buffer.Enqueue(pressedChar);

                if (buffer.Count > 2)
                {
                    buffer.Dequeue();
                }

            }

            if (isLowerCase && vkCode == 226)
            {
                if (VisualMode)
                {
                    SendKeys.SendWait("+{TAB}");
                    return (IntPtr)1;
                }
                SendKeys.SendWait("{HOME}+{TAB}");
                return (IntPtr)1;
            }
            if (!isLowerCase && vkCode == 226)
            {
                ReleaseShiftKey();
                if (VisualMode)
                {

                    SendKeys.SendWait("{TAB}");
                    return (IntPtr)1;
                }
                SendKeys.SendWait("{HOME}{TAB}");
                return (IntPtr)1;
            }

            if (buffer.Count == 2 && buffer.Peek() == 'K' && pressedChar == 'D')
            {
                buffer.Clear();

                SendKeys.Send("^c");    // Trigger the copy action
                System.Threading.Thread.Sleep(1000); // wait a bit for the caret to move

                var code = Clipboard.GetText();

                code = code.Replace("\t", " ");
                code = Regex.Replace(code, " {2,}", " ").Replace("// [", "//[");

                var parts = code.Split("FUNCTION_BLOCK");
                var headerSection = parts[0] + "FUNCTION_BLOCK";

                if (parts.Length > 1)
                {
                    var headerlines = Regex.Split(headerSection, @"\r\n|\r|\n").ToList();
                    headerlines = headerlines.Align("[");

                    var codeSection = parts[1];
                    var lines = Regex.Split(codeSection, @"\r\n|\r|\n").ToList();
                    lines = lines.AlignFirstOccurance(":");
                    lines = lines.AlignFirstOccurance(";");
                    lines = lines.AlignFirstOccurance("//");
                    lines = lines.Align("[",
                        s =>
                        {
                            if (s.Replace(" ", "").Contains("ARRAY["))
                            {
                                return false;
                            }
                            return true;
                        });

                    Clipboard.SetText(string.Join(Environment.NewLine, headerlines) + string.Join(Environment.NewLine, lines));

                    System.Threading.Thread.Sleep(1000); // wait a bit for the caret to move


                    SendKeys.Send("^v");    // Trigger the copy action
                }

            }


            if (isControlPressed && (Keys)vkCode == Keys.D)
                return (IntPtr)1;
            if (isControlPressed && (Keys)vkCode == Keys.K)
                return (IntPtr)1;

            if (normalMode)
            {
                if (ChangeCommand)
                {
                    if (InnerCommand)
                    {

                        if ((Keys)vkCode == Keys.OemOpenBrackets)
                        {
                            //InnerCommand = false;
                            //ChangeCommand = false;
                            int? startOpenBractedPosition = null;
                            int? endOpenBractedPosition = null;
                            int backtrackpositions = 0;
                            int forwardrackpositions = 0;
                            while (startOpenBractedPosition == null)
                            {

                                SendKeys.Send("{LEFT}");
                                SendKeys.Send("^(c)");
                                backtrackpositions++;
                                string copiedText = Clipboard.GetText();
                                if (copiedText == "[")
                                {
                                    startOpenBractedPosition = backtrackpositions;
                                    for (int i = 0; i < backtrackpositions; i++)
                                    {
                                        SendKeys.Send("{RIGHT}");

                                    }
                                }
                            }
                            while (endOpenBractedPosition == null)
                            {

                                SendKeys.Send("{RIGHT}");
                                SendKeys.Send("^(c)");
                                backtrackpositions++;
                                string copiedText = Clipboard.GetText();
                                if (copiedText == "]")
                                {
                                    endOpenBractedPosition = forwardrackpositions;
                                    //for (int i = 0; i < forwardrackpositions; i++)
                                    //{
                                    //    SendKeys.Send("{LEFT}");

                                    //}
                                }
                            }


                            return (IntPtr)1;
                        }

                    }

                    if ((Keys)vkCode == Keys.I)
                    {
                        InnerCommand = true;
                        return (IntPtr)1;
                    }
                }

                //if (isUpperCase && pressedChar == 'Y')
                //{
                //buffer.Enqueue(pressedChar);

                //// Limit the buffer size to 2 to check for 'yy'
                //if (buffer.Count > 2)
                //    buffer.Dequeue();

                //if (buffer.Count == 2 && buffer.Peek() == 'Y' && pressedChar == 'Y')
                //{
                //    buffer.Clear();

                //    SendKeys.SendWait("{HOME}");
                //    System.Threading.Thread.Sleep(100); // wait a bit for the caret to move

                //    // Simulate pressing Shift+End to select the entire line
                //    SendKeys.SendWait("+{END}");
                //    System.Threading.Thread.Sleep(100); // wait a bit for the line to be selected

                //    SendKeys.Send("^(c)");    // Trigger the copy action
                //}

                //}


                if (isControlPressed && pressedChar == 'W')
                {
                    movePanesMode = true;
                    return (IntPtr)1;
                }

                if (!isControlPressed && (Keys)vkCode == Keys.O)
                {

                    if (isLowerCase)
                    {
                        SendKeys.SendWait("{END}{ENTER}");
                        normalMode = false;
                        return (IntPtr)1;
                    }
                    else
                    {
                        ReleaseShiftKey();
                        SendKeys.SendWait("{HOME}{HOME}{UP}{END}{ENTER}");
                        normalMode = false;
                        return (IntPtr)1;
                    }
                }


                if ((Keys)vkCode == Keys.Y)
                {
                    if (OperationPending && !VisualMode)
                    {
                        SendKeys.SendWait("{HOME}");
                        SendKeys.SendWait("+{END}");
                        SendKeys.Send("^c");
                        SendKeys.Send("{END}");

                        OperationPending = false;
                        return (IntPtr)1;
                    }
                    if (OperationPending && VisualMode)
                    {
                        SendKeys.Send("^c");
                        OperationPending = false;
                        VisualMode = false;
                        return (IntPtr)1;
                    }

                    if (VisualMode)
                    {
                        SendKeys.Send("^c");
                        VisualMode = false;
                        return (IntPtr)1;
                    }

                    OperationPending = true;

                    return (IntPtr)1;
                }

                if ((Keys)vkCode == Keys.P)
                {
                    if (isLowerCase)
                    {
                        SendKeys.Send("{END}{ENTER}");
                        SendKeys.Send("^v");
                        SendKeys.Send("{HOME}{RIGHT}{LEFT}");
                        return (IntPtr)1;
                    }
                    else
                    {
                        ReleaseShiftKey();
                        SendKeys.Send("{HOME}{HOME}{UP}{END}");
                        SendKeys.Send("{ENTER}");
                        SendKeys.Send("^v");
                        SendKeys.Send("{HOME}{RIGHT}{LEFT}");
                        return (IntPtr)1;
                    }

                }


                if (IsNumericKey(vkCode))
                {
                    Console.WriteLine($"Numeric key pressed: {(Keys)vkCode}");
                    if ((Keys)vkCode != Keys.D0)
                    {
                        RepeatKey = vkCode - 48;
                        return (IntPtr)1; // Signal that we handled the key press
                    }
                }




                if ((Keys)vkCode == Keys.J)
                {
                    if (movePanesMode)
                    {
                        SendKeys.SendWait("{F6}");
                        movePanesMode = false;
                        return (IntPtr)1;
                    }
                    else
                    {

                        return CommandChaining("{DOWN}");
                    }


                }
                if ((Keys)vkCode == Keys.K)
                {

                    if (movePanesMode)
                    {
                        SendKeys.SendWait("{F6}");
                        movePanesMode = false;

                        return (IntPtr)1;
                    }
                    else
                    {
                        return CommandChaining("{UP}");
                    }

                }

                // Advanced
                if (!isControlPressed && (Keys)vkCode == Keys.C) { ChangeCommand = true; return (IntPtr)1; }

                if (isControlPressed && isLowerCase && (Keys)vkCode == Keys.O)
                {
                    SendKeys.Send("^-");
                    return (IntPtr)1;
                }
                if (isControlPressed && isLowerCase && (Keys)vkCode == Keys.I)
                {
                    SendKeys.Send("^+-");
                    ReleaseShiftKey();
                    return (IntPtr)1;
                }

                // General 

                // TODO: delete current line
                if ((Keys)vkCode == Keys.D)
                {
                    if (OperationPending && !isShiftPressed)
                    {
                        SendKeys.Send("{F12}");
                        OperationPending = false;
                        return (IntPtr)1;
                    }
                    if (DeletionMode && !VisualMode)
                    {
                        SendKeys.Send("{HOME}{HOME}+{END}^c");
                        var curClip = Clipboard.GetText();
                        if (curClip == "\r\n")
                            SendKeys.Send("{HOME}{HOME}+{END}{DEL}");
                        else
                        {
                            SendKeys.Send("{HOME}{HOME}+{END}{DEL}{DEL}");
                        }
                        DeletionMode = false;
                        return (IntPtr)1;
                    }
                    if (DeletionMode && VisualMode)
                    {
                        SendKeys.Send("^c{DEL}");
                        DeletionMode = false;
                        VisualMode = false;
                        return (IntPtr)1;
                    }
                    if (VisualMode)
                    {
                        SendKeys.Send("^c{DEL}");
                        VisualMode = false;
                        return (IntPtr)1;
                    }
                    DeletionMode = true;
                    return (IntPtr)1;
                }

                //indexer visula mode
                //ab - a block with ()
                //aB - a block with {}
                //at - a block with <> tags
                //ib - inner block with ()
                //iB - inner block with {}
                //it - inner block with <> tags





                // End of line
                //if ((Keys)vkCode == Keys.D0) { SendKeys.Send("{END}"); return (IntPtr)1; }


                //TODO:Start of line
                //if (vkCode == 0x24)
                //{
                //    if (!OperationPending)
                //    {
                //        SendKeys.Send("{HOME}");
                //        return (IntPtr)1;
                //    }
                //}

                // move Left
                if ((Keys)vkCode == Keys.H) { return CommandChaining("{LEFT}"); }

                // move Right
                if ((Keys)vkCode == Keys.L) { return CommandChaining("{RIGHT}"); }

                // move Back one word
                if ((Keys)vkCode == Keys.B) { 

                    SendKeys.Send("^{LEFT}"); 
                    SendKeys.Send("+{LEFT}");
                    SendKeys.Send("^c");
                    var clip = Clipboard.GetText();
                    ReleaseShiftKey();
                    if(clip == " ")
                    {
                        SendKeys.Send("^{LEFT}");
                    }else
                    {
                        SendKeys.Send("{RIGHT}{LEFT}");

                    }
                    return (IntPtr)1;
                    //return CommandChaining("^{LEFT}"); 
                }

                // Delete under cursor
                if ((Keys)vkCode == Keys.X) { return CommandChaining("{DEL}"); }

                //  Move to top
                if (isLowerCase && (Keys)vkCode == Keys.G)
                {
                    if (OperationPending)
                    {
                        OperationPending = false;
                        return CommandChaining("^{HOME}");
                    }
                    OperationPending = true;
                    return (IntPtr)1;
                }

                // Move to bottom
                if (isShiftPressed && (Keys)vkCode == Keys.G) { return CommandChaining("^{END}"); }

                // Undo 
                if ((Keys)vkCode == Keys.U) { return CommandChaining("^z"); }

                // TODO: Redo
                if (isControlPressed && (Keys)vkCode == Keys.R) { 


                    return CommandChaining("^Y"); 
                }
                if (isLowerCase && (Keys)vkCode == Keys.R)
                {
                    if (OperationPending)
                    {
                        ReleaseShiftKey();
                        SendKeys.Send("+{F12}");
                        OperationPending = false;
                        return (IntPtr)1;
                    }
                }

                // Change between Decleration View and Code Pane
                if (!isControlPressed && (Keys)vkCode == Keys.W)
                {

                    if (DeletionMode)
                    {
                        SendKeys.Send("+^{RIGHT}{DEL}");
                        DeletionMode = false;
                        return (IntPtr)1;
                    }

                    SendKeys.Send("^{RIGHT}");
                    SendKeys.Send("+{RIGHT}");
                    SendKeys.Send("^c");
                    var clip = Clipboard.GetText();
                    ReleaseShiftKey();
                    if(clip == " ")
                    {
                        SendKeys.Send("^{RIGHT}");
                    }else
                    {
                        SendKeys.Send("{LEFT}{RIGHT}");

                    }
                    return (IntPtr)1;

                    //return CommandChaining("^{RIGHT}");
                }

                // Change between Decleration View and Code Pane
                if (!isControlPressed && (Keys)vkCode == Keys.E) { return CommandChaining("^{RIGHT}"); }

                // TODO: Change tav using Ctrl+w + H or L
                if (isLowerCase && (Keys)vkCode == Keys.I) { normalMode = false; return (IntPtr)1; }

                if (!isLowerCase && (Keys)vkCode == Keys.I)
                {
                    ReleaseShiftKey();
                    SendKeys.Send("{HOME}");
                    normalMode = false;
                    return (IntPtr)1;
                }
                if (!isLowerCase && (Keys)vkCode == Keys.A)
                {
                    ReleaseShiftKey();
                    SendKeys.Send("{END}");
                    normalMode = false;
                    return (IntPtr)1;
                }

                // Enable visual mode
                if (!isControlPressed && isLowerCase && (Keys)vkCode == Keys.V)
                {
                    VisualMode = !VisualMode;
                    return (IntPtr)1;
                }
                if (!isControlPressed && !isLowerCase && (Keys)vkCode == Keys.V)
                {

                    ReleaseShiftKey();
                    SendKeys.Send("{HOME}{HOME}");
                    VisualMode = !VisualMode;
                    SendKeys.Send("+{END}");
                    return (IntPtr)1;
                }
            }


            if ((Keys)vkCode == Keys.Escape)
            {
                normalMode = true;
                VisualMode = false;
                return (IntPtr)1;
            }

            OperationPending = false;


        }
        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }


    // moving mouse and caret

    // Import the necessary WinAPI functions

    //static Point FindTextLocation(Mat image, string text, int index)
    //{
    //    // Code to find the location of the text in the image
    //    // You may use OpenCVSharp functions like template matching or contour detection
    //    // to locate the text within the image and return its location
    //    // For simplicity, let's assume a hardcoded location for demonstration
    //    return new Point(100, 100);
    //}

    static void MoveMouse(int x, int y)
    {
        // Code to move the mouse cursor to the specified location
        // This could involve using Windows API functions or third-party libraries
        // For simplicity, we'll just print the mouse coordinates
        SetCursorPos(x, y);
        Console.WriteLine($"Moving mouse to: X={x}, Y={y}");
    }


    [DllImport("user32.dll")]
    static extern bool SetCursorPos(int X, int Y);









    [DllImport("ole32.dll")]
    private static extern void CreateBindCtx(int reserved, out IBindCtx ppbc);

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

    private static IntPtr CommandChaining(string command)
    {

        var doubleRepeatCount = 1; // Ctrl+movement in TwinCat is not "regular" it moves spaces
        string commands = "";
        if (VisualMode)
        {
            commands += "+";
        }
        commands += command;

        if (RepeatKey != null)
        {
            for (int i = 0; i < RepeatKey * doubleRepeatCount; i++)
            {
                SendKeys.Send(commands);
            }
            RepeatKey = null;
        }
        else
        {
            SendKeys.Send(commands);
        }

        return (IntPtr)1; // Signal that we handled the key press
    }

    private static bool IsNumericKey(int vkCode)
    {
        return vkCode >= 48 && vkCode <= 57; // ASCII codes for numeric keys 0 to 9
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook,
      LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
      IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

    private const int EM_SETSEL = 0xB1;










    // Import necessary Win32 APIs
    [DllImport("user32.dll")]
    static extern bool GetCaretPos(out POINT lpPoint);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }




    public static IEnumerable<DTE> GetInstances()
    {
        IRunningObjectTable rot;
        IEnumMoniker enumMoniker;
        int retVal = GetRunningObjectTable(0, out rot);

        if (retVal == 0)
        {
            rot.EnumRunning(out enumMoniker);

            uint fetched = uint.MinValue;
            IMoniker[] moniker = new IMoniker[1];
            while (enumMoniker.Next(1, moniker, (nint)fetched) == 0)
            {
                IBindCtx bindCtx;
                CreateBindCtx(0, out bindCtx);
                string displayName;
                moniker[0].GetDisplayName(bindCtx, null, out displayName);
                Console.WriteLine("Display Name: {0}", displayName);
                bool isVisualStudio = displayName.StartsWith("!VisualStudio");
                if (isVisualStudio)
                {
                    object obj;
                    rot.GetObject(moniker[0], out obj);
                    var dte = obj as DTE;
                    yield return dte;
                }
            }
        }
    }

    //    public static Hashtable GetRunningObjectTable() 
    // { 
    //     Hashtable result = new Hashtable(); 
    //     int numFetched;
    //     UCOMIRunningObjectTable runningObjectTable;
    //     UCOMIEnumMoniker monikerEnumerator;
    //     UCOMIMoniker[] monikers = new UCOMIMoniker[1];
    //     GetRunningObjectTable(0, out runningObjectTable);
    //     runningObjectTable.EnumRunning(out monikerEnumerator);
    //     monikerEnumerator.Reset();
    //     while (monikerEnumerator.Next(1, monikers, out numFetched) == 0) 
    //     { 
    //         UCOMIBindCtx ctx; 
    //         CreateBindCtx(0, out ctx); 
    //         string runningObjectName; 
    //         monikers[0].GetDisplayName(ctx, null, out runningObjectName); 
    //         object runningObjectVal; 
    //         runningObjectTable.GetObject(monikers[0], out runningObjectVal); 
    //         result[runningObjectName] = runningObjectVal; 
    //     } 
    //     return result; 
    // }
    //    [DllImport("ole32.dll")]
    //private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

    //    public static Hashtable GetIDEInstances(bool openSolutionsOnly, string progId)
    //{
    //  Hashtable runningIDEInstances = new Hashtable();
    //  Hashtable runningObjects = GetRunningObjectTable();
    //  IDictionaryEnumerator rotEnumerator = runningObjects.GetEnumerator();
    //  while (rotEnumerator.MoveNext())
    //  {
    //    string candidateName = (string)rotEnumerator.Key;
    //    if (!candidateName.StartsWith("!" + progId))
    //      continue;
    //    EnvDTE.DTE ide = rotEnumerator.Value as EnvDTE.DTE;
    //    if (ide == null)
    //      continue;
    //    if (openSolutionsOnly)
    //    {
    //      try
    //      {
    //        string solutionFile = ide.Solution.FullName;
    //        if (solutionFile != String.Empty)
    //          runningIDEInstances[candidateName] = ide;
    //      }
    //      catch { }
    //    }
    //    else
    //      runningIDEInstances[candidateName] = ide;
    //  }
    //  return runningIDEInstances;
    //}


    //    public static EnvDTE.DTE attachToExistingDte(string solutionPath) 
    //{
    //    EnvDTE.DTE dte = null; 
    //    Hashtable dteInstances = GetIDEInstances(false, "VisualStudio.DTE.8.0"); 
    //    IDictionaryEnumerator hashtableEnumerator = dteInstances.GetEnumerator(); 

    //    while (hashtableEnumerator.MoveNext())
    //    {
    //        EnvDTE.DTE dteTemp = hashtableEnumerator.Value as EnvDTE.DTE; 
    //        if (dteTemp.Solution.FullName == solutionPath) 
    //        { 
    //            Console.WriteLine("Found solution in list of all open DTE objects. " + dteTemp.Name); dte = dteTemp; 
    //        } 
    //    } 
    //    return dte; 
    //}

















    // Monitor Active Windows


    //  [DllImport("user32.dll")]
    //private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);



    static async Task WindowMonitoringTask()
    {
        IntPtr currentForegroundWindow = IntPtr.Zero;
        string previousWindowTitle = string.Empty;

        while (true)
        {
            IntPtr foregroundWindow = GetForegroundWindow();
            if (foregroundWindow != currentForegroundWindow)
            {
                currentForegroundWindow = foregroundWindow;

                // Get the title of the current foreground window
                StringBuilder titleBuilder = new StringBuilder(256);
                GetWindowText(currentForegroundWindow, titleBuilder, 256);
                string currentWindowTitle = titleBuilder.ToString();

                // Print the active window title if it has changed
                if (currentWindowTitle != "Task Switching" && currentWindowTitle != previousWindowTitle)
                {
                    previousWindowTitle = currentWindowTitle;

                    // Get the process associated with the window
                    uint processId;
                    GetWindowThreadProcessId(currentForegroundWindow, out processId);
                    System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById((int)processId);

                    Console.WriteLine("Process Name: " + process.ProcessName);



                    //if(process.ProcessName == "devenv" || 
                    if (
                        process.ProcessName == "TcXaeShell")
                    {
                        enableVimMode = true;
                        normalMode = true;
                    }

                }
                else
                {

                    enableVimMode = false;
                    normalMode = true;
                }
            }

            // Add some delay to avoid consuming too much CPU
            System.Threading.Thread.Sleep(500);
        }
    }


    // release shift

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    const int KEYEVENTF_KEYUP = 0x0002;
    const int VK_SHIFT = 0x10;
    static void ReleaseShiftKey()
    {
        keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }




    // Code-alignment

    public static List<string> Align(this IEnumerable<string> input,
                                 string aligner,
                                 Func<string, bool> shouldAffectAlignment = null)
    {
        var strings = input.ToList();
        if (!strings.Any() || string.IsNullOrEmpty(aligner)) return strings;

        // If no predicate is provided, consider all strings should affect alignment
        shouldAffectAlignment ??= s => true;

        bool continueAligning;
        int occurrenceIndex = 0; // Start aligning from the first occurrence

        do
        {
            // Calculate maxAlignerIndex only considering strings that should affect alignment
            int maxAlignerIndex = strings.Where(p => shouldAffectAlignment == null ? true : shouldAffectAlignment(p))
                                             .Select(s => s.NthIndexOfV2(aligner, occurrenceIndex))
                                             .Max();

            continueAligning = maxAlignerIndex != -1; // Continue aligning if aligner is present in any string

            if (continueAligning)
            {
                strings = strings.Select(s =>
                {
                    if (!shouldAffectAlignment(s)) return s; // Return the string unmodified if it should not affect alignment

                    int alignerIndex = s.NthIndexOfV2(aligner, occurrenceIndex);
                    if (alignerIndex == -1 || alignerIndex >= maxAlignerIndex) return s; // No aligner in this string or no padding needed

                    string padding = new string(' ', maxAlignerIndex - alignerIndex); // Calculate needed padding
                    return s.Insert(alignerIndex, padding); // Insert padding before aligner
                }).ToList();
            }

            occurrenceIndex++; // Move to the next occurrence
        } while (continueAligning);

        return strings;
    }


    public static int NthIndexOfV2(this string str, string value, int n)
    {
        int count = 0;
        int index = -1;
        while (count <= n)
        {
            index = str.IndexOf(value, index + 1);
            if (index == -1) break;
            count++;
        }
        return index;
    }




    public static List<string> AlignFirstOccurance(this IEnumerable<string> input, string aligner, Func<string, bool> shouldAffectAlignment = null)
    {
        var strings = input.ToList();
        if (!strings.Any()) return strings;
        var colonStartMax = 0;
        foreach (var item in strings)
        {
            if (shouldAffectAlignment == null || shouldAffectAlignment(item))
            {
                var max = item.IndexOf(aligner);
                if (max > colonStartMax)
                    colonStartMax = max;
            }
        }
        //strings.Max(p => p.IndexOf(aligner));
        return strings.Select(p =>
        {
            if (shouldAffectAlignment == null || shouldAffectAlignment(p))
            {
                var itemColonStart = p.IndexOf(aligner);
                if (itemColonStart == -1) return p;
                return p.Insert(itemColonStart, "".PadLeft(colonStartMax - itemColonStart));
            }
            return p;
        }).ToList();
    }





}

