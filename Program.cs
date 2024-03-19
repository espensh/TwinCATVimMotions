using EnvDTE;
using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using EnvDTE80;
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


    private static bool YankCommand { get; set; }
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
            bool isUpperCase = (!isShiftPressed && !isCapsLockOn) || (isShiftPressed && isCapsLockOn);


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

                if (isUpperCase && pressedChar == 'Y')
                {
                    buffer.Enqueue(pressedChar);

                    // Limit the buffer size to 2 to check for 'yy'
                    if (buffer.Count > 2)
                        buffer.Dequeue();

                    if (buffer.Count == 2 && buffer.Peek() == 'Y' && pressedChar == 'Y')
                    {
                        buffer.Clear();

                        SendKeys.SendWait("{HOME}");
                        System.Threading.Thread.Sleep(100); // wait a bit for the caret to move

                        // Simulate pressing Shift+End to select the entire line
                        SendKeys.SendWait("+{END}");
                        System.Threading.Thread.Sleep(100); // wait a bit for the line to be selected

                        SendKeys.Send("^(c)");    // Trigger the copy action
                    }

                }

                if (isControlPressed && pressedChar == 'W')
                {
                    movePanesMode = true;
                    return (IntPtr)1;
                }


                if ((Keys)vkCode == Keys.Y)
                {

                    return (IntPtr)1;
                }

                if (IsNumericKey(vkCode))
                {
                    Console.WriteLine($"Numeric key pressed: {(Keys)vkCode}");
                    RepeatKey = vkCode - 48;
                    return (IntPtr)1; // Signal that we handled the key press
                }

                if ((Keys)vkCode == Keys.C)
                {
                    ChangeCommand = true;
                    return (IntPtr)1;
                }
                if ((Keys)vkCode == Keys.H)
                {
                    return CommandChaining("{LEFT}");
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
                if ((Keys)vkCode == Keys.L)
                {
                    return CommandChaining("{RIGHT}");
                }
                if (!isControlPressed && (Keys)vkCode == Keys.W)
                {
                    return CommandChaining("^{RIGHT}");
                }
                if ((Keys)vkCode == Keys.B)
                {
                    return CommandChaining("^{LEFT}");
                }

                if ((Keys)vkCode == Keys.I)
                {
                    normalMode = false;
                    return (IntPtr)1;
                }


                if ((Keys)vkCode == Keys.V)
                {
                    VisualMode = !VisualMode;

                    return (IntPtr)1; // Signal that we handled the key press
                }
            }


                if ((Keys)vkCode == Keys.Escape)
                {
                    normalMode = true;
                    return (IntPtr)1;
                }


        }
        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    [DllImport("ole32.dll")]
    private static extern void CreateBindCtx(int reserved, out IBindCtx ppbc);

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

    private static IntPtr CommandChaining(string command)
    {

        string commands = "";
        if (VisualMode)
        {
            commands += "+";
        }
        commands += command;

        if (RepeatKey != null)
        {
            for (int i = 0; i < RepeatKey; i++)
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

    [DllImport("user32.dll")]
    static extern bool SetCursorPos(int x, int y);

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
}
