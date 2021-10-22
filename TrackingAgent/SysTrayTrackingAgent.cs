using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Net.Http;
using System.Diagnostics;
using Microsoft.Win32;
using System.Windows.Automation;
using System.Linq;

namespace TrackingAgent
{
    public partial class SysTrayTrackingAgent : Form
    {

        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;

        Timer tmr = new Timer(); // capture the tracking info
        Timer tmr2 = new Timer(); // send the captured info to server
        String filepath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/log.txt";
        readonly HttpClient client = new HttpClient();

        const int WH_KEYBOARD_LL = 13;
        const int WM_KEYDOWN = 0x100;
        private LowLevelKeyboardProc _proc = hookProc;
        public static int keypressedCount = 0;
        private static IntPtr hhook = IntPtr.Zero;


        private const int WH_MOUSE_LL = 14;
        private long timerFrequency = 0;
        private long lastWheelTick = 0;
        private LowLevelMouseProc _procMouse;
        private static IntPtr _hookID = IntPtr.Zero;

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc callback, IntPtr hInstance, uint threadId);

        [DllImport("user32.dll")]
        static extern bool UnhookWindowsHookEx(IntPtr hInstance);

        [DllImport("user32.dll")]
        static extern IntPtr CallNextHookEx(IntPtr idHook, int nCode, int wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        static extern IntPtr LoadLibrary(string lpFileName);

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);


        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern bool QueryPerformanceCounter(out long lpPerformanceCount);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll")]
        private static extern bool QueryPerformanceFrequency(out long lpFrequency);


        //Creating the extern function...  
        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hwnd, StringBuilder ss, int count);


        public SysTrayTrackingAgent()
        {
            InitializeComponent();
            this.TopMost = true;

            tmr.Interval = 1000; // 1sec
            tmr.Tick += Tmr_Tick;

            tmr2.Interval = 5000; // 5secs
            tmr2.Tick += Tmr_Tick2;

            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Exit", OnExit);            

            // Create a tray icon. In this example we use a
            // standard system icon for simplicity, but you
            // can of course use your own custom icon too.
            trayIcon = new NotifyIcon();
            trayIcon.Text = "Tracking Agent";
            trayIcon.Icon = new Icon(Directory.GetParent(System.Environment.CurrentDirectory).Parent.Parent.FullName + @"\images\tt_app.ico");
            trayIcon.Visible = true;

            // Add menu to tray icon and show it.
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;

            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            key.SetValue("TrackingAgent", Application.ExecutablePath);

            // get Windows product Id
            RegistryKey localMachine = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey windowsNTKey = localMachine.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion");
            object productID = windowsNTKey.GetValue("ProductId");
            string device_id = ""; //{\"device_id\":" +  "\"" + productID + "\""  + ",
            string device_name = "\"device_name\":" + "\"" + Environment.MachineName + "\"" + ",";
            string user_name = "\"user_name\":" + "\"" + Environment.UserName + "\"" + ",";
            string label_type = " \"label_type\": \"Activity-Monitoring\"" + ",";

            // Back Ground running Process
            Process[] processCollection = Process.GetProcesses();
            foreach (Process p in processCollection)
            {
                if (!string.IsNullOrEmpty(p.MainWindowTitle))
                {
                    Debug.WriteLine(device_id + device_name + user_name + label_type +
                                    "\"current_time\":" + "\"" + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + "\"" + "," +
                                    "\"label\":" + "\"" + p.ProcessName + "\"" + "," +
                                    "\"value\":" + "\"" + p.MainWindowTitle + "\" }");                    
                }
            }


            label_type = " \"label_type\": \"Brower-Window\"" + ",";
            // Crome Browser
            browserStatus(Process.GetProcessesByName("chrome"), device_id, device_name, user_name, label_type, "chrome");
            browserStatus(Process.GetProcessesByName("msedge"), device_id, device_name, user_name, label_type, "msedge");
            browserStatus(Process.GetProcessesByName("firefox"), device_id, device_name, user_name, label_type, "firefox");


            label_type = " \"label_type\": \"Windows_Explorer\"" + ",";
            // Code to get the system explorer (file maganger) opened file path
            foreach (SHDocVw.InternetExplorer window in new SHDocVw.ShellWindows())
            {
                if (Path.GetFileNameWithoutExtension(window.FullName).ToLowerInvariant() == "explorer")
                {
                    if (Uri.IsWellFormedUriString(window.LocationURL, UriKind.Absolute))

                        Debug.WriteLine(device_id + device_name + user_name + label_type +
                                    "\"current_time\":" + "\"" + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + "\"" + "," +
                                    "\"label\":" + "\"" + "chrome" + "\"" + "," +
                                    "\"value\":" + "\"" + new Uri(window.LocationURL).LocalPath + "\" }");

                    // Console.WriteLine(new Uri(window.LocationURL).LocalPath);
                }
            }

            // Internet Status
            int Desc;
            label_type = " \"label_type\": \"Interactivity-Monitoring\"" + ",";
            Debug.WriteLine(device_id + device_name + user_name + label_type +
                                    "\"current_time\":" + "\"" + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + "\"" + "," +
                                    "\"label\":" + "\"" + "Internet_Status" + "\"" + "," +
                                    "\"value\":" + "\"" + InternetGetConnectedState(out Desc, 0) + "\" }");
            

            // Speaker Status 
            IMMDeviceEnumerator enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            IMMDevice speakers = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
            IAudioMeterInformation meter = (IAudioMeterInformation)speakers.Activate(typeof(IAudioMeterInformation).GUID, 0, IntPtr.Zero);
            float value = meter.GetPeakValue();

            // this is a bit tricky. 0 is the official "no sound" value
            // but for example, if you open a video and plays/stops with it (w/o killing the app/window/stream),
            // the value will not be zero, but something really small (around 1E-09)
            // so, depending on your context, it is up to you to decide
            // if you want to test for 0 or for a small value
            Debug.WriteLine(device_id + device_name + user_name + label_type +
                                    "\"current_time\":" + "\"" + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + "\"" + "," +
                                    "\"label\":" + "\"" + "Speaker_Status" + "\"" + "," +
                                    "\"value\":" + "\"" + (value > 1E-08).ToString() + "\" }");
            


            // WebCam Status 
            Debug.WriteLine(device_id + device_name + user_name + label_type +
                                   "\"current_time\":" + "\"" + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + "\"" + "," +
                                   "\"label\":" + "\"" + "WebCam" + "\"" + "," +
                                   "\"value\":" + "\"" + IsWebCamInUse() + "\" }");
            
            // Mouse Status
            _procMouse = HookCallback;
            QueryPerformanceFrequency(out timerFrequency);
            _hookID = SetMouseHook(_procMouse);

            Debug.WriteLine(device_id + device_name + user_name + label_type +
                                  "\"current_time\":" + "\"" + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + "\"" + "," +
                                  "\"label\":" + "\"" + "Keyboard" + "\"" + "," +
                                  "\"value\":" + "\"" + keypressedCount + "\" }");

        }

        public static void UnHook()
        {
            UnhookWindowsHookEx(hhook);
        }

        public static void browserStatus(Process[] procsChrome, string device_id, string device_name, string user_name, string label_type, string browserType)
        {
            if (procsChrome.Length <= 0)
            {
                Console.WriteLine("Chrome is not running");
            }
            else
            {
                foreach (Process proc in procsChrome)
                {
                    // the chrome process must have a window 
                    if (proc.MainWindowHandle == IntPtr.Zero)
                    {
                        continue;
                    }
                    // to find the tabs we first need to locate something reliable - the 'New Tab' button 
                    AutomationElement root = AutomationElement.FromHandle(proc.MainWindowHandle);
                    //Condition condNewTab = new PropertyCondition(AutomationElement.NameProperty, "New Tab");
                    //AutomationElement elmNewTab = root.FindFirst(TreeScope.Descendants, condNewTab);
                    Condition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
                    var tabs = root.FindFirst(TreeScope.Descendants, condition);

                    // get the tabstrip by getting the parent of the 'new tab' button 
                    TreeWalker treewalker = TreeWalker.ControlViewWalker;
                    if (tabs != null)
                    {
                        AutomationElement elmTabStrip = treewalker.GetParent(tabs);
                        if (elmTabStrip != null)
                        {
                            // loop through all the tabs and get the names which is the page title 
                            Condition condTabItem = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
                            foreach (AutomationElement tabitem in elmTabStrip.FindAll(TreeScope.Children, condTabItem))
                            {
                                Debug.WriteLine(device_id + device_name + user_name + label_type +
                                    "\"current_time\":" + "\"" + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + "\"" + "," +
                                    "\"label\":" + "\"" + browserType + "\"" + "," +
                                    "\"value\":" + "\"" + tabitem.Current.Name + "\" }");                                
                            }
                        }
                    }

                }
            }


        }


        /// <summary>
        /// Keyboard hook function
        /// </summary>
        /// <param name="code"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        public static IntPtr hookProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                keypressedCount = keypressedCount + 1;
                Debug.WriteLine("key was pressed " + keypressedCount);
                return CallNextHookEx(hhook, code, (int)wParam, lParam);
            }
            else
                return CallNextHookEx(hhook, code, (int)wParam, lParam);
        }


        /// <summary>
        /// set mouse hook
        /// </summary>
        /// <param name="proc"></param>
        /// <returns></returns>
        private static IntPtr SetMouseHook(LowLevelMouseProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }


        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {

            if (nCode >= 0 && MouseMessages.WM_MOUSEWHEEL == (MouseMessages)wParam)
            {
                long wheelTick = 0;
                QueryPerformanceCounter(out wheelTick);
                long diff = wheelTick - lastWheelTick;
                long RPM = (timerFrequency / diff) / 60;
                lastWheelTick = wheelTick;
                Debug.WriteLine("Mouse Dfference " + diff);
            }
            // Debug.WriteLine("Mouse  WM_MOUSEMOVE " + MouseMessages.WM_MOUSEMOVE);
            // Debug.WriteLine("Mouse  WM_LBUTTONUP " + MouseMessages.WM_LBUTTONUP);
            // Debug.WriteLine("Mouse  WM_LBUTTONDOWN " + MouseMessages.WM_LBUTTONDOWN);
            //  Debug.WriteLine("Mouse  WM_RBUTTONUP " + MouseMessages.WM_RBUTTONUP);
            // Debug.WriteLine("Mouse  WM_RBUTTONDOWN " + MouseMessages.WM_RBUTTONDOWN);
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private enum MouseMessages
        {
            WM_LBUTTONDOWN = 0x0201,
            WM_LBUTTONUP = 0x0202,
            WM_MOUSEMOVE = 0x0200,
            WM_MOUSEWHEEL = 0x020A,
            WM_RBUTTONDOWN = 0x0204,
            WM_RBUTTONUP = 0x0205
        }

        //---------------------------------------------------------------------------------------------

        private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            UnHook();
        }

        public void SetHook()
        {
            IntPtr hInstance = LoadLibrary("User32");
            hhook = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, hInstance, 0);
        }

        /// <summary>
        /// web cam usage
        /// </summary>
        /// <returns></returns>
        private static bool IsWebCamInUse()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam\NonPackaged"))
            {
                foreach (var subKeyName in key.GetSubKeyNames())
                {
                    using (var subKey = key.OpenSubKey(subKeyName))
                    {
                        if (subKey.GetValueNames().Contains("LastUsedTimeStop"))
                        {
                            var endTime = subKey.GetValue("LastUsedTimeStop") is long ? (long)subKey.GetValue("LastUsedTimeStop") : -1;
                            if (endTime <= 0)
                            {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        private class MMDeviceEnumerator
        {
        }

        private enum EDataFlow
        {
            eRender,
            eCapture,
            eAll,
        }

        private enum ERole
        {
            eConsole,
            eMultimedia,
            eCommunications,
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
        private interface IMMDeviceEnumerator
        {
            void NotNeeded();
            IMMDevice GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role);
            // the rest is not defined/needed
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("D666063F-1587-4E43-81F1-B948E807363F")]
        private interface IMMDevice
        {
            [return: MarshalAs(UnmanagedType.IUnknown)]
            object Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, int dwClsCtx, IntPtr pActivationParams);
            // the rest is not defined/needed
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("C02216F6-8C67-4B5B-9D00-D008E73E0064")]
        private interface IAudioMeterInformation
        {
            float GetPeakValue();
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.

            base.OnLoad(e);
        }

        private void OnExit(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Do you want to exit", "Tracking Agent",
                                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            { 
                tmr.Stop();
                tmr2.Stop();
                UnHook();
                Application.Exit();
            }
        }

        private void Tmr_Tick(object sender, EventArgs e)
        {
            //get title of active window
            string title = ActiveWindowTitle();
            //check if it is null and add it to list if correct
            if (title != "")
            {
                try
                { 
                    using(TextWriter writeFile = new StreamWriter(filepath, true))
                    { 
                        //write data into file
                        writeFile.WriteLine("{\"timestamp\":\"" + DateTime.Now.ToString("hh:mm:ss") + "\",\"title\":\"" + title + "\"}");
                        writeFile.Flush();
                        writeFile.Close();
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }

        private async void Tmr_Tick2(object sender, EventArgs e)
        {   
            /*
            // stop the write timer till reading is done
            tmr.Stop();
            StringBuilder content = new StringBuilder();
            string str = String.Empty;
            try
            {
                using (TextReader readFile = new StreamReader(filepath))
                {
                    while ((str = readFile.ReadLine()) != null)
                    {
                        content.Append(str + ",");
                    }
                    content.Length--;
                    readFile.Close();                    
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            // start the write timer
            tmr.Start();
            Console.WriteLine("file content: " + content);

            // send the data to server
            string Url = "http://localhost:3000/create";
            var strcontent = new StringContent("[" + content + "]", Encoding.UTF8, "application/json");

            HttpResponseMessage responseMessage = null;
            try
            {
                using (responseMessage = await client.PostAsync(Url, strcontent))
                {
                    string result = await responseMessage.Content.ReadAsStringAsync();
                    Console.WriteLine("http result: " + result);
                    // successfully sent the data, hence clear it in the local storage
                    File.WriteAllText(filepath, "");
                }                
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            */
        }

        private string ActiveWindowTitle()
        {
            //Create the variable
            const int nChar = 256;
            StringBuilder ss = new StringBuilder(nChar);

            //Run GetForeGroundWindows and get active window informations
            //assign them into handle pointer variable
            IntPtr handle = IntPtr.Zero;
            handle = GetForegroundWindow();

            if (GetWindowText(handle, ss, nChar) > 0) return ss.ToString();
            else return "";
        }

        private void SysTrayTrackingAgent_Load(object sender, EventArgs e)
        {
            tmr.Start();
            tmr2.Start();
            SetHook();
        }

    }
}
