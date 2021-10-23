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
using System.Collections.Generic;

namespace TrackingAgent
{
    public partial class SysTrayTrackingAgent : Form
    {
        /// initialize the variables
        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
                
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x100;
        private const int WH_MOUSE_LL = 14;

        private static int keypressedCount = 0;
        private static int mouseLeftClickCount = 0;
        private static int mouseRightClickCount = 0;
        private static IntPtr keyboardhook = IntPtr.Zero;
        private static IntPtr mousehook = IntPtr.Zero;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        private LowLevelKeyboardProc _keyboardHookProc = keyboardHookProc;
        private LowLevelMouseProc _mouseHookProc = mouseHookProc;        

        private Timer tmr = new Timer(); // capture the tracking info
        private Timer tmr2 = new Timer(); // send the captured info to server
        private HttpClient client = new HttpClient(); // get the HttpClient
        private String filepath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/log.txt";

        /// DLL imports
        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc proc, IntPtr lParam, uint threadId);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc proc, IntPtr lParam, uint threadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hInstance);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr idHook, int nCode, int wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern IntPtr LoadLibrary(string lpFileName);            
        
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern bool QueryPerformanceCounter(out long lpPerformanceCount);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll")]
        private static extern bool QueryPerformanceFrequency(out long lpFrequency);

        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hwnd, StringBuilder ss, int count);

        /// static methods
        // get active window
        private static string getActiveWindowTitle()
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

        // get browser info
        private static List<string> getInternetBrowserTabNames(Process[] procs)
        {
            List<String> tabNames = new List<string>();
            if (procs.Length > 0)
            {
                foreach (Process proc in procs)
                {
                    // the chrome process must have a window 
                    if (proc.MainWindowHandle == IntPtr.Zero)
                    {
                        continue;
                    }
                    // to find the tabs we first need to locate something reliable - the 'New Tab' button 
                    AutomationElement root = AutomationElement.FromHandle(proc.MainWindowHandle);
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
                                tabNames.Add(tabitem.Current.Name);
                            }
                        }
                    }
                }
            }
            return tabNames;
        }

        // Keyboard hook function
        private static IntPtr keyboardHookProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                keypressedCount = keypressedCount + 1;
            }
            return CallNextHookEx(keyboardhook, code, (int)wParam, lParam);
        }

        // mouse hook function
        private static IntPtr mouseHookProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && wParam == (IntPtr)MouseMessages.WM_LBUTTONDOWN)
            {
                mouseLeftClickCount = mouseLeftClickCount + 1;
            }
            else if (code >= 0 && wParam == (IntPtr)MouseMessages.WM_RBUTTONDOWN)
            {
                mouseRightClickCount = mouseRightClickCount + 1;
            }
            return CallNextHookEx(mousehook, code, (int)wParam, lParam);
        }

        // web cam usage
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

        /// declarations
        private enum MouseMessages
        {
            WM_LBUTTONDOWN = 0x0201,
            // WM_LBUTTONUP = 0x0202,
            // WM_MOUSEMOVE = 0x0200,
            // WM_MOUSEWHEEL = 0x020A,
            WM_RBUTTONDOWN = 0x0204,
            // WM_RBUTTONUP = 0x0205
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
             IMMDevice GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role);
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("D666063F-1587-4E43-81F1-B948E807363F")]
        private interface IMMDevice
        {
            [return: MarshalAs(UnmanagedType.IUnknown)]
            object Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, int dwClsCtx, IntPtr pActivationParams);
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("C02216F6-8C67-4B5B-9D00-D008E73E0064")]
        private interface IAudioMeterInformation
        {
            float GetPeakValue();
        }

        /// constructor
        public SysTrayTrackingAgent()
        {
            InitializeComponent();
            this.TopMost = true;

            // set the log write timer to every 30secs
            tmr.Interval = 1000 * 5; // 5secs
            tmr.Tick += Tmr_Tick;

            // set the server write timer to every 15mins
            tmr2.Interval = 1000 * 60 * 15; // 15mins
            tmr2.Tick += Tmr_Tick2;

            // create a tray menu
            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Exit", OnExit);            

            // Create a tray icon. 
            trayIcon = new NotifyIcon();
            trayIcon.Text = "Tracking Agent";
            trayIcon.Icon = new Icon(Directory.GetParent(System.Environment.CurrentDirectory).Parent.Parent.FullName + @"\images\tt_app.ico");
            trayIcon.Visible = true;

            // Add menu to tray icon and show it.
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;

            // Set the app to run on startup
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            key.SetValue("TrackingAgent", Application.ExecutablePath);

        }

        /// overloaded method
        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.
            base.OnLoad(e);
        }

        private void SysTrayTrackingAgent_Load(object sender, EventArgs e)
        {
            tmr.Start();
            tmr2.Start();
            SetHook();
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

        ///  private functions
        private void SetHook()
        {
            IntPtr hInstance = LoadLibrary("User32");
            keyboardhook = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardHookProc, hInstance, 0);
            mousehook = SetWindowsHookEx(WH_MOUSE_LL, _mouseHookProc, hInstance, 0);
        }

        private void UnHook()
        {
            UnhookWindowsHookEx(keyboardhook);
            UnhookWindowsHookEx(mousehook);
        }

        private void Tmr_Tick(object sender, EventArgs e)
        {
            try
            { 
                using(TextWriter writeFile = new StreamWriter(filepath, true))
                {
                    // get the current date and time
                    string dateTime = DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss");

                    //write (active window monitoring) data into file
                    // get active window title                    
                    string label = "None";
                    string activeWindowTitle = getActiveWindowTitle();
                    if (activeWindowTitle != "")
                    {
                        // seperating the label and value from the Title
                        if (activeWindowTitle.LastIndexOf("|") != -1)
                        {
                            int lastIndex = activeWindowTitle.LastIndexOf("|");
                            label = activeWindowTitle.Substring(lastIndex + 2);
                            activeWindowTitle = activeWindowTitle.Substring(0, lastIndex - 1);
                        }
                        else if (activeWindowTitle.LastIndexOf("-") != -1)
                        {
                            int lastIndex = activeWindowTitle.LastIndexOf("-");
                            label = activeWindowTitle.Substring(lastIndex + 2);
                            activeWindowTitle = activeWindowTitle.Substring(0, lastIndex - 1);                            
                        }
                        writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Active Window Monitoring - Active Window\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"" + label + "\",\"value\":\"" + activeWindowTitle + "\"}");
                    }

                    // Explorer (file maganger) opened with file path
                    foreach (SHDocVw.InternetExplorer window in new SHDocVw.ShellWindows())
                    {
                        if (Path.GetFileNameWithoutExtension(window.FullName).ToLowerInvariant() == "explorer")
                        {
                            if (Uri.IsWellFormedUriString(window.LocationURL, UriKind.Absolute))
                            {
                                writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Active Window Monitoring - Opened Windows\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"File Explorer\",\"value\":\"" + new Uri(window.LocationURL).LocalPath + "\"}");
                            }
                        }
                    }

                    // Chrome browser
                    List<string> chromeTabs = getInternetBrowserTabNames(Process.GetProcessesByName("chrome"));
                    foreach (string chromeTab in chromeTabs)
                    {
                        writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Active Window Monitoring - Opened Windows\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"Google Chrome\",\"value\":\"" + chromeTab + "\"}");
                    }
                    // Edge browser
                    List<string> msedgeTabs = getInternetBrowserTabNames(Process.GetProcessesByName("msedge"));
                    foreach (string msedgeTab in msedgeTabs)
                    {
                        writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Active Window Monitoring - Opened Windows\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"Microsoft Edge\",\"value\":\"" + msedgeTab + "\"}");
                    }
                    // Firefox browser
                    List<string> firefoxTabs = getInternetBrowserTabNames(Process.GetProcessesByName("firefox"));
                    foreach (string firefoxTab in firefoxTabs)
                    {
                        writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Active Window Monitoring - Opened Windows\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"Firefox\",\"value\":\"" + firefoxTab + "\"}");
                    }

                    // Back ground running process
                    Process[] processCollection = Process.GetProcesses();
                    foreach (Process process in processCollection)
                    {
                        if (!string.IsNullOrEmpty(process.MainWindowTitle))
                        {
                            writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Active Window Monitoring - Background Process\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"" + process.ProcessName + "\",\"value\":\"" + process.MainWindowTitle + "\"}");
                        }
                    }

                    // Internet Status
                    int internetStatus;
                    writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Interactivity Monitoring - Internet Status\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"Is internet available?\",\"value\":\"" + InternetGetConnectedState(out internetStatus, 0) + "\"}");

                    // number of keys pressed
                    writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Interactivity Monitoring - Keyboard\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"Number of keys pressed\",\"value\":\"" + keypressedCount + "\"}");
                    keypressedCount = 0; // reset the value

                    // mouse left click count
                    writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Interactivity Monitoring - Mouse\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"Mouse left click count\",\"value\":\"" + mouseLeftClickCount + "\"}");
                    mouseLeftClickCount = 0; // reset the value

                    // mouse right click count
                    writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Interactivity Monitoring - Mouse\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"Mouse right click count\",\"value\":\"" + mouseRightClickCount + "\"}");
                    mouseRightClickCount = 0; // reset the value

                    // WebCam Status 
                    writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Interactivity Monitoring - Camera\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"Is camera in use?\",\"value\":\"" + IsWebCamInUse() + "\"}");

                    // Speaker Status 
                    // IMMDeviceEnumerator enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
                    // IMMDevice speakers = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
                    // IAudioMeterInformation meter = (IAudioMeterInformation)speakers.Activate(typeof(IAudioMeterInformation).GUID, 0, IntPtr.Zero);
                    // float value = meter.GetPeakValue();

                    // this is a bit tricky. 0 is the official "no sound" value
                    // but for example, if you open a video and plays/stops with it (w/o killing the app/window/stream),
                    // the value will not be zero, but something really small (around 1E-09)
                    // so, depending on your context, it is up to you to decide
                    // if you want to test for 0 or for a small value
                    //writeFile.WriteLine("{\"device_name\":\"" + Environment.MachineName + "\",\"user_name\": \"" + Environment.UserName + "\",\"label_type\": \"Interactivity Monitoring - Speaker\",\"date_time\":\"" + dateTime + "\",\"label\":" + "\"Is speaker in use?\",\"value\":\"" + (value > 1E-08).ToString() + "\"}");
                    
                    
                    writeFile.Flush();
                    writeFile.Close();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async void Tmr_Tick2(object sender, EventArgs e)
        {               
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
        }
    }
}
