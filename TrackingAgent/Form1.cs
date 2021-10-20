using System;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Collections;
using System.Collections.Generic;

namespace TrackingAgent
{
    public partial class TrackingAgent : Form
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hwnd, StringBuilder ss, int count);

        Timer tmr = new Timer(); // capture the tracking info
        Timer tmr2 = new Timer(); // send the captured info to server
        String filepath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/log.txt";
        readonly HttpClient client = new HttpClient();

        public TrackingAgent()
        {
            InitializeComponent();
            this.TopMost = true;

            if (!File.Exists(filepath))
            {
                File.Create(filepath);
            }

            tmr.Interval = 1000 * 30; // 30secs
            tmr.Tick += Tmr_Tick;

            tmr2.Interval = 1000 * 60 * 15; // 15mins
            tmr2.Tick += Tmr_Tick2;            
        }

        private void Tmr_Tick(object sender, EventArgs e)
        {
            //get title of active window
            string title = ActiveWindowTitle();
            //check if it is null and add it to list if correct
            if (title != "")
            {
                TextWriter writeFile = new StreamWriter(filepath, true);
                //write data into file
                writeFile.WriteLine("{\"timestamp\":\""+DateTime.Now.ToString("hh:mm:ss")+"\",\"title\":\""+title+"\"}");
                writeFile.Flush();
                writeFile.Close();
            }
        }

        private void Tmr_Tick2(object sender, EventArgs e)
        {            
            tmr.Stop();
            StringBuilder content = new StringBuilder();
            string str = String.Empty;
            TextReader readFile = new StreamReader(filepath);
            while ((str = readFile.ReadLine()) != null)
            {
                content.Append(str + ",");
            }
            content.Length--;
            readFile.Close();
            File.WriteAllText(filepath, "");
            tmr.Start();
            Console.WriteLine("file content: " + content);
            var strcontent = new StringContent("[" + content + "]", Encoding.UTF8, "application/json");
            var response = client.PostAsync("http://localhost:3000/create", strcontent);
            string result = ((StreamContent)response.Result.Content).ReadAsStringAsync().Result;
            Console.WriteLine("http result: " + result);            
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

        private void button_Click(object sender, EventArgs e)
        {
            if (status.Text == "Running")
            {
                status.Text = "Stopped";
                button.Text = "Start";
                tmr.Stop();
                tmr2.Stop();
            }
            else
            {
                status.Text = "Running";
                button.Text = "Stop";
                tmr.Start();
                tmr2.Start();
            }
        }
    }
}
