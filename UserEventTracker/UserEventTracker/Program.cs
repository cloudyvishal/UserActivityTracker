using System;
using System.Xml;
using Microsoft.Win32;
using System.Threading;
using System.Management;
using System.Diagnostics;
using System.Windows.Automation;
using System.Collections.Generic;
using System.Security.Permissions;
using System.Runtime.InteropServices;
using System.IO;

namespace UserEventTracker
{
    class Program
    {
        [DllImport("shell32")]
        static extern bool IsUserAnAdmin();

        [DllImport("User32.dll")]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [DllImport("Kernel32.dll")]
        private static extern uint GetLastError();

        public static string xmlFileName = "UserTrackngApp" + DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH'_'mm'_'ss") + ".xml";

        public static string path = Directory.GetCurrentDirectory() + "\\" + xmlFileName;

        public static XmlWriter xmlWriter = XmlWriter.Create(path);
        static void Main(string[] args)
        {
            xmlWriter.WriteStartDocument();

            xmlWriter.WriteComment("First Comment XmlTextWriter Sample Example");

            xmlWriter.WriteComment("myXmlFile in root dir");

            xmlWriter.WriteStartElement("users");

            GetAllBrowsersTabName();

            //DisableUSBDevice();

            GetListOfUsbDevices();

            GetInstalledApps();

            GetAllProcessMemoryUsage();

            AppExecuteAsAdmin();

            GetIdealTime();

        }

        private static void GetIdealTime()
        {
            xmlWriter.WriteStartElement("r", "RECORD", "urn:record");

            xmlWriter.WriteStartElement("IdealTime");

            xmlWriter.WriteAttributeString("Current_Time", "NA");

            xmlWriter.WriteString($"Current_Time {DateTimeOffset.Now}");

            xmlWriter.WriteEndElement();

           // xmlWriter.WriteAttributeString("Last_input_time", "NA");

            xmlWriter.WriteString($"Last_input_time: {InputTimer.GetLastInputTime()}");

            xmlWriter.WriteEndElement();

           // xmlWriter.WriteAttributeString("Last_input_time", "NA");

            xmlWriter.WriteString($"Idle_time {InputTimer.GetInputIdleTime()}");

            xmlWriter.WriteEndElement();

            xmlWriter.WriteEndDocument();
            // close writer  
            xmlWriter.Close();

            //Console.WriteLine($"Idle time: {DateTimeOffset.Now}");

            // Console.WriteLine($"Last input time: {InputTimer.GetLastInputTime()}");

            // Console.WriteLine($"Idle time: {InputTimer.GetInputIdleTime()}");
        }

        private static void GetAllBrowsersTabName()
        {
            Process[] procsChrome = Process.GetProcessesByName("chrome");
            if (procsChrome.Length <= 0)
            {
                //Console.WriteLine("Chrome is not running");
                xmlWriter.WriteStartElement("BrowserChrome");

                xmlWriter.WriteAttributeString("Chrome", "NA");

                xmlWriter.WriteString("Chrome is not running");

                xmlWriter.WriteEndElement();
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
                    AutomationElement root = AutomationElement.FromHandle(proc.MainWindowHandle);

                    Condition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);

                    var tabs = root.FindAll(TreeScope.Descendants, condition);

                    foreach (AutomationElement tabitem in tabs)
                    {
                        string urlname = tabitem.Current.Name;

                        //Console.WriteLine(urlname);
                        xmlWriter.WriteStartElement("BrowserChrome");

                        xmlWriter.WriteAttributeString("Chrome", tabitem.Current.ProcessId.ToString());

                        xmlWriter.WriteString(urlname);

                        xmlWriter.WriteEndElement();
                    }
                }
            }

            Process[] procsEdge = Process.GetProcessesByName("msedge");
            if (procsEdge.Length <= 0)
            {
                //Console.WriteLine("Edge is not running");
                xmlWriter.WriteStartElement("BrowserEdge");

                xmlWriter.WriteAttributeString("Edge", "NA");

                xmlWriter.WriteString("Edge is not running");

                xmlWriter.WriteEndElement();
            }
            else
            {
                foreach (Process proc in procsEdge)
                {
                    // the chrome process must have a window
                    if (proc.MainWindowHandle == IntPtr.Zero)
                    {
                        continue;
                    }
                    AutomationElement root = AutomationElement.FromHandle(proc.MainWindowHandle);

                    Condition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);

                    var tabs = root.FindAll(TreeScope.Descendants, condition);

                    foreach (AutomationElement tabitem in tabs)
                    {
                        string urlname = tabitem.Current.Name;

                        //Console.WriteLine(urlname);
                        xmlWriter.WriteStartElement("BrowserChrome");

                        xmlWriter.WriteAttributeString("Chrome", tabitem.Current.ProcessId.ToString());

                        xmlWriter.WriteString(urlname);

                        xmlWriter.WriteEndElement();
                    }
                }
            }

            Process[] procsFirefox = Process.GetProcessesByName("firefox");

            if (procsFirefox.Length <= 0)
            {
                //Console.WriteLine("Firefox is not running");
                xmlWriter.WriteStartElement("BrowserFireFox");

                xmlWriter.WriteAttributeString("FireFox", "NA");

                xmlWriter.WriteString("FireFox is not running");

                xmlWriter.WriteEndElement();
            }
            else
            {
                foreach (Process proc in procsFirefox)
                {
                    // the chrome process must have a window
                    if (proc.MainWindowHandle == IntPtr.Zero)
                    {
                        continue;
                    }
                    AutomationElement root = AutomationElement.FromHandle(proc.MainWindowHandle);

                    Condition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);

                    var tabs = root.FindAll(TreeScope.Descendants, condition);

                    foreach (AutomationElement tabitem in tabs)
                    {
                        string urlname = tabitem.Current.Name;

                        //Console.WriteLine(urlname);
                        xmlWriter.WriteStartElement("BrowserChrome");

                        xmlWriter.WriteAttributeString("Chrome", tabitem.Current.ProcessId.ToString());

                        xmlWriter.WriteString(urlname);

                        xmlWriter.WriteEndElement();
                    }
                }
            }
            Console.ReadKey();
        }

        private static void AppExecuteAsAdmin()
        {
            //ProcessStartInfo proc = new ProcessStartInfo();

            //proc.UseShellExecute = true;

            //proc.WorkingDirectory = Environment.CurrentDirectory;

            //proc.FileName = Application.ExecutablePath;

            //proc.Verb = "runas";



            //try

            //{

            //    Process.Start(proc);

            //}

            //catch

            //{

            //    // The user refused the elevation.

            //    // Do nothing and return directly ...

            //    return;

            //}

            //Application.Exit();  // Quit itself
        }

        private static void GetAllProcessMemoryUsage()
        {
            Process[] processCollection = Process.GetProcesses();

            foreach (Process p in processCollection)
            {
                long usedMemory = p.PrivateMemorySize64;

                // Console.WriteLine("Process Name:-\t " + p.ProcessName + "  \t Memory Usage:-  " + usedMemory);
                xmlWriter.WriteStartElement("ProcessMemoryUsage");

                xmlWriter.WriteAttributeString("ProcessName", p.ProcessName);

                xmlWriter.WriteString("MemoryUsage: -" + usedMemory);

                xmlWriter.WriteEndElement();
            }
        }

        /// <summary>
        /// Get All Installed Software Details
        /// </summary>
        private static void GetInstalledApps()
        {
            string registry_key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";

            using (Microsoft.Win32.RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key))
            {
                foreach (string subkey_name in key.GetSubKeyNames())
                {
                    using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                    {
                        //Console.WriteLine(subkey.GetValue("DisplayName"));
                        string dispalyName = (string)subkey.GetValue("DisplayName");

                        xmlWriter.WriteStartElement("InstalledApps");

                        xmlWriter.WriteAttributeString("RegistryKey", "Value");

                        xmlWriter.WriteString(dispalyName);

                        xmlWriter.WriteEndElement();
                    }
                }
            }
        }

        internal struct LASTINPUTINFO
        {
            public uint cbSize;

            public uint dwTime;
        }

        private static void GetListOfUsbDevices()
        {
            ManagementObjectCollection mbsList = null;

            ManagementObjectSearcher mbs = new ManagementObjectSearcher("Select * From Win32_USBHub");

            mbsList = mbs.Get();

            foreach (ManagementObject mo in mbsList)
            {
                // Console.WriteLine("USBHub device Friendly name:{0}", mo["Name"].ToString());
                xmlWriter.WriteStartElement("UsbDevices");

                xmlWriter.WriteAttributeString("USBHubdevice", "Friendlyname");

                xmlWriter.WriteString(mo["Name"].ToString());

                xmlWriter.WriteEndElement();
            }
        }

        private static void DisableUSBDevice()
        {
            bool isAdmin = IsUserAnAdmin();

            RegistryKey Regkey, RegKey2;

            Int32 rValue, rsvalue;

            string Regpath = "System\\CurrentControlSet\\Services\\USBSTOR";

            string ReadAndWriteRegPath2 = "System\\CurrentControlSet\\Control";

            string ReadAndWriteRegPath = "System\\CurrentControlSet\\Control\\StorageDevicePolicies";

            rValue = 4;

            rsvalue = 1;

            Regkey = Registry.LocalMachine.OpenSubKey(Regpath, true);

            Regkey.SetValue("Start", rValue);

            RegKey2 = Registry.LocalMachine.OpenSubKey(ReadAndWriteRegPath2, true);

            RegKey2.CreateSubKey("StorageDevicePolicies");

            RegKey2 = Registry.LocalMachine.OpenSubKey(ReadAndWriteRegPath, true);

            RegKey2.SetValue("WriteProtect", rsvalue);
        }
    }
}
