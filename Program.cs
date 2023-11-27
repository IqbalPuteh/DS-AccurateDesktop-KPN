using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;
using FlaUI.UIA3;
using System.Configuration;
using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;

namespace DSAccurateDesktopKPN
{
    internal class Program
    {
        static Application appx;
        static Window DesktopWindow;
        static UIA3Automation automationUIA3 = new UIA3Automation();
        static ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());
        static AutomationElement window = automationUIA3.GetDesktop();
        static System.Diagnostics.TextWriterTraceListener logListener;
        static int pid;
        static int step = 0;
        static string dtID = ConfigurationManager.AppSettings["dtID"];
        static string dtName = ConfigurationManager.AppSettings["dtName"];
        static string appExe = ConfigurationManager.AppSettings["erpappnamepath"];
        static string loginId = ConfigurationManager.AppSettings["loginId"];
        static string loginPassword = ConfigurationManager.AppSettings["password"];
        static string DBpath = ConfigurationManager.AppSettings["DBaddresspath"].ToUpper();
        static string DBaliasflag = ConfigurationManager.AppSettings["DBaliasflag"].ToUpper();
        static string DBaliasname = ConfigurationManager.AppSettings["DBaliasname"];
        static string issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
        static string issecurehttp = ConfigurationManager.AppSettings["issecurehttp"];
        static string isacctrptincluded = ConfigurationManager.AppSettings["isacctrptincluded"];
        static string isconsolelogenable = ConfigurationManager.AppSettings["isconsolelogenable"].ToUpper();
        static string appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
        static string uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
        static string sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
        static string logfilename;

        [DllImport("user32.dll")]
        public static extern bool BlockInput(bool fBlockIt);

        private static AutomationElement WaitForElement(Func<AutomationElement> findElementFunc)
        {
            AutomationElement element = null;
            for (int i = 0; i < 2000; i++)
            {
                element = findElementFunc();
                if (element != null)
                {
                    break;
                }

                Thread.Sleep(1);
            }
            return element;
        }

        private static void InitializeLogger(string path, string filename, string enableConsoleLogging = "N")
        {
            // Customize the log file path and name here
            var logfname = Path.Combine(path, filename);

            // Create a new text writer for logging to a file
            logListener = new TextWriterTraceListener(logfname);

            // Add the listener to the trace sources
            Trace.Listeners.Add(logListener);

            // Optionally, add a console listener if writeToConsole is true
            if (enableConsoleLogging == "Y")
            {
                Trace.Listeners.Add(new ConsoleTraceListener());
            }

            // Set the level of detail you want to log
            Trace.AutoFlush = true;
        }

        static void Log(string message)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Trace.Write($"[{DateTime.Now.ToString("dd-MM-yy hh:mm:ss")} INF] ");
            Console.ForegroundColor = ConsoleColor.White;
            Trace.WriteLine(message.TrimEnd());
        }

        private static void CloseLogger()
        {
            if (logListener != null)
            {
                logListener.Close();
                logListener.Dispose();
            }
        }

        static void Main(string[] args)
        {
            try
            {
                logfilename = "DEBUG-" + dtID + "-" + dtName + ".log";
                InitializeLogger(appfolder, logfilename, isconsolelogenable);
                var supportFunc = new MyDirectoryManipulator();

                int maxWidth = Console.LargestWindowWidth;
                //int maxHeight = Console.LargestWindowHeight;
                Console.SetWindowPosition(0, 0);
                //Console.SetBufferSize(maxWidth, maxHeight);
                //Console.SetWindowSize(maxWidth, maxHeight);
                Console.BackgroundColor = ConsoleColor.DarkGray;
                Console.ForegroundColor = ConsoleColor.White;
                BlockInput(true);

                if (!Directory.Exists(appfolder))
                {
                    Directory.CreateDirectory(appfolder);
                    Directory.CreateDirectory(uploadfolder);
                    Directory.CreateDirectory(sharingfolder);
                }
                else
                {
                    supportFunc.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Zip);
                    supportFunc.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Excel);
                    supportFunc.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Log);
                }


                Console.WriteLine($" *** Accurate Desktop ver.4 Automation -  by FAIRBANC ***");
                Console.WriteLine($"");
                Console.WriteLine($"Automasi akan dimulai !");
                Console.WriteLine($"******************************************************************");
                Console.WriteLine($"             Keyboard dan Mouse akan di matikan...                ");
                Console.WriteLine($"     Komputer akan menjalankan oleh applikasi robot automasi...   ");
                Console.WriteLine($" Aktifitas penggunakan komputer akan ter-BLOKIR sekitar 10 menit. ");
                Console.WriteLine($"******************************************************************");
                Console.WriteLine($"");
                Console.BackgroundColor = ConsoleColor.Black;
                Thread.Sleep(5000);

                if (!OpenAppAndDBConfig())
                {
                    Log("Application automation failed !!");
                    return;
                }
                if (!LoginProcess())
                {
                    Log("Application automation failed !!");
                    return;
                }
                Log("Now wait for 1 minute before clicking report...");
                Thread.Sleep(35000);
                /* Try to navigare and open 'Sales' report */
                if (!OpenReport("sales"))
                {
                    Log("Application Automation failed !!");
                    return;
                }
                if (!ClosingWorkspace())
                {
                    Log("Application Automation failed !!");
                    return;
                }
                /* Try to navigare and open 'Repayment/AR and Master Outlet' report */
                if (!OpenReport("ar"))
                {
                    Log("Application Automation failed !!");
                    return;
                }
                if (!ClosingWorkspace())
                {
                    Log("Application Automation failed !!");
                    return;
                }
                /* Try to navigare and open 'Master Outlet' report */
                if (!OpenReport("outlet"))
                {
                    Log("Application Automation failed !!");
                    return;
                }
                if (!ClosingWorkspace())
                {
                    Log("Application Automation failed !!");
                    return;
                }
                if (isacctrptincluded != "Y")
                {
                    if (!CloseApp())
                    {
                        Log("Application Automation failed !!");
                        return;
                    }
                    if (automationUIA3 != null)
                    {
                        automationUIA3.Dispose();
                    }
                    ZipAndSendFile();
                    return;
                }
                /* Try to navigare and open 'Stock Valueation' report */
                if (!OpenReport("stock"))
                {
                    Log("Application Automation failed !!");
                    return;
                }
                if (!ClosingWorkspace())
                {
                    Log("Application Automation failed !!");
                    return;
                }
                /* Try to navigare and open 'Cash Flow' report */
                if (!OpenReport("cashflow"))
                {
                    Log("Application Automation failed !!");
                    return;
                }
                if (!ClosingWorkspace())
                {
                    Log("Application Automation failed !!");
                    return;
                }
                /* Try to navigare and open 'Laba/Rugi' report */
                if (!OpenReport("labarugi"))
                {
                    Log("Application Automation failed !!");
                    return;
                }
                if (!ClosingWorkspace())
                {
                    Log("Application Automation failed !!");
                    return;
                }
                /* Try to navigare and open 'Laba/Rugi' report */
                if (!OpenReport("neraca"))
                {
                    Log("Application Automation failed !!");
                    return;
                }
                if (!ClosingWorkspace())
                {
                    Log("Application Automation failed !!");
                    return;
                }
                if (!CloseApp())
                {
                    Log("Application Automation failed !!");
                    return;
                }
                if (automationUIA3 != null)
                {
                    automationUIA3.Dispose();
                }
                ZipAndSendFile();
            }
            catch (Exception ex)
            {
                Log($"Error => {ex.ToString()}");
            }
            finally
            {
                Mouse.MoveTo(10, 100);
                Console.Beep();
                Thread.Sleep(500);
                Console.Beep();
                BlockInput(false);
                Log("Accurate Desktop ver.4 Automation - SELESAI");
                if (automationUIA3 != null)
                {
                    automationUIA3.Dispose();
                }
                CloseLogger();
            }
        }

        static bool ClosingWorkspace()
        {
            try
            {
                /* Travesing back to accurate desktop main windows */
                AutomationElement mainElement = null;
                AutomationElement[] auEle = window.FindAllChildren(cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        mainElement = item;
                        break;
                    }
                }
                if (mainElement is null)
                {
                    Log($"[Step #1 Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                mainElement.SetForeground();
                Thread.Sleep(1000);

                var ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Log($"[Step #2 Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                Log("Element Interaction on property with class named -> " + ele.Properties.ClassName.ToString());
                //ele.SetForeground();
                ele.Focus();
                Thread.Sleep(1000);

                //ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Windows")));
                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Jendela")));
                if (ele is null)
                {
                    Log($"[Step #3] Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.Click();
                //System.Windows.Forms.SendKeys.SendWait("%o");
                //Log("Sending keys 'ALT+o'...");
                Thread.Sleep(1000);

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                if (ele is null)
                {
                    Log("[Step #4] Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(1);
                if (ele is null)
                {
                    Log("[Step #5] Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                Log("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.Click();
                //System.Windows.Forms.SendKeys.SendWait("A");
                //Log("Then sending key 'a'...");
                Thread.Sleep(1000);

                return true;
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return false;
            }
        }

        static bool OpenAppAndDBConfig()
        {
            try
            {
                //appx = Application.Launch(@$"{appExe}");
                //DesktopWindow = appx.GetMainWindow(automationUIA3);
                //pid = appx.ProcessId;

                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = Environment.GetEnvironmentVariable("WINDIR") + "\\explorer.exe";
                psi.Arguments = @$"{appExe}";
                Process p = Process.Start(psi);
                Thread.Sleep(25000);

                window = automationUIA3.GetDesktop();
                AutomationElement mainElement = null;
                AutomationElement[] auEle = window.FindAllDescendants(cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                if (auEle.Length > 0)
                {
                    // Get the process ID of MyUnElevatedProcess.exe
                    string targetProcessName = "Accurate";
                    Process[] processes = Process.GetProcessesByName(targetProcessName);
                    if (processes.Length > 0)
                    {
                        Process latestProcess = processes.OrderByDescending(p => p.StartTime).First();
                        pid = latestProcess.Id;
                    }
                }
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        mainElement = item;
                        break;
                    }
                }
                if (mainElement is null)
                {
                    Log($"[Step #1] Quitting, end of OpenApp automation function.");
                    return false;
                }
                Log("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                AutomationElement selamatWindowEle = null;
                var auEle1 = window.FindAllChildren(cf.ByName("Selamat Datang", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement item in auEle1)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        selamatWindowEle = item;
                    }
                }
                if (selamatWindowEle == null)
                {
                    Log($"[Step #1a] Quitting, end of OpenApp automation function.");
                    return false;
                }
                var par = selamatWindowEle.BoundingRectangle;
                Int16 cordX = Convert.ToInt16(ConfigurationManager.AppSettings["xPosWelcome"]);
                Int16 cordY = Convert.ToInt16(ConfigurationManager.AppSettings["yPosWelcome"]);
                BlockInput(false);
                Thread.Sleep(1000);
                selamatWindowEle.SetForeground();
                Mouse.MoveTo(cordX, cordY);
                Mouse.Click();
                Log($"Action -> Closing 'Welcome to Accurate' window at {(cordX)},{(cordY)}.");
                Thread.Sleep(1500);
                BlockInput(true);

                AutomationElement ele = null;
                auEle = window.FindAllDescendants(cr => cr.ByClassName("TsuiSkinMenuBar"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        ele = item;
                        break;
                    }
                }
                if (ele is null)
                {
                    Log($"[Step #2] Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log("Element Interaction on property class named -> " + ele.Properties.ClassName.ToString());
                ele.SetForeground();

                //ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("File")));
                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Berkas")));
                if (ele is null)
                {
                    Log($"[Step #3] Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.Click();
                Thread.Sleep(1000);

                // Context
                ele = null;
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                if (ele is null)
                {
                    Log("[Step #4] Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                //System.Windows.Forms.SendKeys.SendWait("%F");
                //Log("Sending keys 'ALT+F'...");
                Thread.Sleep(1000);

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(1);
                if (ele is null)
                {
                    Log("[Step #5] Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.Click();
                Thread.Sleep(1000);

                //Using opened Database window
                // ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Buka Database")));
                ele = null;
                auEle = window.FindAllDescendants(cr => cr.ByName("Buka Database"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        ele = item;
                        break;
                    }
                }
                if (ele is null)
                {
                    Log($"[Step #6 Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.SetForeground();

                if (DBaliasflag.Trim() == "Y")
                {

                    //* Opening alias DB by clicking 'Alias'' Button
                    ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Alias")));
                    if (ele is null)
                    {
                        Log("[Step #7] Quitting, end of OpenDB automation function !!");
                        return false;
                    }
                    Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                    ele.AsButton().Click();
                    Thread.Sleep(1000);

                    //* Findng alias window under desktop
                    var MainEle = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Alias")));
                    if (ele is null)
                    {
                        Log("[Step #8] Quitting, end of OpenDB automation function !!");
                        return false;
                    }
                    Log("Element Interaction on property named -> " + MainEle.Properties.Name.ToString());
                    MainEle.SetForeground();

                    //* clicking on {DBaliasname}
                    ele = WaitForElement(() => MainEle.FindFirstDescendant(cr => cr.ByName(DBaliasname)));
                    if (ele is null)
                    {
                        Log("[Step #8] Quitting, end of OpenDB automation function !!");
                        return false;
                    }
                    Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                    ele.Click();

                    //* clicking on 'Buka' button
                    ele = WaitForElement(() => MainEle.FindFirstDescendant(cr => cr.ByName("Buka")));
                    if (ele is null)
                    {
                        Log("[Step #8] Quitting, end of OpenDB automation function !!");
                        return false;
                    }
                    Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                    ele.AsButton().Click();
                    Thread.Sleep(1000);

                }
                else
                {
                    var ele2 = WaitForElement(() => ele.FindFirstChild(cr => cr.ByClassName("TEdit")));
                    if (ele2 is null)
                    {
                        Log($"[Step #7 Quitting, end of OpenDB automation function !!");
                        return false;
                    }
                    Log("Element Interaction on property class named -> " + ele2.Properties.ClassName.ToString());

                    if (ele2.AsTextBox().Text != $@"{DBpath}")
                    {
                        Thread.Sleep(1000);
                        ele2.AsTextBox().Text = $@"{DBpath}";
                    }

                    ele = ele.FindFirstChild(cf => cf.ByName("OK")).AsButton();
                    Log("Clicking 'OK' button...");
                    ele.Click();
                }
                return true;
            }
            catch
            {
                Log("Quitting, end of DB automation function !!");
                return false;
            }
        }

        static bool LoginProcess()
        {
            try
            {
                Thread.Sleep(5000);
                //var ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Login")));
                AutomationElement ele = null;
                AutomationElement[] auEle = window.FindAllDescendants(cr => cr.ByName("Daftar"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        ele = item;
                        break;
                    }
                }
                if (ele is null)
                {
                    Log($"[Step #{step += 1}] Quitting, end of login automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());

                var ele2 = ele.FindFirstDescendant(cf => cf.ByClassName("TEdit")).AsTextBox();
                ele2.Enter(loginId + "\t");
                Log("Sending Login Id...");

                System.Windows.Forms.SendKeys.SendWait(loginPassword);
                Log("Sending password...");

                ele.FindFirstDescendant(cf => cf.ByName("OK")).AsButton().Click();
                Log("Clicking 'OK' button...");
                return true;
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return false;
            }
        }

        static bool IsFileExists(string path, string fileName)
        {
            string fullPath = Path.Combine(path, fileName);
            return File.Exists(fullPath);
        }

        static bool DownloadReport(string reportName)
        {
            try
            {
                Thread.Sleep(10000);
                /** Start downloading report process **/
                /* Travesing back to accurate desktop main windows */
                AutomationElement ele1 = null;
                AutomationElement[] auEle = window.FindAllDescendants(cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        ele1 = item;
                        break;
                    }
                }
                if (ele1 is null)
                {
                    Log($"[Step #1 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele1.Properties.Name.ToString());
                Thread.Sleep(500);

                var ele = ele1.FindFirstDescendant(cf => cf.ByName("PriviewToolBar"));
                if (ele is null)
                {
                    Log($"[Step #2 Quiting, end of DownloadReport automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());

                //Export settings
                ele.FindFirstChild(cf.ByName("Expor", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)).AsButton().Click();
                Thread.Sleep(1000);

                /* The export button action resulting new window opened */
                auEle = window.FindAllChildren(cf => cf.ByName("Export to Excel"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        ele1 = item;
                        break;
                    }
                }
                if (ele1 is null)
                {
                    Log($"Step #3 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele1.Properties.Name.ToString());
                ele1.SetForeground();
                //ele1.AsButton().Click();

                /* Put here the code for iteration of report parameter check box */
                /* End of codes */

                /* Clicking OK button  */
                ele1.FindFirstChild(cf => cf.ByName("OK")).AsButton().Click();
                Log("Clicking 'OK' button...");

                if (!SavingFileDialog(reportName))
                { return false; }

                return true;
            }
            catch (Exception ex) { Log(ex.ToString()); return false; }
        }

        private static bool SavingFileDialog(string reportName)
        {
            Thread.Sleep(2000);
            AutomationElement mainEle = null;
            AutomationElement[] auEle = window.FindAllChildren(cr => cr.ByName("Save As"));
            foreach (AutomationElement item in auEle)
            {
                if (item.Properties.ProcessId == pid)
                {
                    mainEle = item;
                    break;
                }
            }
            if (mainEle is null)
            {
                Log($"Step #1 Quitting, end of OpenReport\\SavingFileDialog automation function !!");
                return false;
            }
            Log("Element Interaction on property named -> " + mainEle.Properties.Name.ToString());
            Thread.Sleep(500);

            AutomationElement ele1 = null;
            AutomationElement[] auEle1 = mainEle.FindAllDescendants(cr => cr.ByName("File name:", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
            foreach (AutomationElement item in auEle1)
            {
                if (item.Properties.ControlType.ToString() == "Edit")
                {
                    ele1 = item;
                    break;
                }
            }
            if (ele1 is null)
            {
                Log($"Step #2 Quitting, end of OpenReport\\SavingFileDialog automation function !!");
                return false;
            }
            Log("Element Interaction on property named -> " + ele1.Properties.Name.ToString());
            Thread.Sleep(500);

            var excelname = "";
            switch (reportName)
            {
                case "sales":
                    excelname = "Sales_Data";
                    break;
                case "ar":
                    excelname = "Repayment_Data";
                    break;
                case "outlet":
                    excelname = "Master_Outlet";
                    break;
                case "stock":
                    excelname = "Laporan_Stock";
                    break;
                case "labarugi":
                    excelname = "Laporan_LabaRugi";
                    break;
                case "cashflow":
                    excelname = "Laporan_ArusKas";
                    break;
                case "neraca":
                    excelname = "Laporan_NeracaSaldo";
                    break;
            }
            ele1.SetForeground();
            ele1.Focus();
            ele1.AsTextBox().Enter($"{appfolder}\\{excelname}");
            Thread.Sleep(500);

            //Save
            AutomationElement ele2 = null;
            ele2 = mainEle.FindFirstChild(cf.ByName("Save"));
            if (ele2 is null)
            {
                Log($"[Step #3 Quitting, end of DownloadReport automation function !!");
                return false;
            }
            Log("Element Interaction on property named -> " + ele2.Properties.Name.ToString());
            ele2.AsButton().Click();
            /* Pause the app to wait file saving is finnished */
            DateTime startTime = DateTime.Now;
            while (DateTime.Now - startTime < TimeSpan.FromMinutes(2))
            {
                if (IsFileExists(appfolder, excelname + ".xls"))
                {
                    //Log("File saved successfully...");
                    break;
                }
                Task.Delay(5000);
            }
            if (!IsFileExists(appfolder, excelname + ".xls"))
            {
                Console.WriteLine("'Time out' when saving file...");
            }
            return true;
        }

        private static bool SendingDate(AutomationElement ele, string date)
        {
            try
            {
                if (ele is null)
                {
                    Log($"[Step #1] Quitting, end of SendingDate automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                ele.Click();

                // Send date parameter
                ele.AsTextBox().Enter("\b\b\b\b\b\b\b\b");
                ele.AsTextBox().Text = date;

                // TWinControl
                var childEle = ele.FindFirstDescendant(cf => cf.ByClassName("TWinControl"));
                if (childEle is null)
                {
                    Log($"[Step #2] Quitting, end of OpenReport01 automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + childEle.Properties.ClassName.ToString());
                childEle.Click();
                Thread.Sleep(500);
                childEle.Click();

                Log($"Sending date parameter");

                return true;
            }
            catch (Exception ex)
            {
                Log( ex.Message);
                return false;
            }
        }

        private static bool OpenReport(string reportType)
        {
            try
            {
                AutomationElement mainElement = null;
                AutomationElement[] auEle = window.FindAllDescendants(cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        mainElement = item;
                        break;
                    }
                }
                if (mainElement is null)
                {
                    Log($"[Step #1] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                Thread.Sleep(500);

                var ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Log($"[Step #2] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                //ele.SetForeground();
                Thread.Sleep(500);

                /* Click on Reports menu */
                //ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Reports")));
                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Laporan")));
                if (ele is null)
                {
                    Log($"[Step #3] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                var pos = ele.GetClickablePoint();
                Mouse.MoveTo(pos.X, pos.Y);
                Mouse.Click();

                Thread.Sleep(1000);

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                if (ele is null)
                {
                    Log("[Step #4] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                Thread.Sleep(1000);

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(0);
                if (ele is null)
                {
                    Log("[Step #5] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.Click();
                Thread.Sleep(3000);

                //var indexToReportsElement = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Index to Reports")));
                var indexToReportsElement = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Daftar Laporan")));
                if (indexToReportsElement == null)
                {
                    Log($"[Step #6] Quitting, end of OpenReport function.");
                    return false;
                }
                Log("Element Interaction on property named -> " + indexToReportsElement.Properties.Name.ToString());
                indexToReportsElement.Focus();
                Thread.Sleep(2000);

                var reportMainTab = "";
                switch (reportType)
                {
                    case "sales":
                        reportMainTab = "Laporan Penjualan";
                        break;
                    case "ar":
                        reportMainTab = "Akun Piutang & Pelanggan";
                        break;
                    case "outlet":
                        reportMainTab = "Akun Piutang & Pelanggan";
                        break;
                    case "stock":
                        reportMainTab = "Persediaan";
                        break;
                    case "labarugi":
                        reportMainTab = "Laporan Keuangan";
                        break;
                    case "cashflow":
                        reportMainTab = "Kas & Bank";
                        break;
                    case "neraca":
                        reportMainTab = "Buku Besar";
                        break;
                }
                var reportElement1 = indexToReportsElement.FindFirstDescendant(cf.ByName(reportMainTab));

                if (reportElement1 == null)
                {
                    Log($"[Step #7] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log("Element Interaction on property named -> " + reportElement1.Properties.Name.ToString());
                reportElement1.Click();
                Thread.Sleep(1000);

                var reportName = "";
                switch (reportType)
                {
                    case "sales":
                        reportName = "Rincian Penjualan per Pelanggan";
                        break;
                    case "ar":
                        reportName = "Rincian Pembayaran Faktur";
                        break;
                    case "outlet":
                        reportName = "Daftar Pelanggan";
                        break;
                    case "stock":
                        reportName = "Ringkasan Valuasi Persediaan";
                        break;
                    case "labarugi":
                        reportName = "Laba/Rugi (Standar)";
                        break;
                    case "cashflow":
                        reportName = "Arus Kas per Akun";
                        break;
                    case "neraca":
                        reportName = "Neraca Saldo";
                        break;
                }
                var reportElement2 = indexToReportsElement.FindFirstDescendant(cf.ByName(reportName));
                if (reportElement2 == null)
                {
                    Log($"[Step #8] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log("Element Interaction on property named -> " + reportElement2.Properties.Name.ToString());
                reportElement2.SetForeground();
                reportElement2.Focus();
                reportElement2.DoubleClick();
                Thread.Sleep(10000);

                // Opening Report Format Window
                AutomationElement reportFormatElement = null;
                auEle = window.FindAllChildren(cr => cr.ByName("Report Format"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        reportFormatElement = item;
                        break;
                    }
                }
                if (reportFormatElement == null)
                {
                    Log($"[Step #9] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log("Element Interaction on property named -> " + reportFormatElement.Properties.Name.ToString());
                reportFormatElement.Focus();
                Thread.Sleep(2000);

                // Filters && Parameters => find it under 'reportFormatElement' windows tree
                var filtersAndParametersElement = reportFormatElement.FindFirstDescendant(cf.ByName("Filter && Parameter", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));

                if (filtersAndParametersElement == null)
                {
                    Log($"[Step #10] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log("Element Interaction on property named -> " + filtersAndParametersElement.Properties.Name.ToString());
                filtersAndParametersElement.Focus();
                Thread.Sleep(500);

                if (reportType == "labarugi")
                {
                    AutomationElement[] checkboxes = filtersAndParametersElement.FindAllDescendants(cr => cr.ByClassName("TCheckBox"));
                    foreach (AutomationElement checkbox in checkboxes)
                    {
                        switch (checkbox.Name)
                        {
                            case "Tampilkan Induk":
                                checkbox.AsCheckBox().IsChecked = true;
                                break;
                            case "Tampilkan Anak":
                                checkbox.AsCheckBox().IsChecked = false;
                                break;
                            case "Termasuk saldo nol":
                                checkbox.AsCheckBox().IsChecked = true;
                                break;
                            case "Tampilkan jumlah Induk":
                                checkbox.AsCheckBox().IsChecked = true;
                                break;
                            case "Tampilkan Total":
                                checkbox.AsCheckBox().IsChecked = false;
                                break;
                        }
                        Thread.Sleep(2000);
                    }
                }
                /* Sending combo box value '<Semua>' when wwhen opening Stock Valuation */
                // "stock"
                if (reportType == "stock")
                {
                    var filterBarang = filtersAndParametersElement.FindFirstDescendant(cf.ByName("<Nihil>", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                    if (filterBarang == null)
                    {
                        Log($"[Step #11] Quitting, end of OpenReport automation function.");
                        return false;
                    }
                    Log("Element Interaction on property named -> " + filterBarang.Properties.Name.ToString());
                    //filterBarang.Focus();
                    filterBarang.AsTextBox().Text = "<Semua>\r";
                    Thread.Sleep(2000);
                }

                // TabDateFromTo
                var tabDateFromToElement = filtersAndParametersElement.FindFirstDescendant(cf.ByName("TabDate", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                if (tabDateFromToElement == null)
                {
                    Log($"[Step #11] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log("Element Interaction on property named -> " + tabDateFromToElement.Properties.Name.ToString());
                tabDateFromToElement.Focus();

                /* Sending Report Date Parameters */
                AutomationElement[] dateElements = tabDateFromToElement.FindAllDescendants(cf.ByClassName("TDateEdit"));

                if (dateElements.Length > 0)
                {
                    Log($"Number of DATE parameters on screen is: {dateElements.Length}");

                    for (int index = dateElements.Length - 1; index > -1; index--)
                    {
                        var dateparam = "";
                        if (index != 0)
                        {
                            dateparam = GetFirstDate() + "/" + GetPrevMonth() + "/" + GetPrevYear();
                            //SendingDate(dateElements[index], "01/01/2000");
                        }
                        else
                        {
                            dateparam = GetLastDayOfPrevMonth() + "/" + GetPrevMonth() + "/" + GetPrevYear();
                            //SendingDate(dateElements[index], "31/12/2023");
                        }
                        SendingDate(dateElements[index], dateparam);

                    }
                }

                reportFormatElement.FindFirstDescendant(cf.ByName("OK")).AsButton().Click();
                return DownloadReport(reportType);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return false;
            }
        }

        private static bool CloseApp()
        {
            try
            {
                AutomationElement mainElement = null;
                AutomationElement[] auEle = window.FindAllDescendants(cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        mainElement = item;
                        break;
                    }
                }
                if (mainElement is null)
                {
                    Log($"[Step #1] Quitting, end of CloseApp automation function.");
                    return false;
                }
                Log("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                mainElement.SetForeground();
                Thread.Sleep(500);

                var ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Log($"[Step #2] Quitting, end of CloseApp automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                ele.SetForeground();
                Thread.Sleep(500);

                /* Click on Reports menu */
                ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Berkas")));
                if (ele is null)
                {
                    Log($"[Step #3] Quitting, end of CloseApp automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.AsMenu().Focus();
                ele.AsMenu().Click();
                Thread.Sleep(1000);

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                if (ele is null)
                {
                    Log("[Step #4] Quitting, end of CloseApp automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                Thread.Sleep(1000);

                // Exit - MenuItem #14
                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(13);
                if (ele is null)
                {
                    Log("[Step #5] Quitting, end of CloseApp automation function !!");
                    return false;
                }
                Log("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.AsMenuItem().Focus();
                Thread.Sleep(1000);
                ele.Click();
                Thread.Sleep(1000);

                /* The Menu 'File' -> 'Close' clicked action resulting new window opened */
                var ele1 = window.FindFirstDescendant(cf => cf.ByName("Confirm"));
                if (ele1 is null)
                {
                    Log($"[Step #6 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele1.Properties.Name.ToString());
                Thread.Sleep(1000);

                /* Clicking 'Yes' button  */
                ele1.FindFirstChild(cf => cf.ByName("Yes")).AsButton().Click();
                Log("Clicking 'Yes' button...");
                Thread.Sleep(5000);

                /* The Menu 'Yes' button clicked action resulting new window opened */
                ele1 = window.FindFirstDescendant(cf => cf.ByName("Confirm"));
                if (ele1 is null)
                {
                    Log($"[Step #7 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                Log("Element Interaction on property named -> " + ele1.Properties.Name.ToString());

                /* Clicking OK button  */
                ele1.FindFirstChild(cf => cf.ByName("No")).AsButton().Click();
                Log("Clicking 'No' button...");
                Thread.Sleep(1000);
                return true;
            }
            catch (Exception ex)
            {
                Log($"Exception: {ex.ToString()}");
                return false;
            }
        }

        // Zip and send files
        static void ZipAndSendFile()
        {
            try
            {
                Log("Prepare data sharing files processing...");
                var strDsPeriod = GetPrevYear() + GetPrevMonth();

                // move transaction reports excel files to Datafolder
                Log("Moving transaction excel reports file to uploaded folder...");
                var path = appfolder + @"\Master_Outlet.xls";
                var path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_OUTLET.xls";
                File.Copy(path, path2, true);
                path = appfolder + @"\Sales_Data.xls";
                path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_SALES.xls";
                File.Copy(path, path2, true);
                path = appfolder + @"\Repayment_Data.xls";
                path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_AR.xls";
                File.Copy(path, path2, true);

                // set zipping name for files
                Log("Zipping transaction file(s)");
                var strZipFile = dtID + "-" + dtName + "_" + strDsPeriod + ".zip";
                ZipFile.CreateFromDirectory(uploadfolder, sharingfolder + Path.DirectorySeparatorChar + strZipFile);

                // Send the accounting files ZIP file to the API server 
                Log("Sending transaction ZIP file to the API server...");
                var strStatusCode = "0"; // varible for debugging Curl test
                strStatusCode = SendReq(sharingfolder + Path.DirectorySeparatorChar + strZipFile, issandbox, issecurehttp);
                Thread.Sleep(5000);
                if (strStatusCode == "200")
                {
                    Log("DATA TRANSACTION SHARING - SELESAI");
                }
                else
                {
                    Log("DATA TRANSACTION SHARING - ERROR, cUrl STATUS CODE :" + strStatusCode);
                }
                var supportFunc = new MyDirectoryManipulator();
                supportFunc.DeleteFiles(uploadfolder, MyDirectoryManipulator.FileExtension.Excel);
                if (isacctrptincluded == "Y")
                {
                    // move acconting reports standart excel
                    Log("Moving standart excel reports file to uploaded folder...");
                    path = appfolder + @"\Laporan_Stock.xls";
                    path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_STOCK.xls";
                    File.Copy(path, path2, true);
                    path = appfolder + @"\Laporan_NeracaSaldo.xls";
                    path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_NERACA.xls";
                    File.Copy(path, path2, true);
                    path = appfolder + @"\Laporan_ArusKas.xls";
                    path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_ARUSKAS.xls";
                    File.Copy(path, path2, true);
                    path = appfolder + @"\Laporan_LabaRugi.xls";
                    path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_LABARUGI.xls";
                    File.Copy(path, path2, true);

                    // set zipping name for files
                    Log("Zipping accounting file(s)");
                    strZipFile = dtID + "-" + dtName + "-Financial_Statement-" + strDsPeriod + ".zip";
                    ZipFile.CreateFromDirectory(uploadfolder, sharingfolder + Path.DirectorySeparatorChar + strZipFile);

                    // Send the ZIP file to the API server 
                    Log("Sending accounting ZIP file to the API server...");
                    strStatusCode = "0"; // varible for debugging Curl test
                    strStatusCode = SendReq(sharingfolder + Path.DirectorySeparatorChar + strZipFile, issandbox, issecurehttp);
                    Thread.Sleep(5000);
                    if (strStatusCode == "200")
                    {
                        Log("DATA ACCOUNTING SHARING - SELESAI");
                    }
                    else
                    {
                        Log("DATA ACCOUNTING SHARING - ERROR, cUrl STATUS CODE :" + strStatusCode);
                    }
                }

                // Send Log file to the API server 
                path = appfolder + Path.DirectorySeparatorChar + logfilename;
                path2 = uploadfolder + Path.DirectorySeparatorChar + logfilename;
                File.Copy(path, path2, true);
                Log("Sending log file to the API server...");
                strStatusCode = SendReq(path2, issandbox, issecurehttp);
                Thread.Sleep(5000);
                supportFunc.DeleteFiles(uploadfolder, MyDirectoryManipulator.FileExtension.Excel);
                supportFunc.DeleteFiles(uploadfolder, MyDirectoryManipulator.FileExtension.Log);
                supportFunc.CopyFolderFiles(appfolder, uploadfolder);
            }
            catch (Exception ex)
            {
                // Log any exceptions that occur during file operations
                Log($"Error during ZIP and send: {ex.Message}");
                //throw ex;
            }
        }

        private static string SendReq(string strFileDataInfo, string strSandboxBool, string isSecureHTTP)
        {
            try
            {
                string text = "";
                string text2 = "";
                if (strSandboxBool == "Y")
                {
                    text2 = "KQtbMk32csiJvm8XDAx2KnRAdbtP3YVAnJpF8R5cb2bcBr8boT3dTvGc23c6fqk2NknbxpdarsdF3M4V";
                    text = ((!(isSecureHTTP == "Y")) ? "http://sandbox.fairbanc.app/api/documents" : "https://sandbox.fairbanc.app/api/documents");
                }
                else
                {
                    text2 = "2S0VtpYzETxDrL6WClmxXXnOcCkNbR5nUCCLak6EHmbPbSSsJiTFTPNZrXKk2S0VtpYzETxDrL6WClmx";
                    text = ((!(isSecureHTTP == "Y")) ? "http://dashboard.fairbanc.app/api/documents" : "https://dashboard.fairbanc.app/api/documents");
                }

                Log("Preparing to send a request to the API server...");
                HttpClient httpClient = new HttpClient();
                HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, text);
                MultipartFormDataContent multipartFormDataContent = new MultipartFormDataContent();
                multipartFormDataContent.Add(new StringContent(text2), "api_token");
                multipartFormDataContent.Add(new ByteArrayContent(File.ReadAllBytes(strFileDataInfo)), "file", Path.GetFileName(strFileDataInfo));
                httpRequestMessage.Content = multipartFormDataContent;
                HttpResponseMessage httpResponseMessage = httpClient.Send(httpRequestMessage);
                Thread.Sleep(5000);
                httpResponseMessage.EnsureSuccessStatusCode();
                var strResponseBody = httpResponseMessage.ToString();
                string[] array = strResponseBody.Split(':', ',');
                Log($"Response from API server: {array[1].Trim()}");
                return array[1].Trim();
            }
            catch (Exception ex)
            {
                // Log any exceptions that occur during the API request
                Log($"Error during API request: {ex.Message}");
                return "-1";
            }
        }

        static string GetPrevMonth()
        {
            return DateTime.Now.AddMonths(-1).ToString("MM");
        }

        static string GetPrevYear()
        {
            return DateTime.Now.AddMonths(-1).ToString("yyyy");
        }

        static string GetDSPeriod()
        {
            return GetPrevYear() + GetPrevMonth();
        }

        static string GetFirstDate()
        {
            return "01";
        }

        static string GetLastDayOfPrevMonth()
        {
            var lastDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1);
            return lastDay.ToString("dd");
        }




    }

}

