using System;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.Conditions;
using FlaUI.Core.AutomationElements;
using Serilog;
using System.Configuration;
using System.Runtime.InteropServices;
using System.IO.Compression;
using System.Diagnostics.Eventing.Reader;
using Windows.ApplicationModel;
using FlaUI.Core.Input;
using System.Drawing.Printing;

namespace DesktopDSPTTest // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static Application appx;
        static Window DesktopWindow;
        static UIA3Automation automationUIA3 = new UIA3Automation();
        static ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());
        static AutomationElement window = automationUIA3.GetDesktop();
        static int step = 0;
        static string dtID = ConfigurationManager.AppSettings["dtID"];
        static string dtName = ConfigurationManager.AppSettings["dtName"];
        static string appExe = ConfigurationManager.AppSettings["erpappnamepath"];
        static string LoginId = ConfigurationManager.AppSettings["loginId"];
        static string LoginPassword = ConfigurationManager.AppSettings["password"];
        static string enableconsolelog = ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
        static string issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
        static string DBpath = ConfigurationManager.AppSettings["DBaddresspath"].ToUpper();
        static string DBaliasflag = ConfigurationManager.AppSettings["DBaliasflag"].ToUpper();
        static string DBaliasname = ConfigurationManager.AppSettings["DBaliasname"];
        static string appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
        static string uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
        static string sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
        //static string screenshotfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotfolder"];
        static string logfilename = "";

        const int EXCELfile = 0;
        const int LOGFile = 1;
        const int ZIPFile = 2;
        const UInt32 WM_CLOSE = 0x0010;

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

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
        static void closeWarningDialogBox(string Title)
        {
            var hWnd = FindWindow(null, Title);
            if (hWnd != IntPtr.Zero)
            {
                SendMessage(hWnd, WM_CLOSE, 0, 0);
                Log.Information("Closing dialog message box.");
            }
        }

        static void Main(string[] args)
        {
            try
            {
                DeleteSupportingFiles(appfolder, EXCELfile);
                DeleteSupportingFiles(appfolder, LOGFile);
                DeleteSupportingFiles(appfolder, ZIPFile);

                if (!Directory.Exists(appfolder))
                {
                    Directory.CreateDirectory(appfolder);
                    Directory.CreateDirectory(uploadfolder);
                    Directory.CreateDirectory(sharingfolder);
                }
                var config = new LoggerConfiguration();
                if (enableconsolelog == "Y")
                {
                    config.WriteTo.Console();
                }
                logfilename = "DEBUG-" + dtID + "-" + dtName + ".log";
                config.WriteTo.File(appfolder + Path.DirectorySeparatorChar + logfilename);
                Log.Logger = config.CreateLogger();

                Log.Information("*** Accurate Desktop ver.4 Automation -  by FAIRBANC ***");
                Console.WriteLine("");
                Log.Information("Automasi akan dimulai !");
                Console.WriteLine("******************************************************************");
                Console.WriteLine("Keyboard dan Mouse akan di matikan...");
                Console.WriteLine("Komputer akan di jalankan oleh applikasi robot automasi...");
                Console.WriteLine("Aktifitas penggunakan komputer akan terblokir selama 10 menit...");
                Console.WriteLine("******************************************************************");
                Thread.Sleep(10000);

                if (!OpenAppAndDBConfig())
                {
                    Log.Information("application automation failed !!");
                    return;
                }
                if (!LoginProcess())
                {
                    Log.Information("application automation failed !!");
                    return;
                }
                Log.Information("now wait for 1 minute before clicking report...");
                Thread.Sleep(35000);


                /* Try to navigare and open 'Sales' report */
                if (!OpenReport("sales"))
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                /* Download opened report on screen */
                if (!DownloadReport("sales"))
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                /* Closing Workspaces that contain all report tab */
                if (!ClosingWorkspace())
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                /* Try to navigare and open 'Sales' report */
                if (!OpenReport("ar"))
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                /* Download opened report on screen */
                if (!DownloadReport("ar"))
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                /* Closing Workspaces that contain all report tab */
                if (!ClosingWorkspace())
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                if (!CloseApp())
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                ZipAndSendFile();
            }
            catch (Exception ex)
            {
                Log.Information($"Error => {ex.ToString()}");
            }
            finally
            {
                Log.Information("Accurate Desktop ver.4 Automation - SELESAI");
                if (automationUIA3 != null)
                {
                    automationUIA3.Dispose();
                }
                Log.CloseAndFlush();

                //Task.Run(() => System.Windows.MessageBox.Show("Tejadi kesalahan, sliahkan tutup applikasi !!!"));
            }
        }

        static bool ClosingWorkspace()
        {
            try
            {
                /* Travesing back to accurate desktop main windows */
                var eleMain = window.FindFirstDescendant(cf => cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                if (eleMain is null)
                {
                    Log.Information($"[Step #1 Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + eleMain.Properties.Name.ToString());
                eleMain.SetForeground();
                Thread.Sleep(1000);

                var ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Log.Information($"[Step #2 Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property with class named -> " + ele.Properties.ClassName.ToString());
                //ele.SetForeground();
                ele.Focus();
                Thread.Sleep(1000);

                //ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Windows")));
                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Jendela")));
                if (ele is null)
                {
                    Log.Information($"[Step #3] Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.Click();
                //System.Windows.Forms.SendKeys.SendWait("%o");
                //Log.Information("Sending keys 'ALT+o'...");
                Thread.Sleep(1000);

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                if (ele is null)
                {
                    Log.Information("[Step #4] Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(1);
                if (ele is null)
                {
                    Log.Information("[Step #5] Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.Click();
                //System.Windows.Forms.SendKeys.SendWait("A");
                //Log.Information("Then sending key 'a'...");
                Thread.Sleep(1000);

                return true;
            }
            catch (Exception ex)
            {
                Log.Information(ex.ToString());
                return false;
            }
        }

        static bool OpenAppAndDBConfig()
        {
            try
            {
                appx = Application.Launch(@$"{appExe}");
                DesktopWindow = appx.GetMainWindow(automationUIA3);

                // Wait until Accurate window ready
                // FlaUI.Core.Input.Wait.UntilResponsive(DesktopWindow.FindFirstChild(), TimeSpan.FromMilliseconds(5000));
                Thread.Sleep(5000);
                closeWarningDialogBox("Selamat Datang di");
                Log.Information("Action -> Closing 'Welcome to Accurate' window.");

                Thread.Sleep(1500);

                /* Closing Warning diaLog box */
                var ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Warning")));
                if (ele is null)
                {
                    Log.Information($"[Step #1] Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property class named -> " + ele.Properties.ClassName.ToString());

                ele.FindFirstChild(cf => cf.ByName("OK")).AsButton().Click();
                Log.Information("Clicking 'OK' button...");

                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Log.Information($"[Step #2] Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property class named -> " + ele.Properties.ClassName.ToString());
                ele.SetForeground();

                //ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("File")));
                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Berkas")));
                if (ele is null)
                {
                    Log.Information($"[Step #3] Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.Click();

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                if (ele is null)
                {
                    Log.Information("[Step #4] Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                //System.Windows.Forms.SendKeys.SendWait("%F");
                //Log.Information("Sending keys 'ALT+F'...");
                Thread.Sleep(1000);

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(1);
                if (ele is null)
                {
                    Log.Information("[Step #5] Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.Click();
                //System.Windows.Forms.SendKeys.SendWait("o");
                //Log.Information("Then sending key 'o'...");
                Thread.Sleep(1000);

                //Open Database
                //ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Open Database")));
                ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Buka Database")));
                if (ele is null)
                {
                    Log.Information($"[Step #6 Quitting, end of OpenDB automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.SetForeground();

                if (DBaliasflag.Trim() == "Y")
                {

                    //* Opening alias DB by clicking 'Alias'' Button
                    ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Alias")));
                    if (ele is null)
                    {
                        Log.Information("[Step #7] Quitting, end of OpenDB automation function !!");
                        return false;
                    }
                    Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                    ele.AsButton().Click();
                    Thread.Sleep(1000);

                    //* Findng alias window under desktop
                    ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Alias")));
                    if (ele is null)
                    {
                        Log.Information("[Step #8] Quitting, end of OpenDB automation function !!");
                        return false;
                    }
                    Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                    ele.SetForeground();

                    //* clicking on {DBaliasname}
                    ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName(DBaliasname)));
                    if (ele is null)
                    {
                        Log.Information("[Step #8] Quitting, end of OpenDB automation function !!");
                        return false;
                    }
                    Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                    ele.Click();

                    //* clicking on 'Buka' button
                    ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Buka")));
                    if (ele is null)
                    {
                        Log.Information("[Step #8] Quitting, end of OpenDB automation function !!");
                        return false;
                    }
                    Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                    ele.AsButton().Click();
                    Thread.Sleep(1000);

                }
                else
                {
                    var ele2 = WaitForElement(() => ele.FindFirstChild(cr => cr.ByClassName("TEdit")));
                    if (ele2 is null)
                    {
                        Log.Information($"[Step #7 Quitting, end of OpenDB automation function !!");
                        return false;
                    }
                    Log.Information("Element Interaction on property class named -> " + ele2.Properties.ClassName.ToString());

                    if (ele2.AsTextBox().Text != $@"{DBpath}")
                    {
                        Thread.Sleep(1000);
                        ele2.AsTextBox().Text = $@"{DBpath}";
                    }

                    ele = ele.FindFirstChild(cf => cf.ByName("OK")).AsButton();
                    Log.Information("Clicking 'OK' button...");
                    ele.Click();
                }


                return true;
            }
            catch
            {
                Log.Information("Quitting, end of DB automation function !!");
                return false;
            }
        }

        static bool LoginProcess()
        {
            try
            {
                //var ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Login")));
                var ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Daftar")));
                if (ele is null)
                {
                    Log.Information($"[Step #{step += 1}] Quitting, end of login automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());

                var ele2 = ele.FindFirstDescendant(cf => cf.ByClassName("TEdit")).AsTextBox();
                ele2.Enter(LoginId + "\t");
                Log.Information("Sending Login Id...");

                System.Windows.Forms.SendKeys.SendWait(LoginPassword);
                Log.Information("Sending password...");

                ele.FindFirstDescendant(cf => cf.ByName("OK")).AsButton().Click();
                Log.Information("Clicking 'OK' button...");
                return true;
            }
            catch (Exception ex)
            {
                Log.Information(ex.ToString());
                return false;
            }
        }

        static bool OpenReport01(string rptType)
        {
            try
            {
                var ele1 = WaitForElement(() => window.FindFirstDescendant(cf => cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)));
                if (ele1 is null)
                {
                    Log.Information($"[Step #1 Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele1.Properties.ClassName.ToString());
                ele1.SetForeground();
                Thread.Sleep(500);

                var ele = ele1.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar"));
                if (ele is null)
                {
                    Log.Information($"[Step #2] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Log.Information(ele.Properties.ClassName.ToString());
                //ele.Focus();

                System.Windows.Forms.SendKeys.SendWait("%R");
                Log.Information("Sending keys 'ALT+R'...");
                Thread.Sleep(1000);

                System.Windows.Forms.SendKeys.SendWait("i");
                Log.Information("Then sending key 'I'...");
                Thread.Sleep(1000);

                /* Travesing to 'Index to Reports' */
                //window = automationUIA3.GetDesktop();
                var eleParent = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Index to Reports")));
                if (eleParent is null)
                {
                    Log.Information($"[Step #3] Quitting, end of OpenReport0\\{rptType}1 automation function !!");
                    return false;
                }
                Log.Information("Class name is {0}", eleParent.Properties.ClassName.ToString());
                Thread.Sleep(1000);

                if (rptType == "sales")
                {
                    //Sales Reports
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Sales Reports"));
                    if (ele is null)
                    {
                        Log.Information($"[Step #4] Quitting, end of OpenReport01\\{rptType} automation function !!");
                        return false;
                    }
                    ele.Click();
                    Thread.Sleep(500);

                    //Sales By Customer Detail
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Sales By Customer Detail"));
                    if (ele is null)
                    {
                        Log.Information($"[Step #5] Quitting, end of OpenReport01\\{rptType} automation function !!"); ;
                        return false;
                    }
                    ele.DoubleClick();
                    Thread.Sleep(3000);
                }

                if (rptType == "ar")
                {
                    //Account Receivables & Customers
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Account Receivables & Customers"));
                    if (ele is null)
                    {
                        Log.Information($"[Step #9] Quitting, end of OpenReport01\\{rptType} automation function !!");
                        return false;
                    }
                    ele.Click();
                    Thread.Sleep(500);

                    //Invoices Paid Summary
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Invoices Paid Summary"));
                    if (ele is null)
                    {
                        Log.Information($"[Step #10] Quitting, end of OpenReport01\\{rptType} automation function !!"); ;
                        return false;
                    }
                    ele.DoubleClick();
                    Thread.Sleep(3000);
                }

                //Report Format
                eleParent = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Report Format")).AsWindow());
                if (eleParent is null)
                {
                    Log.Information($"[Step #6] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Log.Information(eleParent.Properties.ClassName.ToString());
                eleParent.Focus();
                Thread.Sleep(2000);

                //  Filters && Parameters => Under 'Desktop' windows
                ele = eleParent.FindFirstDescendant(cf => cf.ByName("Filters && Parameters", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                //ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Filters && Parameters", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)));
                if (ele is null)
                {
                    Log.Information($"[Step #7] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Log.Information(ele.Properties.ClassName.ToString());
                ele.Focus();
                Thread.Sleep(500);

                //TabDateFromTo
                ele = ele.FindFirstDescendant(cr => cf.ByName("TabDateFromTo"));
                if (ele is null)
                {
                    Log.Information($"[Step #8] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                ele.Focus();

                /* Sending Report Date Parameters */
                AutomationElement[] ArrEle = ele.FindAllDescendants(cf => cf.ByClassName("TDateEdit"));
                if (ArrEle.Length > 0)
                {
                    //TDateEdit
                    Log.Information($"Number of DATE parameter on screen is : {ArrEle.Length}");
                    for (int index = ArrEle.Length - 1; index > -1; index--)
                    {
                        if (index != 0)
                        {
                            SendingDate(ArrEle[index], "01/01/2000");
                        }
                        else
                        {
                            SendingDate(ArrEle[index], "31/12/2023");
                        }
                    }
                }

                eleParent.FindFirstDescendant(cf.ByName("OK")).AsButton().Click();

                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"[Exception] Quitting, end of OpenReport01\\{rptType} automation function !!");
                throw ex;
            }

        }

        static bool DownloadReport(string reportName)
        {
            try
            {
                /** Start downloading report process **/
                /* Travesing back to accurate desktop main windows */
                var ele1 = window.FindFirstDescendant(cf => cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                if (ele1 is null)
                {
                    Log.Information($"[Step #1 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele1.Properties.Name.ToString());
                Thread.Sleep(500);

                var ele = ele1.FindFirstDescendant(cf => cf.ByName("PriviewToolBar"));
                if (ele1 is null)
                {
                    Log.Information($"[Step #2 Quiting, end of DownloadReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());

                //Export settings
                ele.FindFirstChild(cf.ByName("Expor", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)).AsButton().Click();
                Thread.Sleep(1000);

                /* The export button action resulting new window opened */
                ele1 = window.FindFirstDescendant(cf => cf.ByName("Export to Excel"));
                if (ele1 is null)
                {
                    Log.Information($"[Step #3 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele1.Properties.Name.ToString());

                /* Put here the code for iteration of report parameter check box */
                /* End of codes */

                /* Clicking OK button  */
                ele1.FindFirstChild(cf => cf.ByName("OK")).AsButton().Click();
                Log.Information("Clicking 'OK' button...");

                if (!SavingFileDialog(reportName))
                { return false; }

                return true;
            }
            catch (Exception ex) { Log.Information(ex.ToString()); return false; }
        }

        private static bool SavingFileDialog(string reportName)
        {
            //var ele1 = window.FindFirstDescendant(cf.ByClassName("#32770"));
            //if (ele1 is null)
            //{
            //    log.Information($"[Step #1] Quitting, end of OpenReport\\SavingFileDialog automation function !!");
            //    return false;
            //}
            //ele1.Focus();

            //Edit
            System.Windows.Forms.SendKeys.SendWait("%n");
            Log.Information("Saving file by sending keys 'ALT+n'...");
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
                default:
                    break;
            }
            System.Windows.Forms.SendKeys.SendWait($@"{appfolder}\{excelname}.");
            Thread.Sleep(500);

            //Save
            var ele = window.FindFirstDescendant(cf => cf.ByName("Save"));
            if (ele is null)
            {
                Log.Information($"[Step #2] Quitting, end of OpenReport\\SavingFileDialog automation function !!");
                return false;
            }
            Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
            ele.SetForeground();
            ele.AsButton().Click();

            /* Pause the app to wait file saving is finnished */
            Thread.Sleep(10000);

            return true;
        }

        private static bool SendingDate(AutomationElement ele, string date)
        {
            try
            {
                if (ele is null)
                {
                    Log.Information($"[Step #1] Quitting, end of SendingDate automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                ele.Click();

                // Send date parameter
                ele.AsTextBox().Enter("\b\b\b\b\b\b\b\b");
                ele.AsTextBox().Text = date;

                // TWinControl
                var childEle = ele.FindFirstDescendant(cf => cf.ByClassName("TWinControl"));
                if (childEle is null)
                {
                    Log.Information($"[Step #2] Quitting, end of OpenReport01 automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + childEle.Properties.ClassName.ToString());
                childEle.Click();
                Thread.Sleep(500);
                childEle.Click();

                Log.Information($"Sending date parameter");

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void DeleteSupportingFiles(string FolderName, int fileExtEnum)
        {
            try
            {
                // Delete Excel files
                var fileextpattern = "";
                switch (fileExtEnum)
                {
                    case 0:
                        fileextpattern = "*.xl*";
                        break;
                    case 1:
                        fileextpattern = "*.log";
                        break;
                    case 2:
                        fileextpattern = "*.zip";
                        break;
                }
                var supportFiles = Directory.EnumerateFiles(FolderName, fileextpattern, SearchOption.AllDirectories);
                foreach (var excelFile in supportFiles)
                {
                    File.Delete(excelFile);
                    Console.WriteLine($"Deleted Excel file: {excelFile}");
                }
            }
            catch (Exception)
            { throw; }
        }

        private static bool OpenReport(string reportType)
        {
            try
            {
                var mainElement = WaitForElement(() => window.FindFirstDescendant(cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)));
                if (mainElement is null)
                {
                    Log.Information($"[Step #1] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                mainElement.SetForeground();
                Thread.Sleep(500);

                var ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Log.Information($"[Step #2] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                //ele.SetForeground();
                Thread.Sleep(500);

                /* Click on Reports menu */
                //ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Reports")));
                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Laporan")));
                if (ele is null)
                {
                    Log.Information($"[Step #3] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                //ele.AsMenu().Focus();
                //ele.AsMenu().Click();
                var pos = ele.GetClickablePoint();
                Mouse.MoveTo(pos.X, pos.Y);
                Mouse.Click();

                Thread.Sleep(1000);

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                if (ele is null)
                {
                    Log.Information("[Step #4] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                //System.Windows.Forms.SendKeys.SendWait("%R");
                //Log.Information("Sending keys 'ALT+R'...");
                Thread.Sleep(1000);

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(0);
                if (ele is null)
                {
                    Log.Information("[Step #5] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.Click();
                //System.Windows.Forms.SendKeys.SendWait("i");
                //Log.Information("Then sending key 'I'...");
                Thread.Sleep(3000);

                //var indexToReportsElement = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Index to Reports")));
                var indexToReportsElement = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Daftar Laporan")));
                if (indexToReportsElement == null)
                {
                    Log.Information($"[Step #6] Quitting, end of OpenReport function.");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + indexToReportsElement.Properties.Name.ToString());
                indexToReportsElement.Focus();
                Thread.Sleep(2000);

                //* var reportMainTab = (reportType == "sales") ? "Sales Reports" : "Account Receivables & Customers";
                var reportMainTab = (reportType == "sales") ? "Laporan Penjualan" : "Akun Piutang & Pelanggan";
                var reportElement1 = indexToReportsElement.FindFirstDescendant(cf.ByName(reportMainTab));

                if (reportElement1 == null)
                {
                    Log.Information($"[Step #7] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + reportElement1.Properties.Name.ToString());
                reportElement1.Click();
                Thread.Sleep(1000);

                //var reportName = (reportType == "sales") ? "Sales By Customer Detail" : "Invoices Paid Summary";
                var reportName = (reportType == "sales") ? "Rincian Penjualan per Pelanggan" : "Rincian Pembayaran Faktur";
                var reportElement2 = indexToReportsElement.FindFirstDescendant(cf.ByName(reportName));

                if (reportElement2 == null)
                {
                    Log.Information($"[Step #8] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + reportElement2.Properties.Name.ToString());
                reportElement2.DoubleClick();
                Thread.Sleep(3000);

                // Report Format
                var reportFormatElement = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Report Format")).AsWindow());

                if (reportFormatElement == null)
                {
                    Log.Information($"[Step #9] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + reportFormatElement.Properties.Name.ToString());
                reportFormatElement.Focus();
                Thread.Sleep(2000);

                // Filters && Parameters => Under 'Desktop' windows
                var filtersAndParametersElement = reportFormatElement.FindFirstDescendant(cf.ByName(" Filter && Parameter", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));

                if (filtersAndParametersElement == null)
                {
                    Log.Information($"[Step #10] Quitting, end of OpenReport automation10unction.");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + filtersAndParametersElement.Properties.Name.ToString());
                filtersAndParametersElement.Focus();
                Thread.Sleep(500);

                // TabDateFromTo
                var tabDateFromToElement = filtersAndParametersElement.FindFirstDescendant(cf.ByName("TabDateFromTo"));
                if (tabDateFromToElement == null)
                {
                    Log.Information($"[Step #11] Quitting, end of OpenReport aut[Step #11] ction.");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + tabDateFromToElement.Properties.Name.ToString());
                tabDateFromToElement.Focus();

                /* Sending Report Date Parameters */
                AutomationElement[] dateElements = tabDateFromToElement.FindAllDescendants(cf.ByClassName("TDateEdit"));

                if (dateElements.Length > 0)
                {
                    Log.Information($"Number of DATE parameters on screen is: {dateElements.Length}");

                    for (int index = dateElements.Length - 1; index > -1; index--)
                    {
                        if (index != 0)
                        {
                            var dateparam = GetFirstDate() + "/" + GetPrevMonth() + "/" + GetPrevYear();
                            SendingDate(dateElements[index], "01/01/2000");
                        }
                        else
                        {
                            var dateparam = GetLastDayOfPrevMonth() + "/" + GetPrevMonth() + "/" + GetPrevYear();
                            SendingDate(dateElements[index], "31/12/2023");
                        }
                    }
                }

                reportFormatElement.FindFirstDescendant(cf.ByName("OK")).AsButton().Click();
                return true;
            }
            catch (Exception ex)
            {
                Log.Information(ex.ToString());
                return false;
            }
        }

        private static bool CloseApp()
        {
            try
            {
                var mainElement = WaitForElement(() => window.FindFirstDescendant(cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)));
                if (mainElement is null)
                {
                    Log.Information($"[Step #1] Quitting, end of CloseApp automation function.");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                mainElement.SetForeground();
                Thread.Sleep(500);

                var ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Log.Information($"[Step #2] Quitting, end of CloseApp automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                ele.SetForeground();
                Thread.Sleep(500);

                /* Click on Reports menu */
                ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Berkas")));
                if (ele is null)
                {
                    Log.Information($"[Step #3] Quitting, end of CloseApp automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.AsMenu().Focus();
                ele.AsMenu().Click();
                Thread.Sleep(1000);

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                if (ele is null)
                {
                    Log.Information("[Step #4] Quitting, end of CloseApp automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                Thread.Sleep(1000);

                // Exit - MenuItem #14
                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(13);
                if (ele is null)
                {
                    Log.Information("[Step #5] Quitting, end of CloseApp automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.AsMenuItem().Focus();
                Thread.Sleep(1000);
                ele.Click();
                Thread.Sleep(1000);

                /* The Menu 'File' -> 'Close' clicked action resulting new window opened */
                var ele1 = window.FindFirstDescendant(cf => cf.ByName("Confirm"));
                if (ele1 is null)
                {
                    Log.Information($"[Step #6 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele1.Properties.Name.ToString());
                Thread.Sleep(1000);

                /* Clicking 'Yes' button  */
                ele1.FindFirstChild(cf => cf.ByName("Yes")).AsButton().Click();
                Log.Information("Clicking 'Yes' button...");
                Thread.Sleep(5000);

                /* The Menu 'Yes' button clicked action resulting new window opened */
                ele1 = window.FindFirstDescendant(cf => cf.ByName("Confirm"));
                if (ele1 is null)
                {
                    Log.Information($"[Step #7 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele1.Properties.Name.ToString());

                /* Clicking OK button  */
                ele1.FindFirstChild(cf => cf.ByName("No")).AsButton().Click();
                Log.Information("Clicking 'No' button...");
                Thread.Sleep(1000);
                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Exception: {ex.ToString()}");
                return false;
            }
        }

        // Zip and send files
        static void ZipAndSendFile()
        {
            try
            {
                Log.Information("Checking data sharing files...");
                //CheckAndDeleteZipFile(strDataSharingFolder);
                var strDsPeriod = GetPrevYear() + GetPrevMonth();

                Log.Information("Moving standart excel reports file to uploaded folder...");
                // move excels files to Datafolder

                /* var path = strAppDownloadFolder + @"\Master_Outlet.xlsx";
                var path2 = strUploadFolder + @"\ds-" + strDTid + "-" + strDTname + "-" + strDsPeriod + "_OUTLET.xlsx";
                File.Move(path, path2, true);
                */

                var path = appfolder + @"\Sales_Data.xls";
                var path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_SALES.xls";
                File.Move(path, path2, true);
                path = appfolder + @"\Repayment_Data.xls";
                path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_AR.xls";
                File.Move(path, path2, true);

                // set zipping name for files
                Log.Information("Zipping Transaction file(s)");
                var strZipFile = dtID + "-" + dtName + "_" + strDsPeriod + ".zip";
                ZipFile.CreateFromDirectory(uploadfolder, sharingfolder + Path.DirectorySeparatorChar + strZipFile);

                // Send the ZIP file to the API server 
                Log.Information("Sending ZIP file to the API server...");
                var strStatusCode = "0"; // varible for debugging Curl test
                strStatusCode = SendReq(sharingfolder + Path.DirectorySeparatorChar + strZipFile, issandbox, "Y");
                Thread.Sleep(5000);
                if (strStatusCode == "200")
                {
                    Log.Information("DATA TRANSACTION SHARING - SELESAI");
                }
                else
                {
                    Log.Information("DATA TRANSACTION SHARING - ERROR, cUrl STATUS CODE :" + strStatusCode);
                }

                // Send Log file to the API server 
                path = appfolder + Path.DirectorySeparatorChar + logfilename;
                path2 = uploadfolder + Path.DirectorySeparatorChar + logfilename;
                Log.Information("Sending log file to the API server...");
                File.Copy(path, path2, true);
                strStatusCode = SendReq(path2, issandbox, "Y");
                Thread.Sleep(5000);
            }
            catch (Exception ex)
            {
                // Log any exceptions that occur during file operations
                Log.Information($"Error during ZIP and send: {ex.Message}");
                throw ex;
            }
        }

        private static string SendReq(string strFileDataInfo, string strSandboxBool, string SecureHTTP)
        {
            try
            {
                string text = "";
                string text2 = "";
                if (strSandboxBool == "Y")
                {
                    text2 = "KQtbMk32csiJvm8XDAx2KnRAdbtP3YVAnJpF8R5cb2bcBr8boT3dTvGc23c6fqk2NknbxpdarsdF3M4V";
                    text = ((!(SecureHTTP == "Y")) ? "http://sandbox.fairbanc.app/api/documents" : "https://sandbox.fairbanc.app/api/documents");
                }
                else
                {
                    text2 = "2S0VtpYzETxDrL6WClmxXXnOcCkNbR5nUCCLak6EHmbPbSSsJiTFTPNZrXKk2S0VtpYzETxDrL6WClmx";
                    text = ((!(SecureHTTP == "Y")) ? "http://dashboard.fairbanc.app/api/documents" : "https://dashboard.fairbanc.app/api/documents");
                }

                Log.Information("Preparing to send a request to the API server...");
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
                Log.Information($"Response from API server: {array[1].Trim()}");
                return array[1].Trim();
            }
            catch (Exception ex)
            {
                // Log any exceptions that occur during the API request
                Log.Information($"Error during API request: {ex.Message}");
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

        static void CheckAndDeleteZipFile(string folder)
        {
            try
            {
                var pattern = "*.zip";
                var zipFiles = Directory.EnumerateFiles(folder, pattern);
                foreach (var zipFile in zipFiles)
                {
                    File.Delete(zipFile);
                    Log.Information($"Deleted ZIP file(s): {zipFile}");
                }
            }
            catch (Exception ex)
            {
                Log.Information($"Error during ZIP file deletion: {ex.Message}");
            }
        }

    }

}

