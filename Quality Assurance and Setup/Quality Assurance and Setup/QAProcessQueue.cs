using IWshRuntimeLibrary;
using Shell32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quality_Assurance_and_Setup {
    public class QAProcessQueue {
        private List<QAProcess> ProcessQueue = new List<QAProcess>();
        public bool X64 { get;}
        public int OfficeVersion { get; }
        public QAType TypeOfQA { get;}
        public bool FullQueue { get; }
        
        public QAProcessQueue(bool is64Bit, int version, QAType QAtoPerform, bool initializeFullQueue = true) {
            X64 = is64Bit;
            OfficeVersion = version;
            TypeOfQA = QAtoPerform;
            FullQueue = initializeFullQueue;
            if (TypeOfQA == QAType.Customer && FullQueue)
                InitializeFullCustomerQueue();
            else if ((TypeOfQA == QAType.Loaner || TypeOfQA == QAType.Kiosk) && FullQueue)
                InitializeFullOtherQueue();
            else {
                //InitializeQueue();
            }
        }

        private void InitializeFullCustomerQueue() {
            ProcessQueue.Add(ShowVPN());
            
            ProcessQueue.Add(RunTagIT());

            ProcessQueue.Add(RunLync());
            
            ProcessQueue.Add(RepairIESettings());
            
            ProcessQueue.Add(OpenChannelViewer());

            ProcessQueue.Add(OpenMPSPortal());

            ProcessQueue.Add(OpenCertificateManager());
            
            ProcessQueue.Add(OpenExcel());
            
            ProcessQueue.Add(OpenOutlook());

            ProcessQueue.Add(RunTuneUp());
            
            ProcessQueue.Add(OpenPrinterDoc());
            
            ProcessQueue.Add(OpenDocumentsFolder());
        }

        private void InitializeFullOtherQueue() {
            ProcessQueue.Add(RunTagIT());
            
            ProcessQueue.Add(OpenChannelViewer());

            ProcessQueue.Add(OpenExcel());

            ProcessQueue.Add(RunTuneUp());

            ProcessQueue.Add(RemoveOutlook());
        }

        public void ExecuteQueue(MainWindow siht) {
            foreach (QAProcess QAp in ProcessQueue) {
                siht.PrintLine(QAp.Run());
            }

            if (TypeOfQA == QAType.Customer) {
                CompleteCustomerQueue(siht);
            }else if (TypeOfQA == QAType.Loaner) {
                CompleteLoanerQueue(siht);
            }else {
                CompleteKioskQueue(siht);
            }
            
            siht.PrintLine("QA Process COMPLETED!");
        }

        private void CompleteCustomerQueue(MainWindow siht) {
            AddVPNLinkToDesktop(siht);

            PinLyncIconToTaskbar(siht);
        }

        private void CompleteLoanerQueue(MainWindow siht) {
            WebmailLink(siht);

            AddVPNLinkToPublicDesktop(siht);

            OfficeLink(siht);
        }

        private void CompleteKioskQueue(MainWindow siht) {
            WebmailLink(siht);

            OfficeLink(siht);

            ITAccessLink(siht);

            ITSolutionsLink(siht);

            GuestWiFiLink(siht);

            EformsLink(siht);
        }

        private QAProcess ShowVPN() {
            QAProcess process = new QAProcess("Start VPN",
                                              "Starting the VPN software for testing purposes");
            if (!X64) {
                //32-bit specific
                process.App = @"C:\Program Files\Common Files\Juniper Networks\JamUI\Pulse.exe";
                process.Arguments = "-show";
            } else {
                //64-bit specific
                process.App = @"C:\Program Files (x86)\Common Files\Juniper Networks\JamUI\Pulse.exe";
                process.Arguments = "-show";
            }
            return process;
        }

        private QAProcess RunTagIT() {
            //Run TagIT to set the asset tag #
            return new QAProcess("Start TagIT",
                                 "Running TagIT to set the asset tag #",
                                 @"C:\Program Files\Marimba\AddOns\TagIT.exe");
        }

        private QAProcess RunLync() {
            //initialize the Lync process with the standard part, then find the correct year case and add the correct .exe
            QAProcess process = new QAProcess("Start Lync",
                                              "Starting up the Lync program on the target PC, to verify that it is functioning");
            switch (OfficeVersion) {
                case 2013:
                case 2016:
                    process.App = "lync.exe";
                    break;
                default: // 2007 or 2010
                    process.App = "communicator.exe";
                    break;
            }
            return process;
        }

        private QAProcess RepairIESettings() {
            return new QAProcess("Run repair on IE settings",                                                               //name
                                 "Running repairs on the IE settings, just to verify they are set to the P&G defaults",     //description
                                 "wscript.exe",                                                                             //application
                                 @"C:\swsetup\IEProductivityPack\IEHealing\IEHealing.vbs");                                 //argument
        }

        private QAProcess OpenChannelViewer() {
            return new QAProcess("Run Channel Viewer",
                                 "Running Channel Viewer to verify installed apps and fix failed/missing apps",
                                 "wscript.exe",
                                 @"C:\HP\Scripts\StartChannelViewer.VBS");
        }

        private QAProcess OpenMPSPortal() {
            return new QAProcess("Open MPS Portal",
                                 "Running IE and navigating to the MPS Portal to install the default printer",
                                 "http://mpsportal.pg.com");
        }

        private QAProcess OpenCertificateManager() {
            return new QAProcess("Open Certificate Manager",
                                 "Opening the Certificate Manager to verify that the T# certificate is downloaded, for the VPN",
                                 "certmgr.msc");
        }

        private QAProcess OpenExcel() {
            return new QAProcess("Open Excel",
                                 "Opening a blank Macro-Enabled Worksheet to preemptively fix an issue with opening other Excel Workbooks",
                                 "excel.exe",
                                 "/m");
        }

        private QAProcess OpenOutlook() {
            return new QAProcess("Open Outlook",
                                 "Opening Outlook to complete Synchronization of the inbox and add the user's archives, if required",
                                 "outlook.exe");
        }

        private QAProcess RunTuneUp() {
            return new QAProcess("Run all pending Tuner updates",
                                 "Running all pending updates through the tuner",
                                 @"C:\Windows\System32\TuneUp\TuneUp.exe",
                                 "/RunNowAll");
        }

        private QAProcess OpenPrinterDoc() {
            return new QAProcess("Open Old Printer Information.txt",
                                 "Opening the Old Printer Information file for getting the user's default printer information",
                                 "notepad.exe",
                                 @"C:\Users\Public\Desktop\Printer Info\Old Printer Information.txt");
        }

        private QAProcess OpenDocumentsFolder() {
            return new QAProcess("Open documents folder",
                                 "Opening the documents folder to verify that the user's files transferred over from the old PC",
                                 "explorer.exe",
                                 Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        }

        private QAProcess RemoveOutlook() {
            return new QAProcess("Remove Outlook",
                                 "Opening the Control Panel so Outlook can be removed from the installed software on the PC",
                                 "control",
                                 "appwiz.cpl");
        }

        private void AddVPNLinkToPublicDesktop(MainWindow siht) {
            siht.PrintLine("Adding Pulse Secure shortcut to the Public Desktop.");
            WshShell shell = new WshShell();
            string shortcutLocation = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory) + "\\Pulse Secure.lnk";
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);
            shortcut.Arguments = "-show";
            if (!X64) {
                //32-bit specific
                //setting all of the path information for the VPN link to be placed on the desktop
                shortcut.TargetPath = "C:\\Program Files\\Common Files\\Juniper Networks\\JamUI\\Pulse.exe";
                shortcut.WorkingDirectory = "C:\\Program Files\\Common Files\\Juniper Networks\\JamUI";
            } else {
                //64-bit specific
                //setting all of the path information for the VPN link to be placed on the desktop
                shortcut.TargetPath = "C:\\Program Files (x86)\\Common Files\\Juniper Networks\\JamUI\\Pulse.exe";
                shortcut.WorkingDirectory = "C:\\Program Files (x86)\\Common Files\\Juniper Networks\\JamUI";
            }
            shortcut.Save();
        }

        private void OfficeLink(MainWindow siht) {
            siht.PrintLine("Adding Office Portal Link to the Public Desktop.");
            urlShortcutToPublicDesktop("Office Portal", "http://portal.office.com");
        }

        private void WebmailLink(MainWindow siht) {
            siht.PrintLine("Adding P&G Webmail Link to the Public Desktop.");
            urlShortcutToPublicDesktop("P&G Webmail", "http://webmail.pg.com");
        }

        private void ITAccessLink(MainWindow siht) {
            siht.PrintLine("Adding IT Access Link to the Public Desktop.");
            urlShortcutToPublicDesktop("IT Access", "http://itaccess.pg.com");
        }

        private void ITSolutionsLink(MainWindow siht) {
            siht.PrintLine("Adding ITSolutions Link to the Public Desktop.");
            urlShortcutToPublicDesktop("IT Solutions", "http://itsolutions.pg.com");
        }

        private void GuestWiFiLink(MainWindow siht) {
            siht.PrintLine("Adding Guest Wi-Fi Link to the Public Desktop.");
            urlShortcutToPublicDesktop("Guest Wi-Fi", "http://guest.pg.com");
        }

        private void EformsLink(MainWindow siht) {
            siht.PrintLine("Adding Eforms Link to the Public Desktop.");
            urlShortcutToPublicDesktop("IT Services Catalog", "https://pgglobalenterprise.service-now.com/ITSM_Portal/home.do");
        }

        private void AddVPNLinkToDesktop(MainWindow siht) {
            siht.PrintLine("Adding Pulse Secure shortcut to the desktop.");
            WshShell shell = new WshShell();
            string shortcutLocation = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\Pulse Secure.lnk";
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);
            shortcut.Arguments = "-show";
            if (!X64) {
                //32-bit specific
                //setting all of the path information for the VPN link to be placed on the desktop
                shortcut.TargetPath = "C:\\Program Files\\Common Files\\Juniper Networks\\JamUI\\Pulse.exe";
                shortcut.WorkingDirectory = "C:\\Program Files\\Common Files\\Juniper Networks\\JamUI";
            } else {
                //64-bit specific
                //setting all of the path information for the VPN link to be placed on the desktop
                shortcut.TargetPath = "C:\\Program Files (x86)\\Common Files\\Juniper Networks\\JamUI\\Pulse.exe";
                shortcut.WorkingDirectory = "C:\\Program Files (x86)\\Common Files\\Juniper Networks\\JamUI";
            }
            shortcut.Save();
        }

        private void PinLyncIconToTaskbar(MainWindow siht) {
            siht.PrintLine("Pinning Lync shortcut to the taskbar.");
            string linkPath = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\";
            switch (OfficeVersion) {
                case 2007:
                case 2010:
                    try {
                        TaskbarPinUnpin(linkPath + "Microsoft Lync\\", "Microsoft Lync 2010.lnk", true, siht);
                    } catch (Exception e) {
                        siht.PrintLine(e.Message);
                    }
                    break;
                case 2013:
                    try {
                        TaskbarPinUnpin(linkPath + "Microsoft Office 2013\\", "Lync 2013.lnk", true, siht);
                    } catch (Exception e) {
                        siht.PrintLine(e.Message);
                    }
                    break;
                case 2016:
                    try {
                        TaskbarPinUnpin(linkPath, "Skype for Business 2016.lnk", true, siht);
                    } catch (Exception e) {
                        siht.PrintLine(e.Message);
                    }
                    break;
            }
        }

        private static void TaskbarPinUnpin(string filePath, string fileName, bool pin, MainWindow siht) {
            if (!System.IO.File.Exists(filePath + fileName)) {
                throw new FileNotFoundException(filePath + fileName);
            }
            
            // create the shell application object
            dynamic shellApplication = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
                        
            dynamic directory = shellApplication.NameSpace(filePath);
            dynamic link = directory.ParseName(fileName);

            //FolderItemVerbs verbs = link.Verbs();
            dynamic verbs = link.Verbs();
            for (int i = 0; i < verbs.Count; i++) {
                FolderItemVerb verb = verbs.Item(i);
                string verbName = verb.Name.Replace("&", string.Empty).ToLower();
                if ((pin && verbName.Equals("pin to taskbar")) || (!pin && verbName.Equals("unpin from taskbar"))) {
                    verb.DoIt();
                }
            }
            shellApplication = null;
        }

        private void urlShortcutToPublicDesktop(string linkName, string linkUrl) {
            string deskDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);

            using (StreamWriter writer = new StreamWriter(deskDir + "\\" + linkName + ".url")) {
                writer.WriteLine("[InternetShortcut]");
                writer.WriteLine("URL=" + linkUrl);
                writer.Flush();
            }
        }
    }
}