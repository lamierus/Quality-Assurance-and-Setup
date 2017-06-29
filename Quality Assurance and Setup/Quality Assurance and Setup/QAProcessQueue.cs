using IWshRuntimeLibrary;
using Shell32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quality_Assurance_and_Setup {
    public class QAProcessQueue {
        private List<QAProcess> ProcessQueue = new List<QAProcess>();
        public bool X64 { get;}
        public int OfficeVersion { get; }
        public QAType typeOfQA { get;}

        /*public QAProcessQueue() {

        }*/

        public QAProcessQueue(bool is64Bit, int version, QAType QAtoPerform) {
            X64 = is64Bit;
            OfficeVersion = version;
            typeOfQA = QAtoPerform;
            initializeFullQueue();
        }

        private void initializeFullQueue() {

            object shDesktop = (object)"Desktop";
            WshShell shell = new WshShell();
            string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + @"\Pulse Secure.lnk";
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
            shortcut.Arguments = "-show";

            QAProcess StartVPN = new QAProcess("Start VPN",
                                               "Starts the VPN software for testing purposes");
            if (!X64) {
                //32-bit specific

                //setting all of the path information for the VPN link to be placed on the desktop
                shortcut.TargetPath = @"C:\Program Files\Common Files\Juniper Networks\JamUI\Pulse.exe";
                shortcut.WorkingDirectory = @"C:\Program Files\Common Files\Juniper Networks\JamUI";

                //all of these \'s are required for escapes on the " and \ symbols in the command
                StartVPN.Script = "\"\" \"C:\\Program Files\\Common Files\\Juniper Networks\\JamUI\\Pulse.exe\" -show";
            } else {
                //64-bit specific

                //setting all of the path information for the VPN link to be placed on the desktop
                shortcut.TargetPath = @"C:\Program Files (x86)\Common Files\Juniper Networks\JamUI\Pulse.exe";
                shortcut.WorkingDirectory = @"C:\Program Files (x86)\Common Files\Juniper Networks\JamUI";

                StartVPN.Script = "\"\" \"C:\\Program Files (x86)\\Common Files\\Juniper Networks\\JamUI\\Pulse.exe\" -show";
            }

            shortcut.Save();
            ProcessQueue.Add(StartVPN);

            //initialize the Lync process with the standard part --
            QAProcess StartLync = new QAProcess("Start Lync",
                                                "Starts up the Lync program on the target PC, to verify that it is functioning");
            //-- add the correct .exe, depending on the year of the office version
            if (OfficeVersion >= 2013) {
                StartLync.Script = "lync.exe";
            } else {
                StartLync.Script = "communicator.exe";
            }
            ProcessQueue.Add(StartLync);

            //NON-SPECIFIC
            //Run TagIT to set the asset tag #
            QAProcess RunTagIT = new QAProcess("Start TagIT",
                                               "Run TagIT to set the asset tag #",
                                               "\"\" \"C:\\Program Files\\Marimba\\AddOns\\TagIT.exe\"");
            
            /*
            ::Remove the superfluous links, incase they were migrated
            del /f "C:\Users\Public\Desktop\Avaya VPN Client.lnk"
            del /f "C:\Users\Public\Desktop\MyPassword.lnk"
            del /f "C:\Users\Public\Desktop\Adobe Reader*.lnk"
            /*
            ::merge in the IE security zone and Lync fixes
            start wscript.exe /e:vbscript ".\QA Scripts\IEHealing.vbs"
            start regedit.exe /s ".\QA Scripts\Lync Installing Repeat Fix.reg"
            /*
            ::Use the VB script to pin a Lync 2010 icon to the taskbar
            cscript ".\QA Scripts\PinItem.vbs" /taskbar /item:"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Lync\Microsoft Lync 2010.lnk" /q
            cscript ".\QA Scripts\PinItem.vbs" /taskbar /item:"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Lync 2013.lnk" /q
            cscript ".\QA Scripts\PinItem.vbs" /taskbar /item:"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Skype for Business 2016.lnk" /q

            /*
            ::Start Channel Viewer to make sure everything is installed correctly
            start wscript.exe C:\HP\Scripts\StartChannelViewer.VBS
            /*
            ::Start IE to check Favorites and send to the Xerox MPS Portal
            start iexplore http://mpsportal.pg.com
            /*
            ::Open the Old Printer Info text to add the user's default printer
            explorer C:\Users\Public\Desktop\Printer Info\Old Printer Information.txt
            /*
            ::Open the Documents folder to verify transfer of documents
            explorer C:\Users\%userprofile%\Documents\
            /*
            ::Start Certificate Manager to remove the old junos pulse certificate, if required
            start certmgr.msc
            /*
            ::Start Excel to clear any errors on first run
            start "" "excel.exe" /m
            /*
            ::Start Outlook to verify transfer and sync with Exchange
            start outlook.exe
            /*
            ::Run all of the updates
            start "" "C:\Windows\System32\TuneUp\TuneUp.exe" /RunNowAll
            */

        }

        private static void PinUnpinTaskBar(string filePath, bool pin) {
            if (!System.IO.File.Exists(filePath)) throw new System.IO.FileNotFoundException(filePath);

            // create the shell application object
            Shell shellApplication = new Shell();

            string path = System.IO.Path.GetDirectoryName(filePath);
            string fileName = System.IO.Path.GetFileName(filePath);

            Shell32.Folder directory = shellApplication.NameSpace(path);
            FolderItem link = directory.ParseName(fileName);

            FolderItemVerbs verbs = link.Verbs();
            for (int i = 0; i < verbs.Count; i++) {
                FolderItemVerb verb = verbs.Item(i);
                string verbName = verb.Name.Replace(@"&", string.Empty).ToLower();

                if ((pin && verbName.Equals("pin to taskbar")) || (!pin && verbName.Equals("unpin from taskbar"))) {

                    verb.DoIt();
                }
            }

            shellApplication = null;
        }


        public void executeQueue(MainWindow siht) {
            foreach (QAProcess QAp in ProcessQueue) {
                try {
                    if (!QAp.runScript()) {

                    } else {
                        
                    }
                } catch (Exception ex) {
                    siht.printLine(ex.Message);
                }
            }
        }
    }
}
