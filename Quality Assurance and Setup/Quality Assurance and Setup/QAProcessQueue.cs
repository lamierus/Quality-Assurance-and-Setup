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

            QAProcess processToAdd = new QAProcess("Start VPN",
                                                   "Starts the VPN software for testing purposes");
            if (!X64) {
                //32-bit specific

                //setting all of the path information for the VPN link to be placed on the desktop
                shortcut.TargetPath = @"C:\Program Files\Common Files\Juniper Networks\JamUI\Pulse.exe";
                shortcut.WorkingDirectory = @"C:\Program Files\Common Files\Juniper Networks\JamUI";

                //all of these \'s are required for escapes on the " and \ symbols in the command
                processToAdd.Script = "\"\" \"C:\\Program Files\\Common Files\\Juniper Networks\\JamUI\\Pulse.exe\" -show";
            } else {
                //64-bit specific

                //setting all of the path information for the VPN link to be placed on the desktop
                shortcut.TargetPath = @"C:\Program Files (x86)\Common Files\Juniper Networks\JamUI\Pulse.exe";
                shortcut.WorkingDirectory = @"C:\Program Files (x86)\Common Files\Juniper Networks\JamUI";

                processToAdd.Script = "\"\" \"C:\\Program Files (x86)\\Common Files\\Juniper Networks\\JamUI\\Pulse.exe\" -show";
            }
            shortcut.Save();
            ProcessQueue.Add(processToAdd);

            //NON-SPECIFIC

            //Run TagIT to set the asset tag #
            processToAdd = new QAProcess("Start TagIT",
                                         "Run TagIT to set the asset tag #",
                                         "\"\" \"C:\\Program Files\\Marimba\\AddOns\\TagIT.exe\"");
            ProcessQueue.Add(processToAdd);

            //initialize the Lync process with the standard part, then find the correct year case and add the correct .exe,
            //  while pinning the correct link on the taskbar.
            processToAdd = new QAProcess("Start Lync",
                                         "Starts up the Lync program on the target PC, to verify that it is functioning");
            switch (OfficeVersion) {
                case 2007:
                case 2010:
                    pinToTaskBar(@"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Lync\Microsoft Lync 2010.lnk", true);
                    processToAdd.Script = "communicator.exe";
                    break;
                case 2013:
                    pinToTaskBar(@"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Lync 2013.lnk", true);
                    processToAdd.Script = "lync.exe";
                    break;
                case 2016:
                    pinToTaskBar(@"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Skype for Business 2016.lnk", true);
                    processToAdd.Script = "lync.exe";
                    break;
            }
            ProcessQueue.Add(processToAdd);

            //check for and run the IEHealing script installed on all O&G machines.
            processToAdd = new QAProcess("Run repair on IE settings",
                                         "Run repairs on the IE settings, just to verify they are set to the P&G defaults",
                                         "wscript.exe /e:vbscript \"C:\\swsetup\\IEProductivityPack\\IEHealing\\IEHealing.vbs\"");
            ProcessQueue.Add(processToAdd);

            //Open Channel Viewer on the PC to verify what apps are installed/failed/licensed and not installed.
            processToAdd = new QAProcess("Run Channel Viewer",
                                         "Run Channel Viewer to verify installed apps and fix failed/missing apps",
                                         @"wscript.exe C:\HP\Scripts\StartChannelViewer.VBS");
            ProcessQueue.Add(processToAdd);

            //Open Internet Explorer and go to the Managed Print Services Portal
            processToAdd = new QAProcess("Open MPS Portal",
                                         "Run IE and navigate to the MPS Portal to install the default printer",
                                         @"iexplore http://mpsportal.pg.com");
            ProcessQueue.Add(processToAdd);

            //Open Internet Explorer and go to the Managed Print Services Portal
            processToAdd = new QAProcess("",
                                         "",
                                         "certmgr.msc");
            ProcessQueue.Add(processToAdd);
            
            //
            processToAdd = new QAProcess("",
                                         "",
                                         "\"\" \"excel.exe\" /m");
            ProcessQueue.Add(processToAdd);

            //
            processToAdd = new QAProcess("",
                                         "",
                                         "outlook.exe");
            ProcessQueue.Add(processToAdd);

            //
            processToAdd = new QAProcess("",
                                         "",
                                         "\"\" \"C:\\Windows\\System32\\TuneUp\\TuneUp.exe\" /RunNowAll");
            ProcessQueue.Add(processToAdd);

            //
            processToAdd = new QAProcess("",
                                         "",
                                         @"notepad.exe C:\Users\Public\Desktop\Printer Info\Old Printer Information.txt");
            ProcessQueue.Add(processToAdd);

            //
            processToAdd = new QAProcess("",
                                         "",
                                         @"explorer.exe C:\Users\%userprofile%\Documents\");
            ProcessQueue.Add(processToAdd);

        }

        private static void pinToTaskBar(string filePath, bool pin) {
            if (!System.IO.File.Exists(filePath)) {
                throw new System.IO.FileNotFoundException(filePath);
            }

            // create the shell application object
            Shell shellApplication = new Shell();

            string path = System.IO.Path.GetDirectoryName(filePath);
            string fileName = System.IO.Path.GetFileName(filePath);

            Shell32.Folder directory = shellApplication.NameSpace(path);
            FolderItem link = directory.ParseName(fileName);

            FolderItemVerbs verbs = link.Verbs();
            for (int i = 0; i < verbs.Count; i++) {
                FolderItemVerb verb = verbs.Item(i);
                string verbName = verb.Name.Replace("&", string.Empty).ToLower();

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
