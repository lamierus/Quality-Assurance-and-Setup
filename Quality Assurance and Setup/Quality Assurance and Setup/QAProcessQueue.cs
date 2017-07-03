﻿using IWshRuntimeLibrary;
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
        public QAType TypeOfQA { get;}
        
        public QAProcessQueue(bool is64Bit, int version, QAType QAtoPerform) {
            X64 = is64Bit;
            OfficeVersion = version;
            TypeOfQA = QAtoPerform;
            if (TypeOfQA == QAType.Customer)
                InitializeFullCustomerQueue();
        }

        private void InitializeFullCustomerQueue() {
            QAProcess processToAdd = new QAProcess("Start VPN",
                                                   "Starting the VPN software for testing purposes");
            if (!X64) {
                //32-bit specific
                //all of these \'s are required for escapes on the " and \ symbols in the command
                processToAdd.Script = "\"\" \"C:\\Program Files\\Common Files\\Juniper Networks\\JamUI\\Pulse.exe\" -show";
            } else {
                processToAdd.Script = "\"\" \"C:\\Program Files (x86)\\Common Files\\Juniper Networks\\JamUI\\Pulse.exe\" -show";
            }
            ProcessQueue.Add(processToAdd);

            //NON-SPECIFIC

            //Run TagIT to set the asset tag #
            processToAdd = new QAProcess("Start TagIT",
                                         "Running TagIT to set the asset tag #",
                                         "\"\" \"C:\\Program Files\\Marimba\\AddOns\\TagIT.exe\"");
            ProcessQueue.Add(processToAdd);

            //initialize the Lync process with the standard part, then find the correct year case and add the correct .exe,
            //  while pinning the correct link on the taskbar.
            processToAdd = new QAProcess("Start Lync",
                                         "Starting up the Lync program on the target PC, to verify that it is functioning");
            switch (OfficeVersion) {
                case 2007:
                case 2010:
                    processToAdd.Script = "communicator.exe";
                    break;
                case 2013:
                    processToAdd.Script = "lync.exe";
                    break;
                case 2016:
                    processToAdd.Script = "lync.exe";
                    break;
            }
            ProcessQueue.Add(processToAdd);
            
            //the rest of the processes have the basic idea laid out in the Description line
            processToAdd = new QAProcess("Run repair on IE settings",
                         /*Description*/ "Running repairs on the IE settings, just to verify they are set to the P&G defaults",
                         /*Script*/      @"wscript.exe /e:vbscript C:\swsetup\IEProductivityPack\IEHealing\IEHealing.vbs");
            ProcessQueue.Add(processToAdd);
            
            processToAdd = new QAProcess("Run Channel Viewer",
                         /*Description*/ "Running Channel Viewer to verify installed apps and fix failed/missing apps",
                         /*Script*/      @"wscript.exe C:\HP\Scripts\StartChannelViewer.VBS");
            ProcessQueue.Add(processToAdd);
            
            processToAdd = new QAProcess("Open MPS Portal",
                         /*Description*/ "Running IE and navigating to the MPS Portal to install the default printer",
                         /*Script*/      @"iexplore http://mpsportal.pg.com");
            ProcessQueue.Add(processToAdd);
            
            processToAdd = new QAProcess("Open Certificate Manager",
                         /*Description*/ "Opening the Certificate Manager to verify that the T# certificate is downloaded, for the VPN",
                         /*Script*/      "certmgr.msc");
            ProcessQueue.Add(processToAdd);
            
            processToAdd = new QAProcess("Open Excel",
                         /*Description*/ "Opening a blank Macro-Enabled Worksheet to preemptively fix an issue with opening other Excel Workbooks",
                         /*Script*/      "\"\" \"excel.exe\" /m");
            ProcessQueue.Add(processToAdd);
            
            processToAdd = new QAProcess("Open Outlook",
                         /*Description*/ "Opening Outlook to complete Synchronization of the inbox and add the user's archives, if required",
                         /*Script*/      "outlook.exe");
            ProcessQueue.Add(processToAdd);
            
            processToAdd = new QAProcess("Run all pending Tuner updates",
                         /*Description*/ "Running all pending updates through the tuner",
                         /*Script*/      "\"\" \"C:\\Windows\\System32\\TuneUp\\TuneUp.exe\" /RunNowAll");
            ProcessQueue.Add(processToAdd);
            
            processToAdd = new QAProcess("Open Old Printer Information.txt",
                         /*Description*/ "Opening the Old Printer Information file for getting the user's default printer information",
                         /*Script*/      @"notepad.exe C:\Users\Public\Desktop\Printer Info\Old Printer Information.txt");
            ProcessQueue.Add(processToAdd);
            
            processToAdd = new QAProcess("Open documents folder",
                         /*Description*/ "Opening the documents folder to verify that the user's files transferred over from the old PC",
                         /*Script*/      @"explorer.exe C:\Users\%userprofile%\Documents\");
            ProcessQueue.Add(processToAdd);
        }

        public void ExecuteQueue(MainWindow siht) {
            foreach (QAProcess QAp in ProcessQueue) {
                siht.PrintLine(QAp.RunScript());
            }

            object shDesktop = (object)"Desktop";
            WshShell shell = new WshShell();
            string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + @"\Pulse Secure.lnk";
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
            shortcut.Arguments = "-show";
            if (!X64) {
                //32-bit specific
                //setting all of the path information for the VPN link to be placed on the desktop
                shortcut.TargetPath = @"C:\Program Files\Common Files\Juniper Networks\JamUI\Pulse.exe";
                shortcut.WorkingDirectory = @"C:\Program Files\Common Files\Juniper Networks\JamUI";
            } else {
                //64-bit specific
                //setting all of the path information for the VPN link to be placed on the desktop
                shortcut.TargetPath = @"C:\Program Files (x86)\Common Files\Juniper Networks\JamUI\Pulse.exe";
                shortcut.WorkingDirectory = @"C:\Program Files (x86)\Common Files\Juniper Networks\JamUI";
            }
            shortcut.Save();

            switch (OfficeVersion) {
                case 2007:
                case 2010:
                    try {
                        TaskbarPinUnpin(@"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Lync\Microsoft Lync 2010.lnk", true);
                    } catch (Exception e) {
                        siht.PrintLine(e.Message);
                    }
                    break;
                case 2013:
                    try {
                        TaskbarPinUnpin(@"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Lync 2013.lnk", true);
                    } catch (Exception e) {
                        siht.PrintLine(e.Message);
                    }
                    break;
                case 2016:
                    try {
                    TaskbarPinUnpin(@"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Skype for Business 2016.lnk", true);
                    } catch (Exception e) {
                        siht.PrintLine(e.Message);
                    }
                    break;
            }
        }

        private static void TaskbarPinUnpin(string filePath, bool pin) {
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
    }
}
