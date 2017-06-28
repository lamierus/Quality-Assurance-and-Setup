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
            if (X64) {
                /*
                ::64-BIT SPECIFIC
                ::Create a Junos Pulse link on the Desktop
                echo Set oWS = WScript.CreateObject("WScript.Shell") > CreateShortcut.vbs
                echo sLinkFile = "C:\Users\Public\Desktop\Pulse Secure.lnk" >> CreateShortcut.vbs
                echo Set oLink = oWS.CreateShortcut(sLinkFile) >> CreateShortcut.vbs
                echo oLink.TargetPath = "C:\Program Files (x86)\Common Files\Juniper Networks\JamUI\Pulse.exe" >> CreateShortcut.vbs
                echo oLink.Arguments = "-show" >> CreateShortcut.vbs
                echo oLink.Save >> CreateShortcut.vbs
                cscript CreateShortcut.vbs
                del CreateShortcut.vbs

                ::Start Junos Pulse to test
                start "" "C:\Program Files (x86)\Common Files\Juniper Networks\JamUI\Pulse.exe" -show
                */
            } else {
                //add 32-bit specific stuff here
            }

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

            /*
            ::NON-SPECIFIC
            ::Run TagIT to set the asset tag #
            start "" "C:\Program Files\Marimba\AddOns\TagIT.exe"

            ::Remove the superfluous links, incase they were migrated
            del /f "C:\Users\Public\Desktop\Avaya VPN Client.lnk"
            del /f "C:\Users\Public\Desktop\MyPassword.lnk"
            del /f "C:\Users\Public\Desktop\Adobe Reader*.lnk"

            ::merge in the IE security zone and Lync fixes
            start wscript.exe /e:vbscript ".\QA Scripts\IEHealing.vbs"
            start regedit.exe /s ".\QA Scripts\Lync Installing Repeat Fix.reg"

            ::Use the VB script to pin a Lync 2010 icon to the taskbar
            cscript ".\QA Scripts\PinItem.vbs" /taskbar /item:"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Lync\Microsoft Lync 2010.lnk" /q
            cscript ".\QA Scripts\PinItem.vbs" /taskbar /item:"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Lync 2013.lnk" /q
            cscript ".\QA Scripts\PinItem.vbs" /taskbar /item:"%ProgramData%\Microsoft\Windows\Start Menu\Programs\Skype for Business 2016.lnk" /q

            ::Start Channel Viewer to make sure everything is installed correctly
            start wscript.exe C:\HP\Scripts\StartChannelViewer.VBS

            ::Start IE to check Favorites and send to the Xerox MPS Portal
            start iexplore http://mpsportal.pg.com

            ::Open the Old Printer Info text to add the user's default printer
            explorer C:\Users\Public\Desktop\Printer Info\Old Printer Information.txt

            ::Open the Documents folder to verify transfer of documents
            explorer C:\Users\%userprofile%\Documents\

            ::Start Certificate Manager to remove the old junos pulse certificate, if required
            start certmgr.msc

            ::Start Excel to clear any errors on first run
            start "" "excel.exe" /m

            ::Start Outlook to verify transfer and sync with Exchange
            start outlook.exe

            ::Run all of the updates
            start "" "C:\Windows\System32\TuneUp\TuneUp.exe" /RunNowAll
            */

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
