using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Quality_Assurance_and_Setup {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        static bool Is64Bit = Environment.Is64BitOperatingSystem;
        static string ComputerName = Environment.MachineName.ToString();

        QAProcessQueue QAQueue;
        QAType TypeOfQA;
        bool IsOfficeInstalled;

        public MainWindow() {
            InitializeComponent();
            
            if (Is64Bit) {
                lblWindowsVersion.Content =  "64-Bit";
            } else {
                lblWindowsVersion.Content =  "32-Bit";
            }

            string oVersion;
            lblOfficeVersion.Content = oVersion = FindOfficeVersion();
            oVersion = oVersion.Substring(oVersion.IndexOf('2'));

            QAQueue = new QAProcessQueue(Is64Bit, int.Parse(oVersion), TypeOfQA);
        }
        
        private string FindOfficeVersion() {
            string keyString;

            if (Is64Bit) {
                keyString = "SOFTWARE\\Wow6432Node\\Microsoft\\Office\\";
            } else {
                keyString = "SOFTWARE\\Microsoft\\Office\\";
            }
            
            foreach (string key in Registry.LocalMachine.OpenSubKey(keyString).GetSubKeyNames()) {
                if (key == "12.0") {
                    if (CheckIfNull(keyString + key)) {
                        return "Microsoft Office 2007";
                    }
                } else if (key == "14.0") {
                    if (CheckIfNull(keyString + key)) {
                        return "Microsoft Office 2010";
                    }
                } else if (key == "15.0") {
                    if (CheckIfNull(keyString + key)) {
                        return "Microsoft Office 2013";
                    }
                } else if (key == "16.0") {
                    if (CheckIfNull(keyString + key)) {
                        return "Microsoft Office 2016";
                    }
                }
            }

            return "Office Not Installed";
        }

        private bool CheckIfNull(string keyString) {
            RegistryKey exists = Registry.LocalMachine.OpenSubKey(keyString + "\\Excel\\InstallRoot");
            if (exists == null) {
                return IsOfficeInstalled = false;
            }
            return IsOfficeInstalled = true;
        }

        private void RBCustomerPC_Checked(object sender, RoutedEventArgs e) {
            TypeOfQA = QAType.Customer;
            PrintLine(TypeOfQA.ToString() + " QA process selected!");
        }

        private void RBLoanerPC_Checked(object sender, RoutedEventArgs e) {
            TypeOfQA = QAType.Loaner;
            PrintLine(TypeOfQA.ToString() + " QA process selected!");
        }

        private void RBKioskPC_Checked(object sender, RoutedEventArgs e) {
            TypeOfQA = QAType.Kiosk;
            PrintLine(TypeOfQA.ToString() + " QA process selected!");
        }

        public void PrintLine(string textToPrint) {
            tblockOutput.Text += textToPrint + Environment.NewLine;
        }

        private void BTNBeginProcess_Click(object sender, RoutedEventArgs e) {

            if (!IsOfficeInstalled) {
                IsOfficeInstalled = true;
                //prompt user to start process to install office, possibly what version
            } else {
                QAQueue.ExecuteQueue(this);
            }
        }
    }
}