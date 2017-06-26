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
using System.Management;//.Instrumentation;

namespace Quality_Assurance_and_Setup {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        static bool Is64Bit = Environment.Is64BitOperatingSystem;
        static string ComputerName = Environment.MachineName.ToString();

        public MainWindow() {
            InitializeComponent();
            
            if (Is64Bit) {
                lblWindowsVersion.Content =  "64-Bit";
            } else {
                lblWindowsVersion.Content =  "32-Bit";
            }

            lblOfficeVersion.Content = findOfficeVersion(Is64Bit);
        }

        // 32-bit
        //2007 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\12.0\Excel\InstallRoot
        //2010 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Excel\InstallRoot
        //2013 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Excel\InstallRoot
        //2016 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot

        // 64-bit
        //2007 - HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Excel\InstallRoot
        //2010 - HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Excel\InstallRoot
        //2013 - HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Excel\InstallRoot
        //2016 - HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Excel\InstallRoot
        private string findOfficeVersion(bool x64) {
            //string officeVersion = null;
            string keyString;

            if (x64) {
                keyString = "SOFTWARE\\Wow6432Node\\Microsoft\\Office\\";
            } else {
                keyString = "SOFTWARE\\Microsoft\\Office\\";
            }

            RegistryKey localMachine = Registry.LocalMachine.OpenSubKey(keyString);

            //string versionKey = string.Empty;

            foreach (string key in localMachine.GetSubKeyNames()) {
                if (key == "12.0") {
                    //2007
                    if (checkIfNull(keyString + key)) {
                        return "Microsoft Office 2007";
                    }
                } else if (key == "14.0") {
                    //2010
                    if (checkIfNull(keyString + key)) {
                        return "Microsoft Office 2010";
                    }
                } else if (key == "15.0") {
                    //2013
                    if (checkIfNull(keyString + key)) {
                        return "Microsoft Office 2013";
                    }
                } else if (key == "16.0") {
                    //2016
                    if (checkIfNull(keyString + key)) {
                        return "Microsoft Office 2016";
                    }
                }

                /*if (!string.IsNullOrEmpty(versionKey)) {
                    break;
                }*/
            }

            return "Office Not Installed";
        }

        private bool checkIfNull(string keyString) {
            RegistryKey exists = Registry.LocalMachine.OpenSubKey(keyString + "\\Excel\\InstallRoot");
            if (exists == null) {
                return false;
            }
            return true;
        }

        private void rbCustomerPC_Checked(object sender, RoutedEventArgs e) {

        }

        private void rbLoanerPC_Checked(object sender, RoutedEventArgs e) {

        }

        private void rbKioskPC_Checked(object sender, RoutedEventArgs e) {

        }

        private void btnBeginProcess_Click(object sender, RoutedEventArgs e) {

        }
    }
}