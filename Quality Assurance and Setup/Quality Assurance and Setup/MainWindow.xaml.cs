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
        string OfficeVersion;

        public MainWindow() {
            InitializeComponent();

            if (Is64Bit) {
                lblWindowsVersion.Content = "64-Bit";
            } else {
                lblWindowsVersion.Content = "32-Bit";
            }

            //get the currently installed office version and assign this value to the global OfficeVersion
            // as well as set the content of the label to the found version string
            lblOfficeVersion.Content = OfficeVersion = FindOfficeVersion();
            //extract just the year of the office version from the string as a whole.
            OfficeVersion = OfficeVersion.Substring(OfficeVersion.IndexOf('2'));
            rbCustomerPC.IsChecked = true;
        }
        
        private string FindOfficeVersion() {
            string returnString = "Office Not Installed";

            RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe");
            if (key == null) {
                IsOfficeInstalled = false;
            } else {
                IsOfficeInstalled = true;
            }

            string oVersion = key.GetValue("Path").ToString();
            if (oVersion.Contains("Office12")) {
                returnString = "Microsoft Office 2007";
            } 
            else if (oVersion.Contains("Office14")) {
                returnString = "Microsoft Office 2010";
            } 
            else if (oVersion.Contains("Office15")) {
                returnString = "Microsoft Office 2013";
            } 
            else if (oVersion.Contains("Office16")) {
                returnString = "Microsoft Office 2016";
            }

            return returnString;
        }

        private void RBCustomerPC_Checked(object sender, RoutedEventArgs e) {
            TypeOfQA = QAType.Customer;
            PrintLine(TypeOfQA.ToString() + " QA process selected!");
            InitializeQueue();
        }

        private void RBLoanerPC_Checked(object sender, RoutedEventArgs e) {
            TypeOfQA = QAType.Loaner;
            PrintLine(TypeOfQA.ToString() + " QA process selected!");
            InitializeQueue();
        }

        private void RBKioskPC_Checked(object sender, RoutedEventArgs e) {
            TypeOfQA = QAType.Kiosk;
            PrintLine(TypeOfQA.ToString() + " QA process selected!");
            InitializeQueue();
        }

        private void InitializeQueue() {
            QAQueue = new QAProcessQueue(Is64Bit, int.Parse(OfficeVersion), TypeOfQA);
        }

        public void PrintLine(string textToPrint) {
            tboxOutput.Text += textToPrint + Environment.NewLine;
        }

        private void BTNBeginProcess_Click(object sender, RoutedEventArgs e) {

            if (!IsOfficeInstalled) {
                //prompt user to start process to install office, possibly what version
                IsOfficeInstalled = true;
            } else {
                QAQueue.ExecuteQueue(this);
            }
        }

        private void btnCustomize_Click(object sender, RoutedEventArgs e) {

        }
    }
}