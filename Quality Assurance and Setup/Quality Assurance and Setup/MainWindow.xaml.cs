﻿using System;
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
            //string keyString;

            //if (Is64Bit) {
                string x64KeyString = "SOFTWARE\\Wow6432Node\\Microsoft\\Office\\";
            //} else {
                string x86KeyString = "SOFTWARE\\Microsoft\\Office\\";
            //}
            //x86string.Union(x64string)?
            //foreach (string key in Registry.LocalMachine.OpenSubKey(keyString).GetSubKeyNames()) {
            foreach (string key in Registry.LocalMachine.OpenSubKey(x64KeyString).GetSubKeyNames().Union(Registry.LocalMachine.OpenSubKey(x86KeyString).GetSubKeyNames())) {
                if (key == "12.0") {
                    if (CheckIfNull(x86KeyString + key) || CheckIfNull(x64KeyString + key)) {
                        returnString = "Microsoft Office 2007";
                    }
                } else if (key == "14.0") {
                    if (CheckIfNull(x86KeyString + key) || CheckIfNull(x64KeyString + key)) {
                        returnString = "Microsoft Office 2010";
                    }
                } else if (key == "15.0") {
                    if (CheckIfNull(x86KeyString + key) || CheckIfNull(x64KeyString + key)) {
                        returnString = "Microsoft Office 2013";
                    }
                } else if (key == "16.0") {
                    if (CheckIfNull(x86KeyString + key) || CheckIfNull(x64KeyString + key)) {
                        returnString = "Microsoft Office 2016";
                    }
                }
            }
            return returnString;
        }

        private bool CheckIfNull(string fullKeyString) {
            if (Registry.LocalMachine.OpenSubKey(fullKeyString + "\\Excel\\InstallRoot") == null) {
                return IsOfficeInstalled = false;
            }
            return IsOfficeInstalled = true;
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
    }
}