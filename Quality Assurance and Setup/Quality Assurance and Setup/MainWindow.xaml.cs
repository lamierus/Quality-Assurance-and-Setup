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

namespace Quality_Assurance_and_Setup {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        static bool Is64bit = Environment.Is64BitOperatingSystem;

        public MainWindow() {
            InitializeComponent();
            
            if (Is64bit) {
                lblWindowsVersion.Content = "64-Bit";
            } else {
                lblWindowsVersion.Content = "32-Bit";
            }
        }

    }
}