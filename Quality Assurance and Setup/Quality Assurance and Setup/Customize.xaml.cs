using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Quality_Assurance_and_Setup {
    /// <summary>
    /// Interaction logic for Customize.xaml
    /// </summary>
    public partial class Customize : Window {
        public Customize(ref QAProcessQueue QAQueue) {
            InitializeComponent();
            dgQueueComponents.ItemsSource = QAQueue.ProcessQueue;
        }
    }
}
