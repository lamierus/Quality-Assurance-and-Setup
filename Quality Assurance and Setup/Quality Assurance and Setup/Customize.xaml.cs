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
        public QAProcessQueue ModifiedQueue;
        public bool AddDesktopIcons = true;
        private QAProcessQueue AvailableItems = new QAProcessQueue();
        private QAProcessQueue UnmodifiedQueue;
        private QAProcess DesktopIcons = new QAProcess("Add Desktop Icons",
                                                       "Add icons to the desktop");

        public Customize(QAProcessQueue QAQueue) {
            InitializeComponent();
            UnmodifiedQueue = QAQueue;
            ResetQueue();
            dgQueueComponents.ItemsSource = ModifiedQueue.ProcessQueue;
            dgAvailableComponents.ItemsSource = AvailableItems.ProcessQueue;
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e) {
            QAProcess toAdd = (QAProcess)dgAvailableComponents.SelectedItem;
            AvailableItems.RemoveFromQueue(toAdd);
            ModifiedQueue.AddToQueue(toAdd);
        }

        private void btnRemove_Click(object sender, RoutedEventArgs e) {
            QAProcess toRemove = (QAProcess)dgQueueComponents.SelectedItem;
            ModifiedQueue.RemoveFromQueue(toRemove);
            AvailableItems.AddToQueue(toRemove);
        }

        private void ResetQueue() {
            ModifiedQueue = UnmodifiedQueue;
            ModifiedQueue.AddToQueue(DesktopIcons);
            AvailableItems.ClearQueue();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e) {
            DialogResult = false;
            ModifiedQueue = UnmodifiedQueue;
            ModifiedQueue.RemoveFromQueue(DesktopIcons);
            Close();
        }

        private void btnReset_Click(object sender, RoutedEventArgs e) {
            ResetQueue();
        }

        private void btnSave_Click(object sender, RoutedEventArgs e) {
            DialogResult = true;
            if (!ModifiedQueue.Find(DesktopIcons)) {
                AddDesktopIcons = false;
            } else {
                ModifiedQueue.RemoveFromQueue(DesktopIcons);
            }
            Close();
        }
    }
}
