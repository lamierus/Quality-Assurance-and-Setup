using System.Diagnostics;
using System;

namespace Quality_Assurance_and_Setup {
    public class QAProcess {
        public string ProccessName { get; set; }
        public string Description { get; set; }
        private string App;
        private string Arguments;

        public QAProcess() { }

        public QAProcess(string name, string description) {
            ProccessName = name;
            Description = description;
        }

        public QAProcess(string name, string description, string app) {
            ProccessName = name;
            Description = description;
            App = app;
        }

        public QAProcess(string name, string description, string app, string arguments) {
            ProccessName = name;
            Description = description;
            App = app;
            Arguments = arguments;
        }

        public void SetAppString(string app) {
            App = app;
        }

        public void SetArguments(string arguments) {
            Arguments = arguments;
        }

        public string Run() {
            try {
                Process p = new Process();
                p.StartInfo.FileName = App;
                p.StartInfo.Arguments = Arguments;
                p.Start();
                return Description;
            } catch (Exception e) {
                return Description + Environment.NewLine + "Error! - " + e.Message;
            }
        }
    }
}