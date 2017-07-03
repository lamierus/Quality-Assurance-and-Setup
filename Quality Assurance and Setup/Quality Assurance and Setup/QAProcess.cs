using System.Diagnostics;
using System;

namespace Quality_Assurance_and_Setup {
    public class QAProcess {
        public string ProccessName { get; set; }
        public string Description { get; set; }
        public string Script { get; set; }
        
        public QAProcess(string name, string description) {
            ProccessName = name;
            Description = description;
        }

        public QAProcess(string name, string description, string script) {
            ProccessName = name;
            Description = description;
            Script = script;
        }

        public string RunScript() {
            try {
                Process.Start(Script);
                return Description;
            } catch (Exception e){
                return Description + Environment.NewLine + "Error! - " + e.Message;
            }
        }
    }
}