using System.Diagnostics;

namespace Quality_Assurance_and_Setup {
    public class QAProcess {
        public string ProccessName { get; set; }
        public string Description { get; set; }
        public string Script { get; set; }

        public QAProcess() {

        }

        public QAProcess(string name, string description) {
            ProccessName = name;
            Description = description;
        }

        public QAProcess(string name, string description, string script) {
            ProccessName = name;
            Description = description;
            Script = script;
        }

        public bool RunScript() {
            Process.Start(Script);
            return true;
        }
    }
}