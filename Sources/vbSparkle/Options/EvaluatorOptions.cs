using System.Collections.Generic;

namespace vbSparkle.Options
{
    public class LargeStringAllocationObserver
    {
        public List<string> LargeStringAllocated { get; set; }= new List<string>();
        public int MinSize { get; internal set; } = 50;
    }

    public class ExecuteObserver
    {
        public List<string> VBScriptExecuted { get; set; } = new List<string>();
    }


    public class CreateObjectObserver
    {
        public List<string> CreateObjectObserved { get; set; } = new List<string>();
    }

    public class EvaluatorOptions
    {
        public LargeStringAllocationObserver LargeStringAllocationObserver  { get;set;} = null;
        public ExecuteObserver ExecuteObserver { get; set; } = null;
        public CreateObjectObserver CreateObjectObserver { get; set; } = null;

        public SymbolRenamingMode SymbolRenamingMode { get; set; } = SymbolRenamingMode.None;
        public JunkCodeProcessingMode JunkCodeProcessingMode { get; set; } = JunkCodeProcessingMode.Comment;

        public int IndentSpacing { get; set; } = 4;


        //TODO: public bool PerfomPartialEvaluation { get; set; } = true;
        #region Internal
        internal int ConstIdx { get; set; } = 0;
        internal int VarIdx { get; set; } = 0;
        #endregion

    }
}
