using static vbSparkle.VBScriptParser;

namespace vbSparkle
{
    public class VbLtNull : VBLiteral<LtNullContext>
    {
        public VbLtNull(IVBScopeObject context, LtNullContext @object)
            : base(context, @object)
        {
            Value = new DCodeBlock("Null");
        }

        public override string Prettify()
        {
            return "Null";
        }
    }

}