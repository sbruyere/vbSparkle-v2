using static vbSparkle.VBScriptParser;

namespace vbSparkle
{
    public class VbLtNothing : VBLiteral<LtNothingContext>
    {
        public VbLtNothing(IVBScopeObject context, LtNothingContext @object)
            : base(context, @object)
        {
            Value = new DCodeBlock("Nothing");
        }

        public override string Prettify()
        {
            return "Nothing";
        }
    }

}