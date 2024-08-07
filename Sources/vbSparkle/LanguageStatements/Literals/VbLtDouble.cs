using static VBScriptParser;

namespace vbSparkle
{
    public class VbLtDouble : VBLiteral<LtDoubleContext>
    {
        public VbLtDouble(IVBScopeObject context, LtDoubleContext @object)
            : base(context, @object)
        {
            string quoted = @object.GetText();
            Value = new DMathExpression<double>(double.Parse(quoted));
        }

        public override string Prettify()
        {
            DMathExpression<double> val = (DMathExpression<double>)Value;
            return $"{val.GetRealValue()}d";
        }
    }

}