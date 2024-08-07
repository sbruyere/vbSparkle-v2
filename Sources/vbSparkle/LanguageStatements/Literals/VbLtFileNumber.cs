using System;
using static VBScriptParser;

namespace vbSparkle
{
    public class VbLtFileNumber : VBLiteral<LtFilenumberContext>
    {
        public VbLtFileNumber(IVBScopeObject context, LtFilenumberContext @object)
            : base(context, @object)
        {
            string quoted = @object.GetText();
            Value = new DMathExpression<Int32>(int.Parse(quoted.Replace("#", "")));
        }

        public override string Prettify()
        {
            DMathExpression<Int32> val = (DMathExpression<Int32>)Value;
            return $"#{val.GetRealValue()}";
        }
    }

}