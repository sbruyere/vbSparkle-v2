using System;
using vbSparkle.EvaluationObjects;
using static vbSparkle.VBScriptParser;

namespace vbSparkle
{
    public class VbLtBoolean : VBLiteral<LtBooleanContext>
    {
        public VbLtBoolean(IVBScopeObject context, LtBooleanContext @object)
            : base(context, @object)
        {
            string quoted = @object.GetText();
            if (quoted.Equals("True", StringComparison.InvariantCultureIgnoreCase))
                Value = new DBoolExpression(true);
            else 
                Value = new DBoolExpression(false);
        }

        public override string Prettify()
        {
            DBoolExpression val = (DBoolExpression)Value;

            return val.ToExpressionString();
        }
    }

}