using System;
using vbSparkle.EvaluationObjects;
using static VBPreprocessorsParser;

namespace vbSparkle.PreProcessor.Statements
{
    public class VbLtDateTime : VBLiteral<LtDateContext>
    {
        public VbLtDateTime(IVBScopeObject context, LtDateContext @object)
            : base(context, @object)
        {
            string date = @object.GetText();
            date = date.Substring(1, date.Length - 2);

            Value = new DDateTimeExpression(DateTime.Parse(date));
        }

        public override string Prettify()
        {
            DDateTimeExpression val = (DDateTimeExpression)Value;
            return val.ToExpressionString();
        }
    }
}
