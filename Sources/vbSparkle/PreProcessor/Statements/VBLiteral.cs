namespace vbSparkle.PreProcessor.Statements
{
    public class VBLiteral<T> : VBLiteral
        where T : VBPreprocessorsParser.ILiteralContext
    {
        public T Object { get; set; }

        public VBLiteral(IVBScopeObject context, T @object)
        {
            Object = @object;
            Context = context;
            Value = new DCodeBlock(@object?.GetText());
        }

        public override string Prettify()
        {
            return Object.GetText();
        }
    }

    public abstract class VBLiteral
    {
        public IVBScopeObject Context { get; set; }
        public DExpression Value { get; set; }

        public abstract string Prettify();
    }
}
