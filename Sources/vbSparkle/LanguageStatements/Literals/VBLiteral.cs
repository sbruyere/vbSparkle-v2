namespace vbSparkle
{
    public abstract class VBLiteral
    {
        public DExpression Value { get; set; }
        public IVBScopeObject Context { get; set; }

        public abstract string Prettify();
    }

    public class VBLiteral<T> : VBLiteral
        where T : ILiteralContext
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

}