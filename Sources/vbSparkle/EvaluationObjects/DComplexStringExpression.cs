using System.Collections.Generic;
using System.Linq;
using System.Text;
using MathNet.Symbolics;
using vbSparkle;
using vbSparkle.Options;

namespace vbSparkle.EvaluationObjects
{
    public class DArrayExpression : DExpression
    {
        public List<DExpression> Items { get; set; }

        public DArrayExpression(int ubound)
        {
            Items = (new DExpression[ubound+1]).ToList();
        }
        public DArrayExpression(IEnumerable<DExpression> array)
        {
            Items = array.ToList();
        }

        public override bool IsValuable { get => true; set => throw new System.NotImplementedException(); }
        public override bool HasSideEffet { get => false; set => throw new System.NotImplementedException(); }

        public override string ToExpressionString()
        {
            string[] dExpressions = Items.Select(v=> v.ToExpressionString()).ToArray();
            return "Array(" + string.Join(", ", dExpressions) + ")";
        }

        public override string ToValueString()
        {
            return ToExpressionString();
        }

        public DExpression this[int index]
        {
            get => Items[index];
            set => Items[index] = value;
        }

        internal override SymbolicExpression GetSymExp()
        {
            return SymbolicExpression.Variable(ToExpressionString());
        }
    }

    internal class DComplexStringExpression
        : DExpression, IStringExpression
    {

        public List<DExpression> ConcatExpressions { get; set; } = new List<DExpression>();

        public override bool HasSideEffet { get; set; } = false;
        public override bool IsValuable { get; set; } = true;

        public DComplexStringExpression()
        {
        }

        public DComplexStringExpression(DExpression leftExp)
        {
            Concat(leftExp);
        }

        public override string ToExpressionString()
        {

            if (IsValuable)
            {
                return VbUtils.StrValToExp(ToValueString());
            }
            else
            {
                List<string> stringResult = new List<string>();
                foreach (var v in ConcatExpressions)
                {
                    //if (v is DSimpleStringExpression)
                    //{
                    stringResult.Add(v.ToExpressionString());
                    //}
                }
                return string.Join(" & ", stringResult);
            }
        }

        internal override SymbolicExpression GetSymExp()
        {
            return SymbolicExpression.Variable(ToExpressionString());
        }

        public override string ToString()
        {
            return ToExpressionString();
        }

        public void Concat(DExpression expression)
        {
            if (expression is DComplexStringExpression)
            {
                foreach(var v in ((DComplexStringExpression)expression).ConcatExpressions)
                {
                    Concat(v);
                }
                return;
            }

            if (!expression.IsValuable)
                IsValuable = false;

            if (expression.HasSideEffet || !expression.IsValuable)
            {
                HasSideEffet = true;
                ConcatExpressions.Add(expression);
                return;
            }

            var curRightExp = ConcatExpressions.LastOrDefault();

            if (curRightExp is DSimpleStringExpression && expression.IsValuable)
            {
                var options = (curRightExp as DSimpleStringExpression).Options;

                ConcatExpressions[ConcatExpressions.Count - 1] = new DSimpleStringExpression(curRightExp.ToValueString() + expression.ToValueString(), Encoding.Unicode, options);
                //((DSimpleStringExpression)curRightExp).SetValue(curRightExp.ToValueString() + expression.ToValueString()); <= this was causing side effect
                return;
            }

            ConcatExpressions.Add(expression);
        }


        public override string ToValueString()
        {
            if (!this.IsValuable)
                throw new System.Exception("Not valuable");

            DSimpleStringExpression exp;

            if (ConcatExpressions.Count() == 1 && (exp = ConcatExpressions.FirstOrDefault() as DSimpleStringExpression) != null)
            {
                return exp.ToValueString();
            }
            else
            {
                StringBuilder stringResult = new StringBuilder();

                foreach (var v in ConcatExpressions)
                {
                    //if (v is DSimpleStringExpression)
                    //{
                    stringResult.Append(v.ToValueString());
                    //}
                }

                return stringResult.ToString();
            }
        }
    }
}
