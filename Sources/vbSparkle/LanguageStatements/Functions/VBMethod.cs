using MathNet.Symbolics;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using vbSparkle.EvaluationObjects;

namespace vbSparkle.NativeMethods
{
    public class VB_MonitoringFunction
    : VbNativeFunction
    {
        public VB_MonitoringFunction(IVBScopeObject context, string Name)
            : base(context, Name)
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            return DefaultExpression(args);
        }

    }


    public class VB_Time
        : VbNativeFunction
    {
        public VB_Time(IVBScopeObject context)
            : base(context, "Time")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {

            return new DCodeBlock("Time");
        }

    }

    //public class VB_CreateObject
    //    : VbNativeFunction
    //{
    //    public VB_CreateObject(IVBScopeObject context)
    //        : base(context, "CreateObject")
    //    {
    //    }

    //    public override DExpression Evaluate(params DExpression[] args)
    //    {
    //        throw new NotImplementedException();
    //    }

    //}

    //public class VB_Array
    //    : VbNativeFunction
    //{
    //    public VB_Array(IVBScopeObject context)
    //        : base(context, "Array")
    //    {
    //    }

    //    public override DExpression Evaluate(params DExpression[] args)
    //    {
    //        throw new NotImplementedException();
    //    }

    //}

    public class VB_Eval
        : VbNativeFunction
    {
        public VB_Eval(IVBScopeObject context)
            : base(context, "Eval")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            if (arg1.IsValuable)
            {
                return arg1;
            }

            return DefaultExpression(args);
        }

    }
    public class VB_StrReverse
    : VbNativeFunction
    {
        public VB_StrReverse(IVBScopeObject context)
            : base(context, "StrReverse")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
                return DefaultExpression(args);

            string str = new string(strArg.ToCharArray().Reverse().ToArray());

            return new DSimpleStringExpression(str, Encoding.Unicode, Context.Options);
        }
    }
    public class VB_EnvironS : VbNativeFunction
    {
        public VB_EnvironS(IVBScopeObject context)
            : base(context, "Environ$")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
            {
                return DefaultExpression(args);
            }

            return new DSimpleStringExpression($"%{strArg}%", Encoding.Unicode, Context.Options);
        }
    }

    public class VB_Environ
    : VbNativeFunction
    {
        public VB_Environ(IVBScopeObject context)
            : base(context, "Environ")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
            {
                return DefaultExpression(args);
            }

            return new DSimpleStringExpression($"%{strArg}%", Encoding.Unicode, Context.Options);
        }
    }

    public class VB_FV : VbNativeFunction
    {
        public VB_FV(IVBScopeObject context)
            : base(context, "FV")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double rate, nper, pmt, pv = 0.0;
                int type = (int)Financial.DueDate.EndOfPeriod;

                // Retrieve mandatory arguments
                if (!Converter.TryGetDoubleValue(args[0], out rate) ||
                    !Converter.TryGetDoubleValue(args[1], out nper) ||
                    !Converter.TryGetDoubleValue(args[2], out pmt))
                {
                    return DefaultExpression(args);
                }

                // Retrieve optional arguments
                if (args.Length > 3 && !Converter.TryGetDoubleValue(args[3], out pv))
                {
                    pv = 0.0; // Default value if not provided
                }

                if (args.Length > 4 && !Converter.TryGetInt32Value(args[4], out type))
                {
                    type = (int)Financial.DueDate.EndOfPeriod; // Default value if not provided
                }

                // Calculate FV
                double fv = Financial.FV(rate, nper, pmt, pv, (Financial.DueDate)type);

                return new DMathExpression<double>(fv) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }

    }

    public class VB_DDB : VbNativeFunction
    {
        public VB_DDB(IVBScopeObject context)
            : base(context, "DDB")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double cost, salvage, life, period;

                // Retrieve arguments
                if (!Converter.TryGetDoubleValue(args[0], out cost) ||
                    !Converter.TryGetDoubleValue(args[1], out salvage) ||
                    !Converter.TryGetDoubleValue(args[2], out life) ||
                    !Converter.TryGetDoubleValue(args[3], out period))
                {
                    return DefaultExpression(args);
                }

                // Calculate DDB
                double rate = 2.0 / life;
                double depreciation = cost * Math.Pow(1 - rate, period - 1) * rate;

                // Ensure depreciation doesn't go below the salvage value
                double accumulatedDepreciation = cost * (1 - Math.Pow(1 - rate, period));
                if (accumulatedDepreciation > (cost - salvage))
                {
                    depreciation = cost - salvage - accumulatedDepreciation + depreciation;
                }

                return new DMathExpression<double>(depreciation) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }

    public class VB_IPmt : VbNativeFunction
    {
        public VB_IPmt(IVBScopeObject context)
            : base(context, "IPmt")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double rate, per, nper, pv;
                double fv = 0.0; // Default value for future value
                int type = 0;    // Default value for type (end of period)

                // Retrieve mandatory arguments
                if (!Converter.TryGetDoubleValue(args[0], out rate) ||
                    !Converter.TryGetDoubleValue(args[1], out per) ||
                    !Converter.TryGetDoubleValue(args[2], out nper) ||
                    !Converter.TryGetDoubleValue(args[3], out pv))
                {
                    return DefaultExpression(args);
                }

                // Retrieve optional arguments
                if (args.Length > 4 && !Converter.TryGetDoubleValue(args[4], out fv))
                {
                    fv = 0.0; // Default if not provided
                }

                if (args.Length > 5 && !Converter.TryGetInt32Value(args[5], out type))
                {
                    type = 0; // Default if not provided (0 = end of period, 1 = beginning)
                }

                // Calculate the interest payment
                double ipmt = Financial.IPmt(rate, per, nper, pv, fv, (Financial.DueDate)type);

                return new DMathExpression<double>(ipmt) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }


    public class VB_IRR : VbNativeFunction
    {
        public VB_IRR(IVBScopeObject context)
            : base(context, "IRR")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                DArrayExpression cashFlowsArray;
                double guess = 0.1;

                // Retrieve the array of cash flows
                if (!Converter.TryGetArrayExpression(args[0], out cashFlowsArray))
                {
                    return DefaultExpression(args);
                }

                // Optionally retrieve the guess value
                if (args.Length > 1 && !Converter.TryGetDoubleValue(args[1], out guess))
                {
                    return DefaultExpression(args);
                }

                // Convert the array items to doubles, return default if conversion fails
                double[] cashFlows = new double[cashFlowsArray.Items.Count];
                for (int i = 0; i < cashFlowsArray.Items.Count; i++)
                {
                    if (!Converter.TryGetDoubleValue(cashFlowsArray.Items[i], out cashFlows[i]))
                    {
                        return DefaultExpression(args);
                    }
                }

                // Calculate IRR using Newton-Raphson method or similar
                double irr = (double)Financial.CalculateIRR(cashFlows, guess);

                return new DMathExpression<double>(irr) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }

    }

    public class VB_MIRR : VbNativeFunction
    {
        public VB_MIRR(IVBScopeObject context)
            : base(context, "MIRR")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                DArrayExpression cashFlowsArray;
                double financeRate, reinvestRate;

                // Retrieve the array of cash flows
                if (!Converter.TryGetArrayExpression(args[0], out cashFlowsArray))
                {
                    return DefaultExpression(args);
                }

                // Retrieve finance rate and reinvestment rate
                if (!Converter.TryGetDoubleValue(args[1], out financeRate) ||
                    !Converter.TryGetDoubleValue(args[2], out reinvestRate))
                {
                    return DefaultExpression(args);
                }

                // Convert the array items to doubles, return default if conversion fails
                double[] cashFlows = new double[cashFlowsArray.Items.Count];
                for (int i = 0; i < cashFlowsArray.Items.Count; i++)
                {
                    if (!Converter.TryGetDoubleValue(cashFlowsArray.Items[i], out cashFlows[i]))
                    {
                        return DefaultExpression(args);
                    }
                }

                // Calculate MIRR
                double mirr = Financial.MIRR(ref cashFlows, financeRate, reinvestRate);

                return new DMathExpression<double>(mirr) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }

    }

    public class VB_NPV : VbNativeFunction
    {
        public VB_NPV(IVBScopeObject context)
            : base(context, "NPV")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double rate;
                DArrayExpression cashFlowsArray;

                // Retrieve the discount rate and cash flows
                if (!Converter.TryGetDoubleValue(args[0], out rate) ||
                    !Converter.TryGetArrayExpression(args[1], out cashFlowsArray))
                {
                    return DefaultExpression(args);
                }

                // Convert the array items to doubles, return default if conversion fails
                double[] cashFlows = new double[cashFlowsArray.Items.Count];
                for (int i = 0; i < cashFlowsArray.Items.Count; i++)
                {
                    if (!Converter.TryGetDoubleValue(cashFlowsArray.Items[i], out cashFlows[i]))
                    {
                        return DefaultExpression(args);
                    }
                }

                // Calculate NPV
                double npv = 0;
                for (int t = 0; t < cashFlows.Length; t++)
                {
                    npv += cashFlows[t] / Math.Pow(1 + rate, t + 1);
                }

                return new DMathExpression<double>(npv) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }


    public class VB_Pmt : VbNativeFunction
    {
        public VB_Pmt(IVBScopeObject context)
            : base(context, "Pmt")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double rate, nper, pv, fv = 0;
                int type = 0;

                // Retrieve arguments
                if (!Converter.TryGetDoubleValue(args[0], out rate) ||
                    !Converter.TryGetDoubleValue(args[1], out nper) ||
                    !Converter.TryGetDoubleValue(args[2], out pv))
                {
                    return DefaultExpression(args);
                }

                // Optional arguments
                if (args.Length > 3 && !Converter.TryGetDoubleValue(args[3], out fv))
                {
                    return DefaultExpression(args);
                }

                if (args.Length > 4 && !Converter.TryGetInt32Value(args[4], out type))
                {
                    return DefaultExpression(args);
                }

                // Calculate Pmt
                double pmt;
                if (rate == 0)
                {
                    pmt = -(pv + fv) / nper;
                }
                else
                {
                    pmt = -((pv * Math.Pow(1 + rate, nper) + fv) / ((1 + rate * type) * (Math.Pow(1 + rate, nper) - 1) / rate));
                }

                return new DMathExpression<double>(pmt) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }

    public class VB_PPmt : VbNativeFunction
    {
        public VB_PPmt(IVBScopeObject context)
            : base(context, "PPmt")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double rate, per, nper, pv;
                double fv = 0.0; // Default value for future value
                int type = 0;    // Default value for type (end of period)

                // Retrieve mandatory arguments
                if (!Converter.TryGetDoubleValue(args[0], out rate) ||
                    !Converter.TryGetDoubleValue(args[1], out per) ||
                    !Converter.TryGetDoubleValue(args[2], out nper) ||
                    !Converter.TryGetDoubleValue(args[3], out pv))
                {
                    return DefaultExpression(args);
                }

                // Retrieve optional arguments
                if (args.Length > 4 && !Converter.TryGetDoubleValue(args[4], out fv))
                {
                    fv = 0.0; // Default if not provided
                }

                if (args.Length > 5 && !Converter.TryGetInt32Value(args[5], out type))
                {
                    type = 0; // Default if not provided (0 = end of period, 1 = beginning)
                }

                // Calculate PPmt
                double ppmt = Financial.PPmt(rate, per, nper, pv, fv, (Financial.DueDate)type);

                return new DMathExpression<double>(ppmt) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }


    public class VB_PV : VbNativeFunction
    {
        public VB_PV(IVBScopeObject context)
            : base(context, "PV")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double rate, nper, pmt, fv = 0;
                int type = 0;

                // Retrieve arguments
                if (!Converter.TryGetDoubleValue(args[0], out rate) ||
                    !Converter.TryGetDoubleValue(args[1], out nper) ||
                    !Converter.TryGetDoubleValue(args[2], out pmt))
                {
                    return DefaultExpression(args);
                }

                // Optional arguments
                if (args.Length > 3 && !Converter.TryGetDoubleValue(args[3], out fv))
                {
                    return DefaultExpression(args);
                }

                if (args.Length > 4 && !Converter.TryGetInt32Value(args[4], out type))
                {
                    return DefaultExpression(args);
                }

                // Calculate PV
                double pv;
                if (rate == 0)
                {
                    pv = -(pmt * nper + fv);
                }
                else
                {
                    pv = -((fv + pmt * (1 + rate * type) * (Math.Pow(1 + rate, nper) - 1) / rate) / Math.Pow(1 + rate, nper));
                }

                return new DMathExpression<double>(pv) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }
    public class VB_Rate : VbNativeFunction
    {
        public VB_Rate(IVBScopeObject context)
            : base(context, "Rate")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double nper, pmt, pv, fv = 0;
                int type = 0;
                double guess = 0.1;

                // Retrieve arguments
                if (!Converter.TryGetDoubleValue(args[0], out nper) ||
                    !Converter.TryGetDoubleValue(args[1], out pmt) ||
                    !Converter.TryGetDoubleValue(args[2], out pv))
                {
                    return DefaultExpression(args);
                }

                // Optional arguments
                if (args.Length > 3 && !Converter.TryGetDoubleValue(args[3], out fv))
                {
                    return DefaultExpression(args);
                }

                if (args.Length > 4 && !Converter.TryGetInt32Value(args[4], out type))
                {
                    return DefaultExpression(args);
                }

                if (args.Length > 5 && !Converter.TryGetDoubleValue(args[5], out guess))
                {
                    return DefaultExpression(args);
                }

                // Calculate Rate using Newton-Raphson method
                double rate = CalculateRate(nper, pmt, pv, fv, type, guess);

                return new DMathExpression<double>(rate) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }

        private double CalculateRate(double nper, double pmt, double pv, double fv, int type, double guess)
        {
            const double tolerance = 1e-6;
            const int maxIterations = 100;
            double rate = guess;

            for (int i = 0; i < maxIterations; i++)
            {
                double fValue = pv * Math.Pow(1 + rate, nper) + pmt * (1 + rate * type) * (Math.Pow(1 + rate, nper) - 1) / rate + fv;
                double fDerivative = nper * pv * Math.Pow(1 + rate, nper - 1) - pmt * (1 + rate * type) * (Math.Pow(1 + rate, nper) - 1) / (rate * rate) + nper * pmt * (1 + rate * type) * Math.Pow(1 + rate, nper - 1) / rate;

                double newRate = rate - fValue / fDerivative;

                if (Math.Abs(newRate - rate) < tolerance)
                {
                    return newRate;
                }

                rate = newRate;
            }

            throw new Exception("Rate did not converge");
        }
    }
    public class VB_SLN : VbNativeFunction
    {
        public VB_SLN(IVBScopeObject context)
            : base(context, "SLN")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double cost, salvage, life;

                // Retrieve arguments
                if (!Converter.TryGetDoubleValue(args[0], out cost) ||
                    !Converter.TryGetDoubleValue(args[1], out salvage) ||
                    !Converter.TryGetDoubleValue(args[2], out life))
                {
                    return DefaultExpression(args);
                }

                // Calculate SLN
                double depreciation = (cost - salvage) / life;

                return new DMathExpression<double>(depreciation) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }
    public class VB_SYD : VbNativeFunction
    {
        public VB_SYD(IVBScopeObject context)
            : base(context, "SYD")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double cost, salvage, life, period;

                // Retrieve arguments
                if (!Converter.TryGetDoubleValue(args[0], out cost) ||
                    !Converter.TryGetDoubleValue(args[1], out salvage) ||
                    !Converter.TryGetDoubleValue(args[2], out life) ||
                    !Converter.TryGetDoubleValue(args[3], out period))
                {
                    return DefaultExpression(args);
                }

                // Calculate SYD
                double sumOfYearsDigits = (life * (life + 1)) / 2;
                double depreciation = ((cost - salvage) * (life - period + 1)) / sumOfYearsDigits;

                return new DMathExpression<double>(depreciation) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }

    public class VB_NPer : VbNativeFunction
    {
        public VB_NPer(IVBScopeObject context)
            : base(context, "NPer")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                double rate, pmt, pv, fv = 0;
                int type = 0;

                // Retrieve arguments
                if (!Converter.TryGetDoubleValue(args[0], out rate) ||
                    !Converter.TryGetDoubleValue(args[1], out pmt) ||
                    !Converter.TryGetDoubleValue(args[2], out pv))
                {
                    return DefaultExpression(args);
                }

                // Optional arguments
                if (args.Length > 3 && !Converter.TryGetDoubleValue(args[3], out fv))
                {
                    return DefaultExpression(args);
                }

                if (args.Length > 4 && !Converter.TryGetInt32Value(args[4], out type))
                {
                    return DefaultExpression(args);
                }

                // Calculate NPer
                double nper;
                if (rate == 0)
                {
                    nper = -(pv + fv) / pmt;
                }
                else
                {
                    nper = Math.Log((pmt * (1 + rate * type) - fv * rate) / (pv * rate + pmt * (1 + rate * type))) / Math.Log(1 + rate);
                }

                return new DMathExpression<double>(nper) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }

    


    public class VB_Array : VbNativeFunction
    {
        public VB_Array(IVBScopeObject context)
            : base(context, "Array")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            // Create a new DArrayExpression from the provided arguments
            var arrayExpression = new DArrayExpression(args);

            // Return the DArrayExpression
            return arrayExpression;
        }
    }

    public class VB_LBound : VbNativeFunction
    {
        public VB_LBound(IVBScopeObject context)
            : base(context, "LBound")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                DArrayExpression arrayExpression;
                int dimension = 1; // Default dimension is 1

                // Retrieve the array
                if (!Converter.TryGetArrayExpression(args[0], out arrayExpression))
                {
                    return DefaultExpression(args);
                }

                // Optionally retrieve the dimension (1-based index)
                if (args.Length > 1 && !Converter.TryGetInt32Value(args[1], out dimension))
                {
                    return DefaultExpression(args);
                }

                if (dimension != 1)
                {
                    throw new ArgumentException("Only single-dimensional arrays are supported.");
                }

                // In VBA, the default lower bound is typically 0
                return new DMathExpression<int>(0) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }
    public class VB_UBound : VbNativeFunction
    {
        public VB_UBound(IVBScopeObject context)
            : base(context, "UBound")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                DArrayExpression arrayExpression;
                int dimension = 1; // Default dimension is 1

                // Retrieve the array
                if (!Converter.TryGetArrayExpression(args[0], out arrayExpression))
                {
                    return DefaultExpression(args);
                }

                // Optionally retrieve the dimension (1-based index)
                if (args.Length > 1 && !Converter.TryGetInt32Value(args[1], out dimension))
                {
                    return DefaultExpression(args);
                }

                if (dimension != 1)
                {
                    throw new ArgumentException("Only single-dimensional arrays are supported.");
                }

                // Return the upper bound, which is the count of items minus one
                int upperBound = arrayExpression.Items.Count - 1;
                return new DMathExpression<int>(upperBound) { HasSideEffet = false };
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }

    public class VB_Join : VbNativeFunction
    {
        public VB_Join(IVBScopeObject context)
            : base(context, "Join")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                DArrayExpression arrayExpr;
                string delimiter = " "; // Default delimiter

                // Retrieve the array
                if (!Converter.TryGetArrayExpression(args[0], out arrayExpr))
                {
                    return DefaultExpression(args);
                }

                // Retrieve the delimiter if provided
                if (args.Length > 1)
                {
                    if (!Converter.TryGetStringValue(args[1], out delimiter))
                    {
                        delimiter = " "; // Default to space if conversion fails
                    }
                }

                // Convert the array items to strings
                List<string> items = new List<string>();
                foreach (var item in arrayExpr.Items)
                {
                    string strItem;
                    if (Converter.TryGetStringValue(item, out strItem))
                    {
                        items.Add(strItem);
                    }
                    else
                    {
                        return DefaultExpression(args);
                        //items.Add(item.ToExpressionString()); // Fallback to the expression string
                    }
                }

                // Join the items with the specified delimiter
                string result = string.Join(delimiter, items);

                return new DSimpleStringExpression(result, Encoding.Unicode, Context.Options);
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }
    public class VB_Split : VbNativeFunction
    {
        public VB_Split(IVBScopeObject context)
            : base(context, "Split")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                string expression;
                string delimiter = " "; // Default delimiter
                int limit = -1; // Default limit (no limit)
                int compare = 0; // Default comparison type (binary)

                // Retrieve the string to split
                if (!Converter.TryGetStringValue(args[0], out expression))
                {
                    return DefaultExpression(args);
                }

                // Retrieve the delimiter if provided
                if (args.Length > 1 && !Converter.TryGetStringValue(args[1], out delimiter))
                {
                    delimiter = " "; // Default to space if conversion fails
                }

                // Retrieve the limit if provided
                if (args.Length > 2 && !Converter.TryGetInt32Value(args[2], out limit))
                {
                    limit = -1; // Default to no limit if conversion fails
                }

                // Retrieve the comparison type if provided
                if (args.Length > 3 && !Converter.TryGetInt32Value(args[3], out compare))
                {
                    compare = 0; // Default to binary compare if conversion fails
                }

                // Adjust the expression and delimiter for case-insensitive comparison if necessary
                if (compare == 1) // Text comparison
                {
                    expression = expression.ToLowerInvariant();
                    delimiter = delimiter.ToLowerInvariant();
                }

                // Split the string
                string[] resultArray = limit > 0
                    ? expression.Split(new string[] { delimiter }, limit, StringSplitOptions.None)
                    : expression.Split(new string[] { delimiter }, StringSplitOptions.None);

                // Convert the result into a DArrayExpression
                List<DExpression> resultExpressions = new List<DExpression>();
                foreach (var item in resultArray)
                {
                    resultExpressions.Add(new DSimpleStringExpression(item, Encoding.Unicode, Context.Options));
                }

                return new DArrayExpression(resultExpressions);
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }



    public class VB_CreateObject
    : VbNativeFunction
    {
        public VB_CreateObject(IVBScopeObject context)
            : base(context, "CreateObject")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
            {
                return DefaultExpression(args);
            }

            if (Context?.Options?.CreateObjectObserver != null)
            {
                Context.Options.CreateObjectObserver.CreateObjectObserved.Add(strArg.Replace("\"\"", "\""));
            }

            return DefaultExpression(args);
        }
    }

    public class VB_Shell
    : VbNativeFunction
    {
        public VB_Shell(IVBScopeObject context)
            : base(context, "Shell")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
            {
                return DefaultExpression(args);
            }

            if (Context?.Options?.ShellObserver != null)
            {
                Context.Options.ShellObserver.ShellObserved.Add(strArg.Replace("\"\"", "\""));
            }

            return DefaultExpression(args);
        }
    }

    public class VB_Execute
    : VbNativeFunction
    {
        public VB_Execute(IVBScopeObject context)
            : base(context, "Execute")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
            {
                return DefaultExpression(args);
            }

            if (Context?.Options?.ExecuteObserver != null)
            {
                Context.Options.ExecuteObserver.VBScriptExecuted.Add(strArg.Replace("\"\"", "\""));
            }

            return DefaultExpression(args);
        }
    }

    public class VB_Replace
    : VbNativeFunction
    {
        public VB_Replace(IVBScopeObject context)
            : base(context, "Replace")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            if (args.Length < 3)
                throw new Exception();

            DExpression exp = args[0];
            DExpression findExp = args[1];
            DExpression replExp = args[2];

            string expStr;
            string findStr;
            string replStr;

            if (!Converter.TryGetStringValue(exp, out expStr))
                return DefaultExpression(args);
            if (!Converter.TryGetStringValue(findExp, out findStr))
                return DefaultExpression(args);
            if (!Converter.TryGetStringValue(replExp, out replStr))
                return DefaultExpression(args);

            if (args.Length > 3)
            {
                DExpression startExp = args[3];
            }

            if (args.Length > 4)
            {
                DExpression countExp = args[4];
            }

            if (args.Length > 5)
            {
                DExpression compareExp = args[5];
            }

            string str = findStr.Equals(replStr) ? expStr : expStr.Replace(findStr, replStr);

            return new DSimpleStringExpression(str, Encoding.Unicode, Context.Options);
        }

    }

    public class VB_Mid_S
     : VB_Mid
    {
        public VB_Mid_S(IVBScopeObject context)
       : base(context, "Mid$")
        {
        }
    }

    public class VB_InStr : VbNativeFunction
    {
        public VB_InStr(IVBScopeObject context)
            : base(context, "InStr")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            int start = 1;
            string string1, string2;
            StringComparison comparisonType = StringComparison.Ordinal;

            // If 4 arguments are passed, the first is the start position and the last is the comparison type
            if (args.Length == 4)
            {
                if (!Converter.TryGetInt32Value(args[0], out start))
                {
                    return DefaultExpression(args);
                }

                if (!Converter.TryGetStringValue(args[1], out string1) ||
                    !Converter.TryGetStringValue(args[2], out string2))
                {
                    return DefaultExpression(args);
                }

                if (!Converter.TryGetInt32Value(args[3], out int comparisonMode))
                {
                    return DefaultExpression(args);
                }

                // Set the comparison type
                comparisonType = comparisonMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            }
            // If 3 arguments are passed, the first can be start or comparison type based on their type
            else if (args.Length == 3)
            {
                if (args[0] is DMathExpression)
                {
                    if (!Converter.TryGetInt32Value(args[0], out start))
                    {
                        return DefaultExpression(args);
                    }

                    if (!Converter.TryGetStringValue(args[1], out string1) ||
                        !Converter.TryGetStringValue(args[2], out string2))
                    {
                        return DefaultExpression(args);
                    }
                }
                else
                {
                    if (!Converter.TryGetStringValue(args[0], out string1) ||
                        !Converter.TryGetStringValue(args[1], out string2))
                    {
                        return DefaultExpression(args);
                    }

                    if (!Converter.TryGetInt32Value(args[2], out int comparisonMode))
                    {
                        return DefaultExpression(args);
                    }

                    comparisonType = comparisonMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
                }
            }
            // If 2 arguments are passed, they are the strings to compare
            else if (args.Length == 2)
            {
                if (!Converter.TryGetStringValue(args[0], out string1) ||
                    !Converter.TryGetStringValue(args[1], out string2))
                {
                    return DefaultExpression(args);
                }
            }
            else
            {
                return DefaultExpression(args);
            }

            // Adjust for 1-based index
            start = start - 1;

            // Perform the search
            int position = string1.IndexOf(string2, start, comparisonType);

            // Convert the result back to 1-based index
            position = position == -1 ? 0 : position + 1;

            return new DMathExpression<int>(position) { HasSideEffet = false };
        }
    }

    public class VB_InStrB : VbNativeFunction
    {
        public VB_InStrB(IVBScopeObject context)
            : base(context, "InStrB")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            int startByte = 1;
            string string1, string2;
            StringComparison comparisonType = StringComparison.Ordinal;

            // Convert strings to byte arrays using Unicode encoding
            byte[] byteArray1, byteArray2;

            // If 4 arguments are passed, the first is the start byte position and the last is the comparison type
            if (args.Length == 4)
            {
                if (!Converter.TryGetInt32Value(args[0], out startByte))
                {
                    return DefaultExpression(args);
                }

                if (!Converter.TryGetStringValue(args[1], out string1) ||
                    !Converter.TryGetStringValue(args[2], out string2))
                {
                    return DefaultExpression(args);
                }

                if (!Converter.TryGetInt32Value(args[3], out int comparisonMode))
                {
                    return DefaultExpression(args);
                }

                // Set the comparison type
                comparisonType = comparisonMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            }
            // If 3 arguments are passed, the first can be start byte position or comparison type based on their type
            else if (args.Length == 3)
            {
                if (args[0] is DMathExpression)
                {
                    if (!Converter.TryGetInt32Value(args[0], out startByte))
                    {
                        return DefaultExpression(args);
                    }

                    if (!Converter.TryGetStringValue(args[1], out string1) ||
                        !Converter.TryGetStringValue(args[2], out string2))
                    {
                        return DefaultExpression(args);
                    }
                }
                else
                {
                    if (!Converter.TryGetStringValue(args[0], out string1) ||
                        !Converter.TryGetStringValue(args[1], out string2))
                    {
                        return DefaultExpression(args);
                    }

                    if (!Converter.TryGetInt32Value(args[2], out int comparisonMode))
                    {
                        return DefaultExpression(args);
                    }

                    comparisonType = comparisonMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
                }
            }
            // If 2 arguments are passed, they are the strings to compare
            else if (args.Length == 2)
            {
                if (!Converter.TryGetStringValue(args[0], out string1) ||
                    !Converter.TryGetStringValue(args[1], out string2))
                {
                    return DefaultExpression(args);
                }
            }
            else
            {
                return DefaultExpression(args);
            }

            // Convert strings to byte arrays
            byteArray1 = Encoding.Unicode.GetBytes(string1);
            byteArray2 = Encoding.Unicode.GetBytes(string2);

            // Adjust for 1-based index (convert to 0-based)
            startByte = (startByte - 1) / 2;  // Dividing by 2 because each character is 2 bytes in Unicode


            // Perform the search within the byte array
            int bytePosition = IndexOfByteArray(byteArray1, byteArray2, startByte);

            // Convert the result back to 1-based index
            bytePosition = bytePosition == -1 ? 0 : bytePosition + 1;

            return new DMathExpression<int>(bytePosition) { HasSideEffet = false };
        }

        private int IndexOfByteArray(byte[] byteArray1, byte[] byteArray2, int startByte)
        {
            for (int i = startByte; i <= byteArray1.Length - byteArray2.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < byteArray2.Length; j++)
                {
                    if (byteArray1[i + j] != byteArray2[j])
                    {
                        match = false;
                        break;
                    }
                }
                if (match) return i;
            }
            return -1;
        }
    }

    public class VB_InStrRev : VbNativeFunction
    {
        public VB_InStrRev(IVBScopeObject context)
            : base(context, "InStrRev")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            string string1, string2;
            int start = -1;  // Start position defaults to the end of the string
            StringComparison comparisonType = StringComparison.Ordinal;

            // If 4 arguments are passed, the first is the string, second is the substring, third is start, and fourth is comparison mode
            if (args.Length == 4)
            {
                if (!Converter.TryGetStringValue(args[0], out string1) ||
                    !Converter.TryGetStringValue(args[1], out string2))
                {
                    return DefaultExpression(args);
                }

                if (!Converter.TryGetInt32Value(args[2], out start))
                {
                    return DefaultExpression(args);
                }

                if (!Converter.TryGetInt32Value(args[3], out int comparisonMode))
                {
                    return DefaultExpression(args);
                }

                // Set the comparison type
                comparisonType = comparisonMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            }
            // If 3 arguments are passed, they can be the strings and start position, or strings and comparison mode
            else if (args.Length == 3)
            {
                if (args[2] is DMathExpression)
                {
                    if (!Converter.TryGetStringValue(args[0], out string1) ||
                        !Converter.TryGetStringValue(args[1], out string2) ||
                        !Converter.TryGetInt32Value(args[2], out start))
                    {
                        return DefaultExpression(args);
                    }
                }
                else
                {
                    if (!Converter.TryGetStringValue(args[0], out string1) ||
                        !Converter.TryGetStringValue(args[1], out string2))
                    {
                        return DefaultExpression(args);
                    }

                    if (!Converter.TryGetInt32Value(args[2], out int comparisonMode))
                    {
                        return DefaultExpression(args);
                    }

                    comparisonType = comparisonMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
                }
            }
            // If 2 arguments are passed, they are the strings to compare
            else if (args.Length == 2)
            {
                if (!Converter.TryGetStringValue(args[0], out string1) ||
                    !Converter.TryGetStringValue(args[1], out string2))
                {
                    return DefaultExpression(args);
                }
            }
            else
            {
                return DefaultExpression(args);
            }

            // If start is not provided, or if it's set to -1, it means start from the end of the string
            if (start == -1 || start > string1.Length)
            {
                start = string1.Length;
            }

            // Adjust for 1-based index by subtracting 1 (since start in C# is 0-based)
            start = start - 1;

            // Perform the reverse search
            int position = string1.LastIndexOf(string2, start, comparisonType);

            // Convert the result back to 1-based index
            position = position == -1 ? 0 : position + 1;

            return new DMathExpression<int>(position) { HasSideEffet = false };
        }
    }

    public class VB_Len : VbNativeFunction
    {
        public VB_Len(IVBScopeObject context)
            : base(context, "Len")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            string strArg;

            // Check if the argument is valid and can be converted to a string
            if (!Converter.TryGetStringValue(args[0], out strArg))
            {
                return DefaultExpression(args);
            }

            // The length in characters is simply the length of the string
            int charLength = strArg.Length;

            return new DMathExpression<int>(charLength) { HasSideEffet = false };
        }
    }

    public class VB_LenB : VbNativeFunction
    {
        public VB_LenB(IVBScopeObject context)
            : base(context, "LenB")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            string strArg;

            // Check if the argument is valid and can be converted to a string
            if (!Converter.TryGetStringValue(args[0], out strArg))
            {
                return DefaultExpression(args);
            }

            // Convert the string to a byte array using Unicode encoding
            byte[] byteArray = Encoding.Unicode.GetBytes(strArg);

            // The length in bytes is simply the length of the byte array
            int byteLength = byteArray.Length;

            return new DMathExpression<int>(byteLength) { HasSideEffet = false };
        }
    }

    public class VB_Mid : VbNativeFunction
    {
        public VB_Mid(IVBScopeObject context, string identifier)
            : base(context, identifier)
        {
        }
        public VB_Mid(IVBScopeObject context)
            : base(context, "Mid")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            string strArg;
            int start, length;

            // Check if the first argument (string) is valid
            if (!Converter.TryGetStringValue(args[0], out strArg))
            {
                return DefaultExpression(args);
            }

            // Check if the second argument (start position) is valid
            if (!Converter.TryGetInt32Value(args[1], out start))
            {
                return DefaultExpression(args);
            }

            // Adjust for 1-based index (VB uses 1-based, C# uses 0-based)
            start = start - 1;

            // Determine if length argument is provided
            if (args.Length > 2 && Converter.TryGetInt32Value(args[2], out length))
            {
                // Ensure length does not exceed the remaining string
                if (start + length > strArg.Length)
                {
                    length = strArg.Length - start;
                }
            }
            else
            {
                // If length is not provided, extract to the end of the string
                length = strArg.Length - start;
            }

            // Check for valid start position and length
            if (start < 0 || start >= strArg.Length || length < 0)
            {
                return new DSimpleStringExpression(string.Empty, Encoding.Unicode, Context.Options);
            }

            string result = strArg.Substring(start, length);
            return new DSimpleStringExpression(result, Encoding.Unicode, Context.Options);
        }
    }


    public class VB_MidB_S
     : VB_MidB
    {
        public VB_MidB_S(IVBScopeObject context)
       : base(context, "MidB$")
        {
        }
    }
    public class VB_MidB : VbNativeFunction
    {
        public VB_MidB(IVBScopeObject context, string identifier)
            : base(context, identifier)
        {
        }

        public VB_MidB(IVBScopeObject context)
            : base(context, "MidB")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            string strArg;
            int startByte, lengthBytes;
            
            // Check if the first argument (string) is valid
            if (!Converter.TryGetStringValue(args[0], out strArg))
            {
                return DefaultExpression(args);
            }

            // Convert the string to a byte array using the appropriate encoding
            byte[] byteArray = Encoding.Unicode.GetBytes(strArg);

            // Check if the second argument (start byte position) is valid
            if (!Converter.TryGetInt32Value(args[1], out startByte))
            {
                return DefaultExpression(args);
            }

            // Adjust for 1-based index (VB uses 1-based, C# uses 0-based)
            startByte = startByte - 1;

            // Determine if length argument is provided
            if (args.Length > 2 && Converter.TryGetInt32Value(args[2], out lengthBytes))
            {
                // Ensure length does not exceed the remaining bytes
                if (startByte + lengthBytes > byteArray.Length)
                {
                    lengthBytes = byteArray.Length - startByte;
                }
            }
            else
            {
                // If length is not provided, extract to the end of the byte array
                lengthBytes = byteArray.Length - startByte;
            }

            // Check for valid start byte position and length
            if (startByte < 0 || startByte >= byteArray.Length || lengthBytes < 0)
            {
                return new DSimpleStringExpression(string.Empty, Encoding.Unicode, Context.Options);
            }

            // Extract the specified byte range
            byte[] resultBytes = byteArray.Skip(startByte).Take(lengthBytes).ToArray();

            // Convert the result back to a string
            string result = Encoding.Default.GetString(resultBytes);
            return new DSimpleStringExpression(result, Encoding.Unicode, Context.Options);
        }
    }



    public class VB_Trim_S
     : VB_Trim
    {
        public VB_Trim_S(IVBScopeObject context)
       : base(context, "Trim$")
        {
        }
    }


    public class VB_LTrim_S
     : VB_LTrim
    {
        public VB_LTrim_S(IVBScopeObject context)
       : base(context, "LTrim$")
        {
        }
    }

    public class VB_RTrim_S
     : VB_RTrim
    {
        public VB_RTrim_S(IVBScopeObject context)
       : base(context, "RTrim$")
        {
        }
    }

    public class VB_LCase_S
        : VB_LCase
    {
        public VB_LCase_S(IVBScopeObject context)
       : base(context, "LCase$")
        {
        }
    }

    public class VB_UCase_S
    : VB_UCase
    {
        public VB_UCase_S(IVBScopeObject context)
       : base(context, "UCase$")
        {
        }
    }

    public class VB_LCase
     : VbNativeFunction
    {
        public VB_LCase(IVBScopeObject context, string name)
            : base(context, name)
        {
        }
        public VB_LCase(IVBScopeObject context)
            : base(context, "LCase")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
                return DefaultExpression(args);

            string str = strArg.ToLowerInvariant();
            return new DSimpleStringExpression(str, Encoding.Unicode, Context.Options);
        }

    }

    public class VB_UCase
     : VbNativeFunction
    {
        public VB_UCase(IVBScopeObject context, string name)
            : base(context, name)
        {
        }
        public VB_UCase(IVBScopeObject context)
            : base(context, "UCase")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
                return DefaultExpression(args);

            string str = strArg.ToUpperInvariant();
            return new DSimpleStringExpression(str, Encoding.Unicode, Context.Options);
        }

    }

    public class VB_Left_S
    : VB_Left
    {
        public VB_Left_S(IVBScopeObject context)
       : base(context, "Left$")
        {
        }
    }

    public class VB_Right_S
    : VB_Right
    {
        public VB_Right_S(IVBScopeObject context)
       : base(context, "Right$")
        {
        }
    }

    public class VB_Left
     : VbNativeFunction
    {
        public VB_Left(IVBScopeObject context, string name)
            : base(context, name)
        {
        }
        public VB_Left(IVBScopeObject context)
            : base(context, "Left")
        {
        }


        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                DExpression arg1 = args.FirstOrDefault();
                DExpression arg2 = args[1];

                string strArg;

                int count;
                if (!Converter.TryGetInt32Value(arg2, out count))
                    return DefaultExpression(args);

                if (count == 0)
                    return new DSimpleStringExpression(string.Empty, Encoding.Unicode, Context.Options);

                if (!Converter.TryGetStringValue(arg1, out strArg))
                    return DefaultExpression(args);


                if (count < 0)
                    return DefaultExpression(args);

                if (count > strArg.Length) 
                    count = strArg.Length;

                string str = strArg.Substring(0, count);
                return new DSimpleStringExpression(str, Encoding.Unicode, Context.Options);
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }

    public class VB_Right
     : VbNativeFunction
    {
        public VB_Right(IVBScopeObject context, string name)
            : base(context, name)
        {
        }
        public VB_Right(IVBScopeObject context)
            : base(context, "Right")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                DExpression arg1 = args.FirstOrDefault();
                DExpression arg2 = args[1];

                string strArg;

                int count;
                if (!Converter.TryGetInt32Value(arg2, out count))
                    return DefaultExpression(args);

                if (count == 0)
                    return new DSimpleStringExpression(string.Empty, Encoding.Unicode, Context.Options);

                if (!Converter.TryGetStringValue(arg1, out strArg))
                    return DefaultExpression(args);

                if (count < 0)
                    return DefaultExpression(args);

                if (count > strArg.Length)
                    count = strArg.Length;

                string str = strArg.Substring(strArg.Length - count);
                return new DSimpleStringExpression(str, Encoding.Unicode, Context.Options);
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }


    public class VB_LeftB_S
        : VB_LeftB
    {
        public VB_LeftB_S(IVBScopeObject context)
       : base(context, "LeftB$")
        {
        }
    }

    public class VB_RightB_S
        : VB_RightB
    {
        public VB_RightB_S(IVBScopeObject context)
       : base(context, "RightB$")
        {
        }
    }

    public class VB_LeftB : VbNativeFunction
    {
        public VB_LeftB(IVBScopeObject context, string identifier)
            : base(context, identifier)
        {
        }

        public VB_LeftB(IVBScopeObject context)
            : base(context, "LeftB")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                DExpression arg1 = args.FirstOrDefault();
                DExpression arg2 = args[1];

                string strArg;

                int byteCount;
                if (!Converter.TryGetInt32Value(arg2, out byteCount))
                    return DefaultExpression(args);

                if (byteCount == 0)
                    return new DSimpleStringExpression(string.Empty, Encoding.Unicode, Context.Options);

                if (!Converter.TryGetStringValue(arg1, out strArg))
                    return DefaultExpression(args);


                if (byteCount < 0)
                    return DefaultExpression(args);

                // Convert the string to a byte array using Unicode encoding
                byte[] byteArray = Encoding.Unicode.GetBytes(strArg);

                // Ensure the byteCount does not exceed the length of the byte array
                if (byteCount > byteArray.Length)
                    byteCount = byteArray.Length;

                // Extract the specified number of bytes
                byte[] resultBytes = byteArray.Take(byteCount).ToArray();

                // Convert the result back to a string
                string result = Encoding.Unicode.GetString(resultBytes);
                return new DSimpleStringExpression(result, Encoding.Unicode, Context.Options);
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }
    public class VB_RightB : VbNativeFunction
    {
        public VB_RightB(IVBScopeObject context, string identifier)
            : base(context, identifier)
        {
        }

        public VB_RightB(IVBScopeObject context)
            : base(context, "RightB")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            try
            {
                DExpression arg1 = args.FirstOrDefault();
                DExpression arg2 = args[1];

                string strArg;

                int byteCount;
                if (!Converter.TryGetInt32Value(arg2, out byteCount))
                    return DefaultExpression(args);


                if (byteCount == 0)
                    return new DSimpleStringExpression(string.Empty, Encoding.Unicode, Context.Options);

                if (!Converter.TryGetStringValue(arg1, out strArg))
                    return DefaultExpression(args);


                if (byteCount < 0)
                    return DefaultExpression(args);

                // Convert the string to a byte array using Unicode encoding
                byte[] byteArray = Encoding.Unicode.GetBytes(strArg);

                // Ensure the byteCount does not exceed the length of the byte array
                if (byteCount > byteArray.Length)
                    byteCount = byteArray.Length;

                // Extract the specified number of bytes from the end of the byte array
                byte[] resultBytes = byteArray.Skip(byteArray.Length - byteCount).Take(byteCount).ToArray();

                // Convert the result back to a string
                string result = Encoding.Unicode.GetString(resultBytes);
                return new DSimpleStringExpression(result, Encoding.Unicode, Context.Options);
            }
            catch (Exception ex)
            {
                return DefaultExpression(args);
            }
        }
    }


    public class VB_Trim
     : VbNativeFunction
    {
        public VB_Trim(IVBScopeObject context, string name)
            : base(context, name)
        {
        }
        public VB_Trim(IVBScopeObject context)
            : base(context, "Trim")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
                return DefaultExpression(args);

            string str = strArg.Trim(' ');
            return new DSimpleStringExpression(str, Encoding.Unicode, Context.Options);
        }
    }


    public class VB_LTrim
     : VbNativeFunction
    {
        public VB_LTrim(IVBScopeObject context, string name)
            : base(context, name)
        {
        }
        public VB_LTrim(IVBScopeObject context)
            : base(context, "LTrim")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
                return DefaultExpression(args);

            string str = strArg.TrimStart(' ');
            return new DSimpleStringExpression(str, Encoding.Unicode, Context.Options);
        }
    }
    public class VB_RTrim
        : VbNativeFunction
    {
        public VB_RTrim(IVBScopeObject context, string name)
            : base(context, name)
        {
        }
        public VB_RTrim(IVBScopeObject context)
            : base(context, "RTrim")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string strArg;

            if (!Converter.TryGetStringValue(arg1, out strArg))
                return DefaultExpression(args);

            string str = strArg.TrimEnd(' ');
            return new DSimpleStringExpression(str, Encoding.Unicode, Context.Options);
        }
    }

    public class VB_Space_S : VB_Space
    {
        public VB_Space_S(IVBScopeObject context)
            : base(context, "Space$")
        {
        }

    }

    public class VB_Space
    : VbNativeFunction
    {
        public VB_Space(IVBScopeObject context, string name)
            : base(context, name)
        {
        }

        public VB_Space(IVBScopeObject context)
            : base(context, "Space")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            int count;
            if (!Converter.TryGetInt32Value(arg1, out count))
                return DefaultExpression(args);
            

            string value = new string(' ', count);

            return new DSimpleStringExpression(value, Encoding.Unicode, Context.Options);
        }

    }

    public class VB_ChrW 
        : VbNativeFunction
    {
        public VB_ChrW(IVBScopeObject context, string name)
            : base(context, name)
        {
        }

        public VB_ChrW(IVBScopeObject context) 
            : base(context, "ChrW")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            int ascii;

            if (!Converter.TryGetInt32Value(arg1, out ascii))
                return DefaultExpression(args);

            char test = (char)ascii;
            string value = new string(new char[] { test });

            //string value = Char.ConvertFromUtf32((int)ascii); //(byte) (UInt32)Math.Round(ascii) & 0x0000FFFF);

            return new DSimpleStringExpression(value, Encoding.Unicode, Context.Options);
        }

    }

    public class VB_ChrB_S
        : VB_ChrB
    {
        public VB_ChrB_S(IVBScopeObject context)
            : base(context, "ChrB$")
        {
        }
    }

        public class VB_ChrB
    : VbNativeFunction
    {
        public VB_ChrB(IVBScopeObject context, string name)
            : base(context, name)
        {
        }
        public VB_ChrB(IVBScopeObject context)
            : base(context, "ChrB")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            int ascii;

            if (!Converter.TryGetInt32Value(arg1, out ascii))
                return DefaultExpression(args);

            //byte[] test = { (byte) ascii };
            ////throw new Exception("Not managed !");
            //string value = Encoding.ASCII.GetString(test);
            string value = new string(new char[] { VbUtils.Chr(ascii) });

            return new DSimpleStringExpression(value, Encoding.Unicode, Context.Options);
        }
    }

    public class VB_Chr_S
        : VB_Chr
    {
        public VB_Chr_S(IVBScopeObject context)
            : base(context, "Chr$")
        {
        }
    }

    public class VB_Chr 
        : VbNativeFunction
    {
        public VB_Chr(IVBScopeObject context, string identifier)
            : base(context, identifier)
        {
        }

        public VB_Chr(IVBScopeObject context)
            : base(context, "Chr")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            int ascii;

            if (!Converter.TryGetInt32Value(arg1, out ascii))
                return DefaultExpression(args);

            string value = new string(new char[]{ VbUtils.Chr(ascii) }); 
            
            return new DSimpleStringExpression(value, Encoding.Unicode, Context.Options);
        }

    }


    public class VB_Abs
    : VbNativeFunction
    {
        public VB_Abs(IVBScopeObject context)
            : base(context, "Abs")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            double inputVal;

            if (!Converter.TryGetDoubleValue(arg1, out inputVal))
                return DefaultExpression(args);

            double value = Math.Abs(inputVal);

            return new DMathExpression<double>(value) { HasSideEffet = false };
        }
    }
    public class VB_Sgn
        : VbNativeFunction
    {
        public VB_Sgn(IVBScopeObject context)
            : base(context, "Sgn")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            double inputVal;

            if (!Converter.TryGetDoubleValue(arg1, out inputVal))
                return DefaultExpression(args);

            double value = Math.Sign(inputVal);

            return new DMathExpression<double>(value) { HasSideEffet = false };
        }
    }
    public class VB_Sqr
        : VbNativeFunction
    {
        public VB_Sqr(IVBScopeObject context)
            : base(context, "Sqr")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            double inputVal;

            if (!Converter.TryGetDoubleValue(arg1, out inputVal))
                return DefaultExpression(args);
            
            double value = Math.Sqrt(inputVal);

            return new DMathExpression<double>(value) { HasSideEffet = false };
        }
    }


    public class VB_Cos
    : VbNativeFunction
    {
        public VB_Cos(IVBScopeObject context)
            : base(context, "Cos")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            double input;
            if (!Converter.TryGetDoubleValue(arg1, out input))
                return DefaultExpression(args);
            
            double value = Math.Cos(input);

            return new DMathExpression<double>(value) { HasSideEffet = false };
        }
    }

    public class VB_CInt
        : VbNativeFunction
    {
        public VB_CInt(IVBScopeObject context)
            : base(context, "CInt")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            Int16 input;
            if (!Converter.TryGetInt16Value(arg1, out input))
                return DefaultExpression(args);


            return new DMathExpression<Int16>(input) { HasSideEffet = false };
        }
    }

    public class VB_CBool
        : VbNativeFunction
    {
        public VB_CBool(IVBScopeObject context)
            : base(context, "CBool")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            bool input;
            if (!Converter.TryGetBool(arg1, out input))
                return DefaultExpression(args);

            return new DBoolExpression(input) { HasSideEffet = false };
        }
    }

    public class VB_Exp
    : VbNativeFunction
    {
        public VB_Exp(IVBScopeObject context)
            : base(context, "Exp")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            double inputVal;

            if (!Converter.TryGetDoubleValue(arg1, out inputVal))
                return DefaultExpression(args);

            double value = Math.Exp(inputVal);

            return new DMathExpression<double>(value) { HasSideEffet = false };
        }
    }


    public class VB_Round
    : VbNativeFunction
    {
        public VB_Round(IVBScopeObject context)
            : base(context, "Round")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            return Evaluate(
                args.FirstOrDefault(),
                args.Skip(1).FirstOrDefault(), 
                args);
        }

        public DExpression Evaluate(DExpression number, DExpression numDecimal, DExpression[] args)
        {
            double numberVal;

            if (!Converter.TryGetDoubleValue(number, out numberVal))
                return DefaultExpression(args);

            int numDecimalVal;

            if (numDecimal != null)
            {
                if (!Converter.TryGetInt32Value(numDecimal, out numDecimalVal))
                    return DefaultExpression(args);


                double ret = Math.Round(numberVal, numDecimalVal);

                return new DMathExpression<double>(ret) { HasSideEffet = false };
            }

            double ret2 = Math.Round(numberVal);

            return new DMathExpression<double>(ret2) { HasSideEffet = false };
        }
    }


    public class VB_Rnd
    : VbNativeFunction
    {
        public VB_Rnd(IVBScopeObject context)
            : base(context, "Rnd")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            //DExpression arg1 = args.FirstOrDefault();

            return DefaultExpression(args);
        }
    }

    public class VB_Sin
    : VbNativeFunction
    {
        public VB_Sin(IVBScopeObject context)
            : base(context, "Sin")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            double input;
            if (!Converter.TryGetDoubleValue(arg1, out input))
                return DefaultExpression(args);


            double value = Math.Sin(input);

            return new DMathExpression<double>(value) { HasSideEffet = false };
        }
    }

    public class VB_Tan
        : VbNativeFunction
    {
        public VB_Tan(IVBScopeObject context)
            : base(context, "Tan")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            double input;
            if (!Converter.TryGetDoubleValue(arg1, out input))
                return DefaultExpression(args);


            double value = Math.Tan(input);

            return new DMathExpression<double>(value) { HasSideEffet = false };
        }
    }


    public class VB_Atn
    : VbNativeFunction
    {
        public VB_Atn(IVBScopeObject context)
            : base(context, "Atn")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            double input;
            if (!Converter.TryGetDoubleValue(arg1, out input))
                return DefaultExpression(args);


            double value = Math.Atan(input);

            return new DMathExpression<double>(value) { HasSideEffet = false };
        }

    }

    public class VB_Asc 
        : VbNativeFunction
    {
        public VB_Asc(IVBScopeObject context)
            : base(context, "Asc")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string chrInput;

            if (!Converter.TryGetStringValue(arg1, out chrInput))
                return DefaultExpression(args);

            int value = VbUtils.Asc(chrInput);

            return new DMathExpression<int>(value) { HasSideEffet = false };
        }

    }
    public class VB_AscB
        : VbNativeFunction
    {
        public VB_AscB(IVBScopeObject context)
            : base(context, "AscB")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string chrInput;

            if (!Converter.TryGetStringValue(arg1, out chrInput))
                return DefaultExpression(args);

            byte value = (byte)(int)(chrInput[0] & 0x000000FF);

            return new DMathExpression<byte>(value) { HasSideEffet = false };
        }

    }
    public class VB_AscW
    : VbNativeFunction
    {
        public VB_AscW(IVBScopeObject context)
            : base(context, "AscW")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            string chrInput;

            if (!Converter.TryGetStringValue(arg1, out chrInput))
                return DefaultExpression(args);

            int value = chrInput[0];

            return new DMathExpression<int>(value) { HasSideEffet = false };
        }

    }

    public class VB_Hex_S
        : VB_Hex
    {
        public VB_Hex_S(IVBScopeObject context)
            : base(context, "Hex$")
        {
        }
    }

    public class VB_Hex
    : VbNativeFunction
    {
        public VB_Hex(IVBScopeObject context, string name)
            : base(context, name)
        {
        }
        public VB_Hex(IVBScopeObject context)
            : base(context, "Hex")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            long input;
            if (!Converter.TryGetInt64Value(arg1, out input))
                return DefaultExpression(args);


            string hexStr = $"{input:X}";

            return new DSimpleStringExpression(hexStr, null, Context.Options);
        }
    }

    public class VB_Log 
        : VbNativeFunction
    {
        public VB_Log(IVBScopeObject context)
            : base(context, "Log")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            double input;
            if (!Converter.TryGetDoubleValue(arg1, out input))
                return DefaultExpression(args);

            return new DMathExpression<double>(Math.Log(input)) { HasSideEffet = false };
        }

    }


    public class VB_Val
    : VbNativeFunction
    {
        public VB_Val(IVBScopeObject context)
            : base(context, "Val")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            double? res = null;
            if (arg1 is IStringExpression)
            {
                string str = arg1.ToValueString();

                try
                {
                    bool hadValError;
                    res = VbUtils.Val(str, false, out hadValError);
                } 
                catch
                {
                    // If engine is VBA, return 0, if engine is vbscript, should throw exception
                    res = 0;
                }
            }

            if (arg1 is DMathExpression)
            {
                DMathExpression arg1_c = (DMathExpression)arg1;

                res = Convert.ToDouble(arg1_c.GetValueObject());
            }

            if (!res.HasValue)
                return DefaultExpression(args);

            return new DMathExpression<double>(res) { HasSideEffet = false };
        }

    }

    public class VB_CLng
    : VbNativeFunction
    {
        public VB_CLng(IVBScopeObject context)
            : base(context, "CLng")
        {
        }

        public override DExpression Evaluate(params DExpression[] args)
        {
            DExpression arg1 = args.FirstOrDefault();

            Int32 input;
            if (!Converter.TryGetInt32Value(arg1, out input))
                return DefaultExpression(args);

            return new DMathExpression<Int32>(input) { HasSideEffet = false };
        }

    }

    //public class VB_Execute
    //    : VbNativeFunction
    //{
    //    public VB_Execute(IVBScopeObject context)
    //        : base(context, "Execute")
    //    {
    //    }

    //    public override DExpression Evaluate(params DExpression[] args)
    //    {
    //        DExpression arg1 = args.FirstOrDefault();

    //        Console.WriteLine($"{Name}({arg1.ToExpressionString()})");

    //        return new DCodeBlock($"{Name}({arg1.ToExpressionString()})");
    //    }

    //}

    //public class VB_MsgBox 
    //    : VbNativeFunction
    //{
    //    public VB_MsgBox(IVBScopeObject context)
    //        : base(context, "MsgBox")
    //    {
    //    }

    //    public override DExpression Evaluate(params DExpression[] args)
    //    {
    //        DExpression arg1 = args.FirstOrDefault();

    //        Console.WriteLine($"{Name}({arg1.ToExpressionString()})");

    //        return new DCodeBlock($"{Name}({arg1.ToExpressionString()})");
    //    }

    //}

    public static class Converter
    {
        internal static DSimpleStringExpression ToStringExp(DExpression arg1)
        {
            throw new NotImplementedException();
        }

        internal static bool TryGetInt16Value(DExpression arg1, out Int16 output)
        {
            double dbl;

            if (TryGetSingleValue(arg1, out dbl))
            {
                output = (Int16)Math.Round(dbl, 0);
                return true;
            }

            output = 0;
            return false;
        }

        internal static bool TryGetInt32Value(DExpression arg1, out Int32 output)
        {
            double dbl;

            if (TryGetDoubleValue(arg1, out dbl))
            {
                output = (int)Math.Round(dbl, 0);
                return true;
            }

            output = 0;
            return false;
        }

        internal static bool TryGetInt64Value(DExpression arg1, out Int64 output)
        {
            double dbl;

            if (TryGetDoubleValue(arg1, out dbl))
            {
                output = (Int64)Math.Round(dbl, 0);
                return true;
            }

            output = 0;
            return false;
        }

        internal static bool TryGetDoubleValue(DExpression arg1, out double doubleArg)
        {
            if (arg1 is IStringExpression)
            {
                try
                {
                    if (!arg1.IsValuable)
                    {
                        doubleArg = 0;
                        return false;
                    }

                    string str = arg1.ToValueString();

                    bool hadValError;
                    double? val = VbUtils.Val(str, false, out hadValError);

                    if (val.HasValue && !hadValError)
                    {
                        doubleArg = val.Value;
                        return true;
                    }
                }
                catch
                {
                    doubleArg = 0;
                    return false;
                }
            }

            if (arg1 is DMathExpression)
            {
                DMathExpression arg1_c = (DMathExpression)arg1;

                doubleArg = Convert.ToDouble(arg1_c.GetValueObject());
                return true;
            }


            doubleArg = 0;
            return false;
        }

        internal static bool TryGetSingleValue(DExpression arg1, out double doubleArg)
        {
            if (arg1 is IStringExpression)
            {
                string str = arg1.ToValueString();

                try
                {
                    bool hadValError;
                    float? val = (float)VbUtils.Val(str, true, out hadValError);

                    if (val.HasValue && !hadValError)
                    {
                        doubleArg = val.Value;
                        return true;
                    }
                }
                catch
                {
                    doubleArg = 0;
                    return false;
                }
            }

            if (arg1 is DMathExpression)
            {
                DMathExpression arg1_c = (DMathExpression)arg1;

                doubleArg = Convert.ToDouble(arg1_c.GetValueObject());
                return true;
            }


            doubleArg = 0;
            return false;
        }


        internal static bool TryGetStringValue(DExpression arg1, out string strArg)
        {
            if (arg1 is DCodeBlock)
            {
                strArg = null;
                return false;
            }

            if (arg1 is IStringExpression)
            {
                if (arg1.IsValuable)
                {
                    strArg = arg1.ToValueString();
                    return true;
                }
            }

            if (arg1 is DMathExpression)
            {
                if (arg1.IsValuable)
                {
                    strArg = arg1.ToValueString();
                    return true;
                }
            }


            if (arg1 is DEmptyVariable)
            {
                strArg = string.Empty;
                return true;
            }

            if (arg1 is DUndefinedVariable)
            {
                strArg = null;
                return false;
            }

            strArg = null;
            return false;
        }

        internal static bool TryGetBool(DExpression arg1, out bool output)
        {
            string outStr;
            if (TryGetStringValue(arg1, out outStr))
            {
                if (outStr.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    output = false;
                    return true;
                }
                if (outStr.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    output = true;
                    return true;
                }
            }


            double dbl;

            if (TryGetSingleValue(arg1, out dbl))
            {
                output = dbl != 0;
                return true;
            }

            output = false;
            return false;
        }


        public static bool TryGetArrayExpression(DExpression expr, out DArrayExpression arrayExpression)
        {
            if (expr is DArrayExpression)
            {
                arrayExpression = (DArrayExpression)expr;
                return true;
            }
            else
            {
                arrayExpression = null;
                return false;
            }
        }
    }
}
