using System;

namespace vbSparkle.NativeMethods
{
    public class Financial
    {

        //
        // Summary:
        //     Indicates when payments are due when calling financial methods.
        public enum DueDate
        {
            //
            // Summary:
            //     Falls at the end of the date interval.
            EndOfPeriod = 0,
            //
            // Summary:
            //     Falls at the beginning of the date interval.
            BegOfPeriod = 1
        }

        public static double CalculateIRR(double[] values, double guess)
        {
            const double tolerance = 1e-10; // Mimicking VBA's precision
            const int maxIterations = 20;   // VBA limits iterations to 20
            double irr = guess;

            for (int i = 0; i < maxIterations; i++)
            {
                double npv = 0.0;
                double npvDerivative = 0.0;

                for (int t = 0; t < values.Length; t++)
                {
                    double denominator = Math.Pow(1.0 + irr, t);
                    npv += values[t] / denominator;
                    npvDerivative -= t * values[t] / (denominator * (1.0 + irr));
                }

                double newIrr = irr - npv / npvDerivative;

                if (Math.Abs(newIrr - irr) < tolerance)
                {
                    return Financial.TruncateToPrecision(newIrr, 15); // Truncate to match VBA precision
                }

                irr = newIrr;
            }

            throw new Exception("IRR did not converge");
        }

        public static double TruncateToPrecision(double value, int decimalPlaces)
        {
            double factor = Math.Pow(10, decimalPlaces);
            return Math.Truncate(value * factor) / factor;
        }

        public static double MIRR(ref double[] ValueArray, double FinanceRate, double ReinvestRate)
        {
            if (ValueArray.Rank != 1)
            {
                throw new InvalidOperationException();
            }
            int num7 = 0;
            int num6 = (ValueArray.GetUpperBound(0) - num7) + 1;
            if (FinanceRate == -1.0)
            {
                throw new InvalidOperationException();
            }
            if (ReinvestRate == -1.0)
            {
                throw new InvalidOperationException();
            }
            if (num6 <= 1)
            {
                throw new InvalidOperationException();
            }
            double num = LDoNPV(FinanceRate, ref ValueArray, -1);
            if (num == 0.0)
            {
                throw new InvalidOperationException();
            }
            double num2 = LDoNPV(ReinvestRate, ref ValueArray, 1);
            double x = ReinvestRate + 1.0;
            double y = num6;
            double num4 = (-num2 * Math.Pow(x, y)) / (num * (FinanceRate + 1.0));
            if (num4 < 0.0)
            {
                throw new InvalidOperationException();
            }
            x = 1.0 / (num6 - 1.0);
            return (Math.Pow(num4, x) - 1.0);
        }

        private static double LDoNPV(double Rate, ref double[] ValueArray, int iWNType)
        {
            bool flag2 = iWNType < 0;
            bool flag = iWNType > 0;
            double num = 1.0;
            double num2 = 0.0;
            int num6 = 0;
            int num8 = ValueArray.GetUpperBound(0);
            for (int i = num6; i <= num8; i++)
            {
                double num3 = ValueArray[i];
                num += num * Rate;
                if ((!flag2 || (num3 <= 0.0)) && (!flag || (num3 >= 0.0)))
                {
                    num2 += num3 / num;
                }
            }
            return num2;
        }
        public static double FV(double Rate, double NPer, double Pmt, double PV = 0.0, DueDate Due = 0)
        {
            return FV_Internal(Rate, NPer, Pmt, PV, Due);
        }

        private static double FV_Internal(double Rate, double NPer, double Pmt, double PV = 0.0, DueDate Due = 0)
        {
            double num;
            if (Rate == 0.0)
            {
                return (-PV - (Pmt * NPer));
            }
            if (Due != DueDate.EndOfPeriod)
            {
                num = 1.0 + Rate;
            }
            else
            {
                num = 1.0;
            }
            double x = 1.0 + Rate;
            double num2 = Math.Pow(x, NPer);
            return ((-PV * num2) - (((Pmt / Rate) * num) * (num2 - 1.0)));
        }


        public static double IPmt(double Rate, double Per, double NPer, double PV, double FV = 0.0, DueDate Due = 0)
        {
            double num;
            if (Due != DueDate.EndOfPeriod)
            {
                num = 2.0;
            }
            else
            {
                num = 1.0;
            }
            if ((Per <= 0.0) || (Per >= (NPer + 1.0)))
            {
                throw new InvalidOperationException();
            }
            if ((Due != DueDate.EndOfPeriod) && (Per == 1.0))
            {
                return 0.0;
            }
            double pmt = PMT_Internal(Rate, NPer, PV, FV, Due);
            if (Due != DueDate.EndOfPeriod)
            {
                PV += pmt;
            }
            return (FV_Internal(Rate, Per - num, pmt, PV, DueDate.EndOfPeriod) * Rate);
        }

        public static double Pmt(double Rate, double NPer, double PV, double FV = 0.0, DueDate Due = 0)
        {
            return PMT_Internal(Rate, NPer, PV, FV, Due);
        }

        private static double PMT_Internal(double Rate, double NPer, double PV, double FV = 0.0, DueDate Due = 0)
        {
            double num;
            if (NPer == 0.0)
            {
                throw new InvalidOperationException();
            }
            if (Rate == 0.0)
            {
                return ((-FV - PV) / NPer);
            }
            if (Due != DueDate.EndOfPeriod)
            {
                num = 1.0 + Rate;
            }
            else
            {
                num = 1.0;
            }
            double x = Rate + 1.0;
            double num2 = Math.Pow(x, NPer);
            return (((-FV - (PV * num2)) / (num * (num2 - 1.0))) * Rate);
        }

        public static double PPmt(double Rate, double Per, double NPer, double PV, double FV = 0.0, DueDate Due = 0)
        {
            if ((Per <= 0.0) || (Per >= (NPer + 1.0)))
            {
                throw new InvalidOperationException();
            }
            double num2 = PMT_Internal(Rate, NPer, PV, FV, Due);
            double num = IPmt(Rate, Per, NPer, PV, FV, Due);
            return (num2 - num);
        }
    }
}
