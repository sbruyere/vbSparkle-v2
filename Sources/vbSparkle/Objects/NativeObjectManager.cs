using System.Collections.Generic;
using System.Text;
using vbSparkle.EvaluationObjects;
using vbSparkle.Options;

namespace vbSparkle
{
    public class NativeObjectManager : IVBScopeObject
    {

        public Dictionary<string, VbNativeIdentifiedObject> NativeObjects { get; private set; } =
            new Dictionary<string, VbNativeIdentifiedObject>();

        public VbAnalyser Analyser { get; set; }
        public EvaluatorOptions Options { get; set; }

        public NativeObjectManager()
        {
            Add(new VbNativeConstants(this, "vbCrLf",           new DSimpleStringExpression("\r\n", Encoding.Unicode, this.Options)));
            Add(new VbNativeConstants(this, "vbNewLine",        new DSimpleStringExpression("\r\n", Encoding.Unicode, this.Options)));
            Add(new VbNativeConstants(this, "vbCr",             new DSimpleStringExpression("\r", Encoding.Unicode, this.Options)));
            Add(new VbNativeConstants(this, "vbLf",             new DSimpleStringExpression("\n", Encoding.Unicode, this.Options)));
            Add(new VbNativeConstants(this, "vbTab",            new DSimpleStringExpression("\x9", Encoding.Unicode, this.Options)));
            Add(new VbNativeConstants(this, "vbBack",           new DSimpleStringExpression("\x8", Encoding.Unicode, this.Options)));
            Add(new VbNativeConstants(this, "vbNullChar",       new DSimpleStringExpression("\x0", Encoding.Unicode, this.Options)));
            Add(new VbNativeConstants(this, "vbFormFeed",       new DSimpleStringExpression("\xC", Encoding.Unicode, this.Options)));
            Add(new VbNativeConstants(this, "vbVerticalTab",    new DSimpleStringExpression("\xB", Encoding.Unicode, this.Options)));
            Add(new VbNativeConstants(this, "vbNullString",     new DSimpleStringExpression("", Encoding.Unicode, this.Options)));
            Add(new VbNativeConstants(this, "vbObjectError",    new DMathExpression<int>(-0x7FFC0000)));

            Add(new VbNativeConstants(this, "vbUseCompareOption", new DMathExpression<int>(-1)));
            Add(new VbNativeConstants(this, "vbBinaryCompare", new DMathExpression<int>(0)));
            Add(new VbNativeConstants(this, "vbTextCompare", new DMathExpression<int>(1)));
            Add(new VbNativeConstants(this, "vbDatabaseCompare", new DMathExpression<int>(2)));

            // Strings
            Add(new NativeMethods.VB_Chr(this));
            Add(new NativeMethods.VB_Chr_S(this));
            Add(new NativeMethods.VB_ChrW(this));
            Add(new NativeMethods.VB_ChrW(this, "ChrW$"));
            Add(new NativeMethods.VB_ChrB(this));
            Add(new NativeMethods.VB_ChrB_S(this));
            Add(new NativeMethods.VB_Asc(this));
            Add(new NativeMethods.VB_AscW(this));
            Add(new NativeMethods.VB_AscB(this));

            Add(new NativeMethods.VB_MonitoringFunction(this, "Filter"));           // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "Format"));           // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "Format$"));          // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "FormatCurrency"));   // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "FormatDateTime"));   // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "FormatNumber"));     // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "FormatPercent"));    // TODO

            Add(new NativeMethods.VB_InStr(this));
            Add(new NativeMethods.VB_InStrB(this));
            Add(new NativeMethods.VB_InStrRev(this));
            Add(new NativeMethods.VB_LCase(this));
            Add(new NativeMethods.VB_LCase_S(this));
            Add(new NativeMethods.VB_UCase(this));
            Add(new NativeMethods.VB_UCase_S(this));
            Add(new NativeMethods.VB_Len(this));
            Add(new NativeMethods.VB_LenB(this));
            Add(new NativeMethods.VB_Mid(this));
            Add(new NativeMethods.VB_Mid_S(this));
            Add(new NativeMethods.VB_MidB(this));
            Add(new NativeMethods.VB_MidB_S(this));
            Add(new NativeMethods.VB_Replace(this));

            Add(new NativeMethods.VB_Left(this));
            Add(new NativeMethods.VB_Left_S(this));
            Add(new NativeMethods.VB_LeftB(this));
            Add(new NativeMethods.VB_LeftB_S(this));

            Add(new NativeMethods.VB_Right(this));
            Add(new NativeMethods.VB_Right_S(this));
            Add(new NativeMethods.VB_RightB(this));
            Add(new NativeMethods.VB_RightB_S(this));

            Add(new NativeMethods.VB_Space_S(this));
            Add(new NativeMethods.VB_Space(this));

            Add(new NativeMethods.VB_Join(this));
            Add(new NativeMethods.VB_Split(this));

            Add(new NativeMethods.VB_MonitoringFunction(this, "StrComp"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "StrConv"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "String"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "String$"));
            Add(new NativeMethods.VB_StrReverse(this));

            // Triming
            Add(new NativeMethods.VB_Trim(this));
            Add(new NativeMethods.VB_Trim_S(this));
            Add(new NativeMethods.VB_RTrim(this));
            Add(new NativeMethods.VB_RTrim_S(this));
            Add(new NativeMethods.VB_LTrim(this));
            Add(new NativeMethods.VB_LTrim_S(this));

            // Math
            Add(new NativeMethods.VB_Abs(this));
            Add(new NativeMethods.VB_Atn(this));
            Add(new NativeMethods.VB_Cos(this));
            Add(new NativeMethods.VB_Exp(this));
            Add(new NativeMethods.VB_Log(this));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Randomize"));
            Add(new NativeMethods.VB_Rnd(this));
            Add(new NativeMethods.VB_Round(this));
            Add(new NativeMethods.VB_Sgn(this));
            Add(new NativeMethods.VB_Sin(this));
            Add(new NativeMethods.VB_Sqr(this));
            Add(new NativeMethods.VB_Tan(this));


            // Interaction
            Add(new NativeMethods.VB_MonitoringFunction(this, "AppActivate"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Beep"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "CallByName"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Choose"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Command"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Command$"));
            Add(new NativeMethods.VB_CreateObject(this));
            Add(new NativeMethods.VB_MonitoringFunction(this, "DeleteSetting"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "DoEvents"));
            Add(new NativeMethods.VB_Environ(this));
            Add(new NativeMethods.VB_EnvironS(this));
            Add(new NativeMethods.VB_Execute(this));
            Add(new NativeMethods.VB_MonitoringFunction(this, "GetAllSettings"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "GetObject"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "GetSetting"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "IIf"));              // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "InputBox"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "MsgBox"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Partition"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "SaveSetting"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "SendKeys"));
            Add(new NativeMethods.VB_Shell(this));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Switch"));           // TODO

            // Information
            Add(new NativeMethods.VB_MonitoringFunction(this, "Err"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "IMEStatus"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "IsArray"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "IsDate"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "IsEmpty"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "IsError"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "IsMissing"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "IsNull"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "IsNumeric"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "IsObject"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "QBColor"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "RGB"));              // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "TypeName"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "VarType"));

            // Financial
            Add(new NativeMethods.VB_DDB(this));
            Add(new NativeMethods.VB_FV(this));
            Add(new NativeMethods.VB_IPmt(this));
            Add(new NativeMethods.VB_IRR(this));
            Add(new NativeMethods.VB_MIRR(this));
            Add(new NativeMethods.VB_NPer(this));
            Add(new NativeMethods.VB_NPV(this));
            Add(new NativeMethods.VB_Pmt(this));
            Add(new NativeMethods.VB_PPmt(this));
            Add(new NativeMethods.VB_PV(this));
            Add(new NativeMethods.VB_Rate(this));
            Add(new NativeMethods.VB_SLN(this));
            Add(new NativeMethods.VB_SYD(this));

            // Arrays
            Add(new NativeMethods.VB_Array(this));
            Add(new NativeMethods.VB_LBound(this));
            Add(new NativeMethods.VB_UBound(this));

            // DateTime

            Add(new NativeMethods.VB_Time(this));
            Add(new NativeMethods.VB_MonitoringFunction(this, "DateAdd"));          // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "DateDiff"));         // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "DatePart"));         // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "DateSerial"));       // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "DateValue"));        // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "Day"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Hour"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Minute"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Month"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Second"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "TimeSerial"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "TimeValue"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "WeekDay"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Year"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "WeekdayName"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "MonthName"));

            // Conversion
            Add(new NativeMethods.VB_CBool(this));
            Add(new NativeMethods.VB_MonitoringFunction(this, "CByte"));            // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "CCur"));             // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "CDate"));            // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "CDbl"));             // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "CDec"));             // TODO
            Add(new NativeMethods.VB_CInt(this));
            Add(new NativeMethods.VB_CLng(this));
            Add(new NativeMethods.VB_MonitoringFunction(this, "CLngLng"));          // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "CLngPtr"));          // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "CSng"));             // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "CStr"));             // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "CVar"));             // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "CVDate"));           // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "CVErr"));            // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "Error"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Error$"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Fix"));              // TODO
            Add(new NativeMethods.VB_Hex(this));
            Add(new NativeMethods.VB_Hex_S(this));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Int"));              // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "Oct"));              // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "Oct$"));             // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "Str"));              // TODO
            Add(new NativeMethods.VB_MonitoringFunction(this, "Str$"));             // TODO
            Add(new NativeMethods.VB_Val(this));

            // Non Deterministics

            // FileSystem
            Add(new NativeMethods.VB_MonitoringFunction(this, "CurDir"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "CurDir$"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Dir"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "EOF"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "FileAttr"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "FileCopy"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "FileDateTime"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "FileLen"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "FreeFile"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "GetAttr"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Kill"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Loc"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "LOF"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "MkDir"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Reset"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "RmDir"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "Seek"));
            Add(new NativeMethods.VB_MonitoringFunction(this, "SetAttr"));
            // Specials
            Add(new NativeMethods.VB_Eval(this));
        }


        private VbNativeFunction Add(VbNativeFunction wrapper)
        {
            var res = wrapper;
            NativeObjects.Add(wrapper.Identifier.ToUpper(), res);
            return res;
        }
        private VbNativeConstants Add(VbNativeConstants wrapper)
        {
            var res = wrapper;
            NativeObjects.Add(wrapper.Identifier.ToUpper(), res);
            return res;
        }

        public VbIdentifiedObject GetIdentifiedObject(string identifier)
        {
            if (identifier == "Execute")
                (0).ToString();

            if (NativeObjects.ContainsKey(identifier.ToUpper()))
            {
                return NativeObjects[identifier.ToUpper()];
            }

            return null;
        }

        public void DeclareVariable(VbUserVariable variable)
        {
            throw new System.NotImplementedException();
        }

        public void SetVarValue(string dest, DExpression value)
        {
            throw new System.NotImplementedException();
        }

        public void DeclareConstant(VbSubConstStatement vbSubConstStatement)
        {
            throw new System.NotImplementedException();
        }


        //public bool TryGet(string identifier, VbNativeIdentifiedObject nativeObject)
        //{
        //    if (NativeObjects.ContainsKey(identifier.ToUpper()))
        //    {
        //        nativeObject = NativeObjects[identifier.ToUpper()];
        //        return true;
        //    }

        //    return false;
        //}

    }
}
