using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft;
using Microsoft.VisualBasic;
using System.Data;
using System.IO;

namespace ExcelReports
{
    public class ER_Common
    {
        #region _Variables

        public struct Str_Parameter
        {
            public String Name;
            public Object Value;
        }

        public struct Str_ParsedExcelRange
        {
            public Int32 X1;
            public Int32 Y1;
            public Int32 X2;
            public Int32 Y2;
        }

        public enum eExcelFileFormat
        {
            xlNormal = 0,
            xlExcel5 = 1,
            xlExcel97 = 2,
            xlOpenXMLWorkbook = 3,
            xlHtml = 4,
            xlCSV = 5,
            xlText = 6,
            xlUnicodeCSV = 7,
            xlUnicodeText = 8
        }

        #endregion

        #region _Methods

        internal static object IsNull(object Obj_Input, object Obj_NullOutput)
        {
            if (Obj_Input == null || Information.IsDBNull(Obj_Input))
            { return Obj_NullOutput; }
            else
            { return Obj_Input; }
        }

        internal static string GenerateChr(Int32 Input)
        {
            Int32 Ct = 0;
            Int32 TmpRes = 0;
            string OutputChr = "";

            //while ((26 ^ Ct) < Input)
            while (Math.Pow(26, Ct) < Input)
            { Ct++; }

            if (Ct > 0)
            { Ct--; }

            while (Input > 0)
            {
                TmpRes = Input / ER_Common.RaiseTruncate(26, Ct);
                if ((Input % 26) == 0 && Ct > 0)
                { TmpRes = TmpRes - Ct; }
                OutputChr = OutputChr + Strings.Chr(TmpRes + 64);
                Input = Input - (ER_Common.RaiseTruncate(26, Ct) * TmpRes);
                Ct--;
            }

            return OutputChr;
        }

        internal static Int32 RaiseTruncate(Double X, Double Y)
        { return Convert.ToInt32(Math.Truncate((Math.Pow(X, Y)))); }

        internal static object[,] ConvertDataTo2DimArray(DataTable Dt, String[] Fields, Int32 RowStart = -1, Int32 RowEnd = -1)
        {
            if (Dt.Rows.Count == 0)
            { return null; }

            Int32 RowCt = 0;
            Int32 RowCtEnd = 0;

            if (RowStart > -1)
            { RowCt = RowStart; }
            else
            { RowStart = 0; }

            if (RowEnd > -1)
            { RowCtEnd = RowEnd; }
            else
            { RowCtEnd = Dt.Rows.Count - 1; }

            if (RowCt >= Dt.Rows.Count)
            { RowCt = Dt.Rows.Count - 1; }

            if (RowCtEnd >= Dt.Rows.Count)
            { RowCtEnd = Dt.Rows.Count - 1; }

            Int32 RowLength = RowCtEnd - RowCt;
            Int32 ColumnLength = 0;
            if (Fields != null)
            { ColumnLength = Fields.Length; }
            else
            { ColumnLength = Dt.Columns.Count; }

            object[,] RV = new object[RowLength, ColumnLength];

            while (RowCt <= RowCtEnd)
            {
                Int32 RV_RowCt = RowCt - RowStart;

                if (Fields != null)
                {
                    for (Int32 Ct = 0; Ct < Fields.Length; Ct++)
                    {
                        Int32 RV_ColumnCt = Ct;
                        for (Int32 Ct2 = 0; Ct2 < Dt.Columns.Count; Ct2++)
                        {
                            if (Fields[Ct].Trim() == "")
                            {
                                RV[RV_RowCt, RV_ColumnCt] = "";
                                break;
                            }
                            if (Fields[Ct].ToUpper() == Dt.Columns[Ct2].ColumnName.ToUpper())
                            { ConvertDataTo2DimArray_Ex(Dt.Rows[RowCt], Ct2, ref RV[RV_RowCt, RV_ColumnCt]); }
                        }
                    }
                }
                else
                {
                    for (Int32 Ct = 0; Ct < Dt.Columns.Count; Ct++)
                    {
                        Int32 RV_ColumnCt = Ct;
                        ConvertDataTo2DimArray_Ex(Dt.Rows[RowCt], Ct, ref RV[RV_RowCt, RV_ColumnCt]);
                    }
                }
                RowCt++;
            }

            return RV;
        }

        internal static object[,] ConvertDataTo2DimArray(DataRow Dr, String[] Fields, Int32 RowStart = -1, Int32 RowEnd = -1)
        {
            Int32 RowLength = 1;
            Int32 ColumnLength = 0;

            if (Fields != null)
            { ColumnLength = Fields.Length; }
            else
            { ColumnLength = Dr.Table.Columns.Count; }

            Object[,] RV = new Object[RowLength, ColumnLength];

            Int32 RV_RowCt = 0;
            if (Fields != null)
            {
                for (Int32 Ct = 0; Ct < Fields.Length; Ct++)
                {
                    Int32 RV_ColumnCt = Ct;

                    for (Int32 Ct2 = 0; Ct2 < Dr.Table.Columns.Count; Ct2++)
                    {
                        if (Fields[Ct].Trim() == "")
                        {
                            RV[RV_RowCt, RV_ColumnCt] = "";
                            break;
                        }

                        if (Fields[Ct].ToUpper() == Dr.Table.Columns[Ct2].ColumnName.ToUpper())
                        { ConvertDataTo2DimArray_Ex(Dr, Ct2, ref RV[RV_RowCt, RV_ColumnCt]); }
                    }
                }
            }
            else
            {
                for (Int32 Ct = 0; Ct < Dr.Table.Columns.Count; Ct++)
                {
                    Int32 RV_ColumnCt = Ct;
                    ConvertDataTo2DimArray_Ex(Dr, Ct, ref RV[RV_RowCt, RV_ColumnCt]);
                }
            }

            return RV;
        }

        static void ConvertDataTo2DimArray_Ex(DataRow Dr_Source, Int32 RowFieldIndex, ref Object Value_Target)
        {
            if (Dr_Source[RowFieldIndex] == DBNull.Value)
            { Value_Target = ""; }
            else
            {
                if (Dr_Source.Table.Columns[RowFieldIndex].DataType == typeof(Guid))
                { Value_Target = ((Guid)Dr_Source[RowFieldIndex]).ToString(); }
                else
                { Value_Target = Dr_Source[RowFieldIndex]; }
            }
        }

        internal static Int32[] ParseExcelRange_Old(String Excel_Range)
        {
            Int32[] Return_Int = new int[3];
            for (Int32 Ct = 0; Ct < Return_Int.Length; Ct++)
            { Return_Int[Ct] = 0; }

            String Tmp_Excel_Range = Excel_Range;

            String St_Range1 = "";
            String St_Range2 = "";

            St_Range1 = Strings.Mid(Tmp_Excel_Range, 1, Strings.InStr(Tmp_Excel_Range, ":") - 1);
            St_Range2 = Strings.Mid(Tmp_Excel_Range, Strings.InStr(Tmp_Excel_Range, ":") + 1);

            ParseExcelRange_Ex(St_Range1, ref Return_Int[0], ref Return_Int[1]);
            ParseExcelRange_Ex(St_Range2, ref Return_Int[2], ref Return_Int[3]);

            return Return_Int;
        }

        internal static Str_ParsedExcelRange ParseExcelRange(String Excel_Range)
        {
            Str_ParsedExcelRange Return_Parsed = new Str_ParsedExcelRange();

            String Tmp_Excel_Range = Excel_Range;

            String St_Range1 = "";
            String St_Range2 = "";

            St_Range1 = Strings.Mid(Tmp_Excel_Range, 1, Strings.InStr(Tmp_Excel_Range, ":") - 1);
            St_Range2 = Strings.Mid(Tmp_Excel_Range, Strings.InStr(Tmp_Excel_Range, ":") + 1);

            ParseExcelRange_Ex(St_Range1, ref Return_Parsed.X1, ref Return_Parsed.Y1);
            ParseExcelRange_Ex(St_Range2, ref Return_Parsed.X2, ref Return_Parsed.Y2);

            return Return_Parsed;
        }

        internal static void ParseExcelRange_Ex(String St_Range, ref Int32 Value1, ref Int32 Value2)
        {
            String St_Parsed1 = "";
            String St_Parsed2 = "";

            foreach (Char C in St_Range)
            {
                if (!Information.IsNumeric(C))
                { St_Parsed1 = St_Parsed1 + C.ToString(); }
                else
                { St_Parsed2 = St_Parsed2 + C.ToString(); }
            }

            St_Parsed1 = St_Parsed1.ToUpper();
            St_Parsed2 = St_Parsed2.ToUpper();

            Int32 Digit = St_Parsed1.Length - 1;
            Int32 Result = 0;

            foreach (Char C in St_Parsed1)
            {
                //Result = Result + ((26 ^ Digit) * (Strings.Asc(C) - 64));
                Result = Result + ((26 ^ Digit) * (Convert.ToInt32(C) - 64));
                Digit--;
            }

            Value1 = Result;
            Value2 = Convert_Int32(St_Parsed2);
        }

        internal static Int32 ParseExcelRange_GetHeight(String Excel_Range)
        {
            Str_ParsedExcelRange Range = ParseExcelRange(Excel_Range);
            Int32 Length = Range.Y2 - Range.Y1;
            return Length;
        }

        internal static Int32 ParseExcelRange_GetWidth(String Excel_Range)
        {
            Str_ParsedExcelRange Range = ParseExcelRange(Excel_Range);
            Int32 Length = Range.X2 - Range.X1;
            return Length;
        }

        internal static Int32 Convert_Int32(Object Value)
        { return Convert_Int32(Value, 0); }

        internal static Int32 Convert_Int32(Object Value, Int32 DefaultValue)
        {
            string ValueString = string.Empty;
            if (Value != null)
            {
                try { ValueString = Value.ToString(); }
                catch { }
            }

            Int32 ReturnValue;
            if (!Int32.TryParse(ValueString, out ReturnValue))
            { ReturnValue = DefaultValue; }
            return ReturnValue;
        }

        internal static String Convert_String(Object Value)
        { return Convert_String(Value, ""); }

        internal static String Convert_String(Object Value, String DefaultValue)
        { return Convert.ToString(IsNull(Value, DefaultValue)); }

        internal static Boolean Convert_Boolean(Object Value)
        { return Convert_Boolean(Value, false); }

        internal static bool Convert_Boolean(Object Value, bool DefaultValue)
        {
            string ValueString = string.Empty;
            try
            { ValueString = Value.ToString(); }
            catch { }

            bool ReturnValue;
            if (!bool.TryParse(ValueString, out ReturnValue))
            { ReturnValue = DefaultValue; }
            return ReturnValue;
        }

        internal static T ParseEnum<T>(String Value)
            where T : struct, IComparable, IFormattable, IConvertible
        { return ParseEnum<T>(Value, default(T)); }

        internal static T ParseEnum<T>(string Value, T DefaultValue)
            where T : struct, IComparable, IFormattable, IConvertible
        {
            if (Enum.IsDefined(typeof(T), Value))
            { return (T)Enum.Parse(typeof(T), Value, true); }
            return DefaultValue;
        }

        internal static Boolean WriteFile(Stream S, String FilePath)
        {
            try
            {
                FileInfo Fi = new FileInfo(FilePath);
                if (!Fi.Directory.Exists)
                { Fi.Directory.Create(); }

                if (Fi.Exists)
                { Fi.Delete(); }

                const Int32 BufferLength = 2097152;

                //Fi.Create();
                //Fi.Attributes = FileAttributes.Hidden & FileAttributes.Temporary & FileAttributes.System;
                FileStream Fs = Fi.OpenWrite();

                Byte[] Buffer;
                Int32 Offset = 0;
                while (true)
                {
                    Buffer = new Byte[BufferLength - 1];
                    Int32 BytesRead = S.Read(Buffer, Offset, BufferLength);

                    if (BytesRead == 0)
                    { break; }

                    Fs.Write(Buffer, Offset, BufferLength);
                    Offset += BytesRead;
                }

                Fs.Close();

                return true;
            }
            catch
            { return false; }
        }

        #endregion
    }
}
