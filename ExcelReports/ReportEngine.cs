using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.Data;

namespace ExcelReports
{
    public class ReportEngine
    {
        public static void CreateExcelDocument(
        String TemplateFileName
        , List<ER_Common.Str_Parameter> Parameters
        , DataSet Ds_Source
        , DataSet Ds_Source_Pivot
        , DataSet Ds_Source_Pivot_Desc
        , DataSet Ds_Source_Pivot_Totals
        , String SaveFileName
        , Boolean IsProtected = false
        , ER_Common.eExcelFileFormat FileFormat = ER_Common.eExcelFileFormat.xlNormal)
        {
            ReportEngine_NativeExcel.CreateExcelDocument(
                TemplateFileName
                , Parameters
                , Ds_Source
                , Ds_Source_Pivot
                , Ds_Source_Pivot_Desc
                , Ds_Source_Pivot_Totals
                , SaveFileName
                , IsProtected
                , FileFormat);

            //ExcelReports_VB.ClsParameters P = new ExcelReports_VB.ClsParameters();
            //Parameters.ForEach(O => { P.Add(O.Name, O.Value); });

            //NativeExcel.XlFileFormat NxlFileFormat = ER_Common.ParseEnum<NativeExcel.XlFileFormat>(FileFormat.ToString());

            //ExcelReports_VB.Methods_NativeExcel.NativeExcel_CreateExcelDocument(
            //    TemplateFileName
            //    , P
            //    , Ds_Source
            //    , Ds_Source_Pivot
            //    , Ds_Source_Pivot_Desc
            //    , Ds_Source_Pivot_Totals
            //    , SaveFileName
            //    , IsProtected
            //    , NxlFileFormat);
        }

    }
}
