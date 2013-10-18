using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using NativeExcel;
using Microsoft.VisualBasic;
using System.IO;

namespace ExcelReports
{
    internal class ReportEngine_NativeExcel
    {
        #region _Variables

        const Int32 CnsExcelMaxWidth = 256;
        const Int32 CnsExcelMaxHeight = 65536;
        const String CnsExcelKeyword_TemplateVersion = "@TemplateVersion";
        const String CnsExcelKeyword_Settings = "[#]Settings";
        const String CnsExcelKeyword_Settings_End = "[#]End_Settings";
        const String CnsExcelKeyword_Parameters = "[#]Parameters";
        const String CnsExcelKeyword_Parameters_End = "[#]End_Parameters";
        const String CnsExcelKeyword_Sections = "[#]Sections";
        const String CnsExcelKeyword_Sections_End = "[#]End_Sections";
        const String CnsExcelKeyword_DataTable = "[#]DataTable";
        const String CnsExcelKeyword_DataTable_End = "[#]End_DataTable";
        const String CnsExcelKeyword_DataTable_Header = "[#]DataTable_Header";
        const String CnsExcelKeyword_DataTable_Header_End = "[#]End_DataTable_Header";
        const String CnsExcelKeyword_DataTable_Footer = "[#]DataTable_Footer";
        const String CnsExcelKeyword_DataTable_Footer_End = "[#]End_DataTable_Footer";
        const String CnsExcelKeyword_DataTable_Pivot = "[#]DataTable_Pivot";
        const String CnsExcelKeyword_DataTable_Pivot_End = "[#]End_DataTable_Pivot";
        const String CnsExcelKeyword_DataTable_Pivot_Header = "[#]DataTable_Pivot_Header";
        const String CnsExcelKeyword_DataTable_Pivot_Header_End = "[#]End_DataTable_Pivot_Header";
        const String CnsExcelKeyword_DataTable_Pivot_Totals = "[#]DataTable_Pivot_Totals";
        const String CnsExcelKeyword_DataTable_Pivot_Totals_End = "[#]End_DataTable_Pivot_Totals";

        struct Str_Settings
        {
            public Int32 DocumentLimit;
            public Int32 DocumentWidth;
            public Boolean IsRepeatHeader;
        }

        struct Str_Sections
        {
            public Int32 Ct;
            public String Type;
            public String Location;
        }

        struct Str_Parameters
        {
            public String Name;
            public String Type;
            public String Value;
        }

        struct Str_DataTable
        {
            public Int32 Ct;
            public String Name;
            public String GroupName;
            public String SourceKey;
            public String TargetKey;
            public Boolean IsSubTable;
            public String Location;
            public Int32 Items;
            public String ParentName;
        }

        struct Str_DataTable_Section
        {
            public Int32 Ct;
            public String Name;
            public String Location;
        }

        struct Str_DataTable_Field
        {
            public Int32 Ct;
            public Int32 DataTable_Ct;
            public String Name;
            public Int32 Position;
        }

        #endregion

        #region _Methods

        public static Boolean CreateExcelDocument(
            String TemplateFileName
            , List<ER_Common.Str_Parameter> Parameters
            , DataSet Ds_Source
            , DataSet Ds_Source_Pivot
            , DataSet Ds_Source_Pivot_Desc
            , DataSet Ds_Source_Pivot_Totals
            , String SaveFileName
            , Boolean IsProtected
            , ER_Common.eExcelFileFormat FileFormat)
        {
            Int32 TemplateVersion = 0;

            IWorkbook Wb_Template = Factory.OpenWorkbook(TemplateFileName);
            IWorksheet Ws_Parameters = Wb_Template.Worksheets["Parameters"];

            Int32[] LineInfo = CreateExcelDocument_ReadLineInfo(Ws_Parameters, CnsExcelKeyword_Settings, CnsExcelKeyword_Settings_End);

            Int32 Ct_Start = LineInfo[0];
            Int32 Ct_End = LineInfo[1];
            Int32 Ct = 0;

            for (Ct = Ct_Start; Ct <= Ct_End; Ct++)
            {
                String ExcelText = Ws_Parameters.Range["A" + Ct.ToString()].Characters.Text;

                if (Strings.InStr(ExcelText, CnsExcelKeyword_TemplateVersion) > 0)
                {
                    TemplateVersion = ER_Common.Convert_Int32(Strings.Mid(ExcelText, Strings.Len(CnsExcelKeyword_TemplateVersion) + 1).Trim());
                    break;
                }
            }

            Boolean Result = false;

            switch (TemplateVersion)
            {
                case 1:
                    throw new NotImplementedException("Version 1 Template is not implemented.");
                case 2:
                    Result = ReportEngine_NativeExcel.CreateExcelDocument_V2(
                         TemplateFileName
                         , Parameters
                         , Ds_Source
                         , SaveFileName
                         , IsProtected
                         , FileFormat);
                    break;
                case 3:
                    Result = ReportEngine_NativeExcel.CreateExcelDocument_V3(
                        TemplateFileName
                        , Parameters
                        , Ds_Source
                        , Ds_Source_Pivot
                        , Ds_Source_Pivot_Desc
                        , Ds_Source_Pivot_Totals
                        , SaveFileName
                        , IsProtected
                        , FileFormat);
                    break;
            }

            return Result;
        }

        public static Boolean CreateExcelDocument_V2(
            String TemplateFileName
            , List<ER_Common.Str_Parameter> Parameters
            , DataSet Ds_Source
            , String SaveFileName
            , Boolean IsProtected
            , ER_Common.eExcelFileFormat FileFormat)
        {
            IWorkbook Wb_Document;
            IWorkbook Wb_Template;

            IWorksheet Ws_Document;
            IWorksheet Ws_Parameters;
            IWorksheet Ws_Template;

            //[-]

            CreateExcelDocument_CheckStringFields(Ds_Source);

            //[-]

            Wb_Template = Factory.OpenWorkbook(TemplateFileName);
            Ws_Parameters = Wb_Template.Worksheets["Parameters"];
            Ws_Template = Wb_Template.Worksheets["Template"];

            Wb_Document = Factory.OpenWorkbook(TemplateFileName);
            foreach (IWorksheet Ws in Wb_Document.Worksheets)
            { Ws.Delete(); }

            Ws_Document = Wb_Document.Worksheets["Template"];
            Ws_Document.Name = "Document";

            //[-]

            //Get Settings
            Str_Settings Document_Settings = CreateExcelDocument_GetSettings(Ws_Parameters);

            //Get Parameters
            List<Str_Parameters?> Document_Parameters = CreateExcelDocument_GetParameters(Ws_Parameters, Parameters);

            //Get Sections
            List<Str_Sections?> Document_Sections = CreateExcelDocument_GetSections(Ws_Parameters);

            //Get DataTables
            List<Str_DataTable?> Document_Tables = CreateExcelDocument_GetDataTables(Ws_Parameters);

            //Get DataTable Fields
            List<Str_DataTable_Field?> Document_Tables_Fields = CreateExcelDocument_GetFields(Document_Tables, Ws_Template);

            //Populate Parameter Values
            for (Int32 Inner_Ct1 = 0; Inner_Ct1 < Document_Settings.DocumentLimit; Inner_Ct1++)
            {
                for (Int32 Inner_Ct2 = 0; Inner_Ct2 < Document_Settings.DocumentWidth; Inner_Ct2++)
                {
                    String Excel_Text = Ws_Template.Range[ER_Common.GenerateChr(Inner_Ct2 + 1) + (Inner_Ct1 + 1)].Characters.Text;
                    String Parameter_Name = "";

                    if (Strings.InStr(Excel_Text, "[@") > 0)
                    {
                        Parameter_Name =
                            Strings.Mid(
                            Excel_Text
                            , Strings.InStr(Excel_Text, "[@") + Strings.Len("[@")
                            , (Strings.InStrRev(Excel_Text, "]") - Strings.Len("]")) - Strings.Len("[@"));
                    }

                    Str_Parameters? Sp = Document_Parameters.FirstOrDefault(O => O.Value.Name == Parameter_Name);
                    if (Sp != null)
                    {
                        Str_Parameters Sp_Ex = Sp.Value;

                        TypeCode Parameter_Type = ER_Common.ParseEnum<TypeCode>(Sp.Value.Type, TypeCode.String);
                        switch (Parameter_Type)
                        {
                            case TypeCode.String:
                                Sp_Ex.Value = Sp_Ex.Value.Replace(Constants.vbTab, " ");
                                Sp_Ex.Value = Sp_Ex.Value.Replace(Constants.vbCrLf, " ");
                                Sp_Ex.Value = Sp_Ex.Value.Replace(Constants.vbCr, " ");
                                Sp_Ex.Value = Sp_Ex.Value.Replace(Constants.vbLf, " ");
                                break;
                        }

                        Ws_Template.Cells[Inner_Ct1 + 1, Inner_Ct2 + 1].Value = Sp_Ex.Value;
                    }
                }
            }

            //Populate Ws_Document

            DataTable Dt_DocumentParameters = new DataTable();
            Dt_DocumentParameters.Columns.Add("Sheet", typeof(Int32));
            Dt_DocumentParameters.Columns.Add("Page", typeof(Int32));

            for (Int32 Inner_Ct = 0; Inner_Ct < Document_Tables.Count(); Inner_Ct++)
            {
                Str_DataTable Item_Table = Document_Tables[Inner_Ct].Value;
                Dt_DocumentParameters.Columns.Add(@"Dt_" + Item_Table.Ct, typeof(Int32));
                Item_Table.Items = Ds_Source.Tables[Item_Table.Ct - 1].Rows.Count;
            }

            Int32 Ct_Page = 0;
            Int32 Ct_Sheet = 0;

            while (true)
            {
                Boolean IsItem = Document_Tables.Any(O => O.HasValue && O.Value.Items > 0);
                if (IsItem)
                {
                    DataRow Inner_Dr_New = Dt_DocumentParameters.NewRow();
                    Dt_DocumentParameters.Rows.Add(Inner_Dr_New);

                    for (Int32 Inner_Ct = 0; Inner_Ct < Document_Tables.Count(); Inner_Ct++)
                    {
                        Str_DataTable Item_Table = Document_Tables[Inner_Ct].Value;
                        ER_Common.Str_ParsedExcelRange Inner_PR = ER_Common.ParseExcelRange(Item_Table.Location);

                        Int32 Inner_Limit = Math.Abs(Inner_PR.Y2 - Inner_PR.Y1) + 1;
                        Int32 Inner_Ct_Tmp = 0;

                        if (Item_Table.Items > Inner_Limit)
                        { Inner_Ct_Tmp = Inner_Limit; }
                        else
                        { Inner_Ct_Tmp = Item_Table.Items; }

                        Item_Table.Items = Item_Table.Items - Inner_Ct_Tmp;
                        Inner_Dr_New["Dt_" + Item_Table.Ct] = Inner_Ct_Tmp;
                    }

                    Inner_Dr_New["Page"] = Ct_Page;
                    Inner_Dr_New["Sheet"] = Ct_Sheet;
                }
                else
                { break; }

                Ct_Page++;
            }

            //[-]

            foreach (DataRow Dr_Page in Dt_DocumentParameters.Select("", "Page"))
            {
                Int32 Page_TopLimit = ER_Common.Convert_Int32(Dr_Page["Page"]) * Document_Settings.DocumentLimit;
                Int32 Page_BottomLimit = (ER_Common.Convert_Int32(Dr_Page["Page"]) + 1) * Document_Settings.DocumentLimit; ;

                String Location_Template =
                    @"A1:"
                    + ER_Common.GenerateChr(Document_Settings.DocumentWidth)
                    + Document_Settings.DocumentLimit.ToString();

                String Location_Document =
                    @"A"
                    + (Page_TopLimit + 1).ToString()
                    + @":"
                    + ER_Common.GenerateChr(Document_Settings.DocumentWidth)
                    + Page_BottomLimit.ToString();

                Ws_Template.Range[Location_Template].Copy(Ws_Document.Range[Location_Document], XlPasteType.xlPasteAll);

                var List_Table = (from O in Document_Tables where O.HasValue orderby O.Value.Ct select O.Value).ToList();
                foreach (var Item_Table in List_Table)
                {
                    foreach (DataColumn Dc in Dt_DocumentParameters.Columns)
                    {
                        if (Dc.ColumnName == @"Dt_" + Item_Table.Ct.ToString())
                        {
                            if (ER_Common.Convert_Int32(Dr_Page[Dc.ColumnName]) > 0)
                            {
                                if (Ds_Source.Tables[Item_Table.Ct - 1].Rows.Count > 0)
                                {
                                    ER_Common.Str_ParsedExcelRange Inner_PR = ER_Common.ParseExcelRange(Item_Table.Location);
                                    Int32 Inner_ItemCount = 0;

                                    foreach (DataRow Inner_Dr in Dt_DocumentParameters.Select("Page < " + ER_Common.Convert_Int32(Dr_Page["Page"]).ToString()))
                                    { Inner_ItemCount = Inner_ItemCount + ER_Common.Convert_Int32(Inner_Dr[Dc.ColumnName]); }

                                    Int32 Inner_RowStart = Inner_ItemCount;
                                    Int32 Inner_RowEnd = (Inner_ItemCount + ER_Common.Convert_Int32(Dr_Page[Dc.ColumnName])) - 1;

                                    for (Int32 Inner_Ct_Row = 0; Inner_Ct_Row < (Inner_RowEnd - Inner_RowStart); Inner_Ct_Row++)
                                    {
                                        var Inner_List_Dtf =
                                            (
                                            from O in Document_Tables_Fields
                                            orderby O.Value.Ct
                                            where O.HasValue && O.Value.DataTable_Ct == Item_Table.Ct
                                            select O.Value).ToList();
                                        foreach (Str_DataTable_Field Inner_Dtf in Inner_List_Dtf)
                                        {
                                            Ws_Document.Cells[
                                                (Inner_PR.Y1 + Page_TopLimit) + Inner_Ct_Row
                                                , Inner_PR.X1 + Inner_Dtf.Position].Value =
                                                    ER_Common.Convert_String(Ds_Source.Tables[Item_Table.Ct - 1].Rows[Inner_Ct_Row + Inner_RowStart][Inner_Dtf.Name]);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                ER_Common.Str_ParsedExcelRange Inner_PR = ER_Common.ParseExcelRange(Item_Table.Location);
                                String Inner_Location =
                                        ER_Common.GenerateChr(Inner_PR.X1)
                                        + (Inner_PR.Y1 + Page_TopLimit).ToString()
                                        + @":"
                                        + ER_Common.GenerateChr(Inner_PR.X2)
                                        + (Inner_PR.Y1 + Page_TopLimit).ToString();

                                Ws_Document.Range[Inner_Location].ClearContents();
                            }
                        }
                    }
                }

                Ws_Document.HPageBreaks.Add(Ws_Document.Range[@"A" + Page_BottomLimit + 1]);
            }

            Ws_Document.Activate();
            Ws_Document.Range["A2"].Select();

            //Save the Document
            if (IsProtected)
            {
                String RandomPassword = Guid.NewGuid().ToString();
                Ws_Document.EnableSelection = XlEnableSelection.xlNoSelection;
                Ws_Document.Protect(RandomPassword);
                Wb_Document.Protect(RandomPassword);
            }

            if (SaveFileName == "")
            { SaveFileName = "Excel_File"; }

            XlFileFormat NxlFileFormat = ER_Common.ParseEnum<XlFileFormat>(FileFormat.ToString());

            return Ws_Document.SaveAs(SaveFileName, NxlFileFormat);
        }

        public static Boolean CreateExcelDocument_V3(
            String TemplateFileName
            , List<ER_Common.Str_Parameter> Parameters
            , DataSet Ds_Source
            , DataSet Ds_Source_Pivot
            , DataSet Ds_Source_Pivot_Desc
            , DataSet Ds_Source_Pivot_Totals
            , String SaveFileName
            , Boolean IsProtected
            , ER_Common.eExcelFileFormat FileFormat)
        {
            IWorkbook Wb_Document;
            IWorkbook Wb_Template;

            IWorksheet Ws_Document;
            IWorksheet Ws_Parameters;
            IWorksheet Ws_Template;

            //[-]

            //Clear all String Fields of all TAB and RETURN chars
            CreateExcelDocument_CheckStringFields(Ds_Source);
            CreateExcelDocument_CheckStringFields(Ds_Source_Pivot);
            CreateExcelDocument_CheckStringFields(Ds_Source_Pivot_Desc);
            CreateExcelDocument_CheckStringFields(Ds_Source_Pivot_Totals);

            //[-]

            Wb_Template = Factory.OpenWorkbook(TemplateFileName);
            Ws_Parameters = Wb_Template.Worksheets["Parameters"];
            Ws_Template = Wb_Template.Worksheets["Template"];

            Wb_Document = Factory.OpenWorkbook(TemplateFileName);
            foreach (IWorksheet Ws in Wb_Document.Worksheets)
            { Ws.Delete(); }

            Ws_Document = Wb_Document.Worksheets["Template"];
            Ws_Document.Name = "Document";

            //[-]

            //Clear the contents of Ws_Document
            Ws_Document.Range[@"A1:" + ER_Common.GenerateChr(CnsExcelMaxWidth) + CnsExcelMaxHeight.ToString()].Clear();

            //[-]

            //Get Settings
            Str_Settings Document_Settings = CreateExcelDocument_GetSettings(Ws_Parameters);

            //Get Parameters
            List<Str_Parameters?> Document_Parameters = CreateExcelDocument_GetParameters(Ws_Parameters, Parameters);

            //Get Sections
            List<Str_Sections?> Document_Sections = CreateExcelDocument_GetSections(Ws_Parameters);

            //Get DataTables
            List<Str_DataTable?> Document_Tables = CreateExcelDocument_GetDataTables(Ws_Parameters);

            //Get DataTable Headers
            List<Str_DataTable_Section?> Document_Tables_Headers = CreateExcelDocument_GetDataTables_Headers(Ws_Parameters);

            //Get DataTable Footers
            List<Str_DataTable_Section?> Document_Tables_Footers = CreateExcelDocument_GetDataTables_Footers(Ws_Parameters);

            //Get DataTable Fields
            List<Str_DataTable_Field?> Document_Tables_Fields = CreateExcelDocument_GetFields(Document_Tables, Ws_Template);

            //Get Pivot DataTable
            List<Str_DataTable?> Document_PivotTables = CreateExcelDocument_GetDataTables_PivotTables(Ws_Parameters);

            //Get Pivot DataTable Fields
            List<Str_DataTable_Field?> Document_PivotTables_Fields = CreateExcelDocument_GetFields(Document_PivotTables, Ws_Template);

            //Get Pivot DataTable Headers
            List<Str_DataTable_Section?> Document_PivotHeaders = CreateExcelDocument_GetDataTables_PivotTables_Headers(Ws_Parameters);

            //Get Pivot DataTable Header Fields
            List<Str_DataTable_Field?> Document_PivotHeaders_Fields = CreateExcelDocument_GetFields(Document_PivotHeaders, Ws_Template);

            //Get Pivot Totals
            List<Str_DataTable_Section?> Document_PivotTotals = CreateExcelDocument_GetDataTables_PivotTables_Totals(Ws_Parameters);

            //Get Pivot Totals Fields
            List<Str_DataTable_Field?> Document_PivotTotals_Fields = CreateExcelDocument_GetFields(Document_PivotTotals, Ws_Template);

            //Populate Parameter Values
            for (Int32 Ct = 0; Ct < Document_Settings.DocumentLimit; Ct++)
            {
                for (Int32 Ct2 = 0; Ct2 < Document_Settings.DocumentWidth; Ct2++)
                {
                    String Excel_Text = Ws_Template.Range[ER_Common.GenerateChr(Ct2 + 1) + (Ct + 1)].Characters.Text;
                    String Parameter_Name = "";

                    if (Strings.InStr(Excel_Text, "[@") > 0)
                    {
                        Parameter_Name =
                            Strings.Mid(
                            Excel_Text
                            , Strings.InStr(Excel_Text, "[@") + Strings.Len("[@")
                            , (Strings.InStrRev(Excel_Text, "]") - Strings.Len("]")) - Strings.Len("[@"));
                    }

                    Str_Parameters? Sp = Document_Parameters.FirstOrDefault(O => O.Value.Name == Parameter_Name);
                    if (Sp != null)
                    {
                        Str_Parameters Sp_Ex = Sp.Value;

                        TypeCode Parameter_Type = ER_Common.ParseEnum<TypeCode>(Sp.Value.Type, TypeCode.String);
                        switch (Parameter_Type)
                        {
                            case TypeCode.String:
                                Sp_Ex.Value = Sp_Ex.Value.Replace(Constants.vbTab, " ");
                                Sp_Ex.Value = Sp_Ex.Value.Replace(Constants.vbCrLf, " ");
                                Sp_Ex.Value = Sp_Ex.Value.Replace(Constants.vbCr, " ");
                                Sp_Ex.Value = Sp_Ex.Value.Replace(Constants.vbLf, " ");
                                break;
                        }

                        Ws_Template.Cells[Ct + 1, Ct2 + 1].Value = Sp_Ex.Value;
                    }
                }
            }

            //Populate Ws_Document

            Int32 Ct_CurrentRow = 1;

            //Set Header
            Str_Sections? Section_Header = Document_Sections.FirstOrDefault(O => O.Value.Type.ToUpper() == "HEADER");
            if (Section_Header != null)
            {
                Str_Sections Inner_Section_Header = Section_Header.Value;
                ER_Common.Str_ParsedExcelRange Inner_PR;
                String Location_Template;
                String Location_Document;
                Int32 Length;

                Inner_PR = ER_Common.ParseExcelRange(Inner_Section_Header.Location);
                Length = (Inner_PR.Y2 - Inner_PR.Y1) + 1;
                Location_Template = @"A" + Inner_PR.Y1.ToString() + @":" + ER_Common.GenerateChr(Document_Settings.DocumentWidth) + Inner_PR.Y2.ToString();
                Location_Document = @"A" + Ct_CurrentRow.ToString() + @":" + ER_Common.GenerateChr(Document_Settings.DocumentWidth) + (Ct_CurrentRow + Length).ToString();

                Ws_Template.Range[Location_Template].Copy(Ws_Document.Range[Location_Document]);

                if (Document_Settings.IsRepeatHeader)
                {
                    ER_Common.Str_ParsedExcelRange Inner2_PR;
                    Str_Sections? Section_Repeat = Document_Sections.FirstOrDefault(O => O.Value.Type.ToUpper() == "REPEAT");
                    if (Section_Repeat != null)
                    { Inner2_PR = ER_Common.ParseExcelRange(Section_Repeat.Value.Location); }
                    else
                    { Inner2_PR = Inner_PR; }

                    Ws_Document.PageSetup.PrintTitleRows = @"$" + Inner2_PR.Y1.ToString() + @":" + "$" + Inner2_PR.Y2;
                }

                Ct_CurrentRow = Ct_CurrentRow + Length;
            }

            //Set Tables
            var List_Tables =
                (from O in Document_Tables
                 where O.Value.IsSubTable == false
                 select O.Value).ToList();

            //foreach (Str_DataTable Item_Table in List_Tables)
            for (Int32 Ct = 0; Ct < List_Tables.Count(); Ct++)
            {
                Str_DataTable Item_Table = List_Tables[Ct];

                String Location = Item_Table.Location;
                Int32 Ct_Table = Item_Table.Ct - 1;
                Int32 ItemCount = CreateExcelDocument_V3_CountItem(Item_Table, List_Tables, Ds_Source, null);
                ER_Common.Str_ParsedExcelRange Inner_PR = ER_Common.ParseExcelRange(Location);

                Int32 Inner_ItemCount = 0;
                var Inner_List_Tables =
                    from O in Document_Tables
                    where
                        O.Value.IsSubTable == false
                        && O.Value.Ct != (Ct_Table + 1)
                    select O;

                foreach (Str_DataTable Inner2_Table in Inner_List_Tables)
                {
                    ER_Common.Str_ParsedExcelRange Inner2_PR = ER_Common.ParseExcelRange(Inner2_Table.Location);
                    if (
                            (
                                ((Inner2_PR.X1 <= Inner_PR.X1) || (Inner2_PR.X1 >= Inner_PR.X2))
                                ||
                                ((Inner2_PR.X2 <= Inner_PR.X1) || (Inner2_PR.X2 >= Inner_PR.X2))
                            )
                            && (Inner_PR.Y1 >= Inner2_PR.Y2)
                        )
                    {
                        Inner_ItemCount = Inner_ItemCount + CreateExcelDocument_V3_CountItem(Inner2_Table, List_Tables, Ds_Source, null);

                        Str_DataTable_Section? Inner2_Dth = Document_Tables_Headers.FirstOrDefault(O => O.Value.Name == Inner2_Table.Name);
                        if (Inner2_Dth != null)
                        { Inner_ItemCount = Inner_ItemCount + (ER_Common.ParseExcelRange_GetHeight(Inner2_Dth.Value.Location) + 1); }

                        Str_DataTable_Section? Inner2_Dtf = Document_Tables_Footers.FirstOrDefault(O => O.Value.Name == Inner2_Table.Name);
                        if (Inner2_Dtf != null)
                        { Inner_ItemCount = Inner_ItemCount + (ER_Common.ParseExcelRange_GetHeight(Inner2_Dtf.Value.Location) + 1); }
                    }
                }

                Int32 Inner_Ct_CurrentRow = Ct_CurrentRow + Inner_ItemCount;

                //Table Header
                Str_DataTable_Section? Inner_Dth = Document_Tables_Headers.FirstOrDefault(O => O.Value.Name == Item_Table.Name);
                if (Inner_Dth != null)
                {
                    Inner_PR = ER_Common.ParseExcelRange(Inner_Dth.Value.Location);
                    Int32 Inner_Length = ER_Common.ParseExcelRange_GetHeight(Inner_Dth.Value.Location);
                    String Inner_Source_Location = Inner_Dth.Value.Location;
                    String Inner_Target_Location =
                        ER_Common.GenerateChr(Inner_PR.X1)
                        + Inner_Ct_CurrentRow.ToString()
                        + @":"
                        + ER_Common.GenerateChr(Inner_PR.X2)
                        + (Inner_Ct_CurrentRow + Inner_Length).ToString();

                    Inner_Ct_CurrentRow = Inner_Ct_CurrentRow + Inner_Length + 1;
                    Inner_ItemCount = Inner_ItemCount + Inner_Length + 1;
                }

                //Table Items
                if (ItemCount > 0)
                {
                    CreateExcelDocument_V3_SubTable(
                     Ds_Source
                     , (from O in List_Tables select new Str_DataTable?(O)).ToList()
                     , Document_Tables_Fields
                     , Ws_Template
                     , Ws_Document
                     , ref Inner_Ct_CurrentRow
                     , new List<Str_DataTable?>() { Item_Table }
                     , null);
                }

                //Table Footer
                Int32 Inner_Tables_Footer_Length = 0;
                Str_DataTable_Section? Inner_Dtf = Document_Tables_Footers.FirstOrDefault(O => O.Value.Name == Item_Table.Name);
                if (Inner_Dtf != null)
                {
                    Inner_PR = ER_Common.ParseExcelRange(Inner_Dth.Value.Location);
                    Int32 Inner_Length = ER_Common.ParseExcelRange_GetHeight(Inner_Dth.Value.Location);
                    String Inner_Source_Location = Inner_Dth.Value.Location;
                    String Inner_Target_Location =
                        ER_Common.GenerateChr(Inner_PR.X1)
                        + Inner_Ct_CurrentRow.ToString()
                        + @":"
                        + ER_Common.GenerateChr(Inner_PR.X2)
                        + (Inner_Ct_CurrentRow + Inner_Length).ToString();

                    Inner_ItemCount = Inner_ItemCount + Inner_Length + 1;
                    Inner_Tables_Footer_Length = Inner_Tables_Footer_Length + Inner_Length + 1;
                }

                //Pivot Tables

                Int32 Inner_OffsetColumn = 0;
                Int32 Inner_OffsetRow = ItemCount + Inner_Tables_Footer_Length;

                //Reuse Inner_CtCurrentRow, reassign with CtCurrentRow for the current iteration
                Inner_Ct_CurrentRow = Ct_CurrentRow;

                var List_PivotTables =
                    from O in Document_PivotTables
                    where
                        O.HasValue
                        && O.Value.ParentName == Item_Table.Name
                    select O.Value;
                foreach (Str_DataTable Item_PivotTable in List_PivotTables)
                {
                    Int32 Ct_PivotTable = Item_PivotTable.Ct - 1;

                    //Pivot Header

                    Str_DataTable_Section? Dts_Pivot_Header = Document_PivotHeaders.FirstOrDefault(O => O.Value.Name == Item_PivotTable.Name);
                    if (Dts_Pivot_Header == null)
                    { throw new Exception("Pivot Table must have a Header."); }

                    String Inner_PivotHeader_Location = Dts_Pivot_Header.Value.Location;
                    ER_Common.Str_ParsedExcelRange Inner_PivotHeader_Range = ER_Common.ParseExcelRange(Inner_PivotHeader_Location);
                    Int32 Inner_PivotHeader_Count = Ds_Source_Pivot.Tables[Ct_PivotTable].Rows.Count;
                    Int32 Inner_PivotHeader_Field_Count = (from O in Document_PivotHeaders_Fields where O.HasValue && O.Value.DataTable_Ct == Item_PivotTable.Ct select O).Count();
                    String Location_Target =
                        ER_Common.GenerateChr(Inner_PivotHeader_Range.X1 + Inner_OffsetColumn)
                            + Ct_CurrentRow.ToString()
                            + @":"
                            + ER_Common.GenerateChr(Inner_PivotHeader_Range.X1 + (((((Inner_PivotHeader_Range.X2 - Inner_PivotHeader_Range.X1) + 1) * (Inner_PivotHeader_Count - 1)) - 1) + Inner_OffsetColumn))
                            + (Ct_CurrentRow + Inner_OffsetRow).ToString();

                    Ws_Document.Range[Location_Target].Insert(XlInsertShiftDirection.xlShiftToRight);

                    ER_Common.Str_ParsedExcelRange Location_Target_Range = ER_Common.ParseExcelRange(Location_Target);
                    Location_Target =
                        ER_Common.GenerateChr(Location_Target_Range.X1)
                        + Location_Target_Range.Y1.ToString()
                        + @":"
                        + ER_Common.GenerateChr(Location_Target_Range.X2)
                        + Location_Target_Range.Y1.ToString();

                    Ws_Template.Range[Inner_PivotHeader_Location].Copy(Ws_Document.Range[Location_Target], XlPasteType.xlPasteFormats);

                    Int32 Outer_Ct = 0;
                    foreach (DataRow Inner_Dr in Ds_Source_Pivot.Tables[Ct_PivotTable].Rows)
                    {
                        Int32 Inner2_OffsetColumn = Outer_Ct * ((Inner_PivotHeader_Range.X2 - Inner_PivotHeader_Range.X1) + 1);

                        var Inner_List =
                            (
                            from O in Document_PivotHeaders_Fields
                            where
                                O.HasValue
                                && O.Value.DataTable_Ct == Item_PivotTable.Ct
                            select O.Value);
                        foreach (Str_DataTable_Field Inner_Item in Inner_List)
                        {
                            String Inner2_Location = ER_Common.GenerateChr((Inner_PivotHeader_Range.X1 + Inner_OffsetColumn) + Inner_Item.Position + Inner2_OffsetColumn) + Ct_CurrentRow;
                            Ws_Document.Range[Inner2_Location].Characters.Text = ER_Common.Convert_String(Inner_Dr[Inner_Item.Name]);
                        }
                        Outer_Ct++;
                    }

                    Inner_Ct_CurrentRow = Inner_Ct_CurrentRow + ER_Common.ParseExcelRange_GetHeight(Item_Table.Location);

                    //Pivot Items

                    String Inner_Pivot_Location = Item_PivotTable.Location;
                    ER_Common.Str_ParsedExcelRange Inner_Pivot_Range = ER_Common.ParseExcelRange(Inner_Pivot_Location);

                    Int32 Inner_Row_Start = Inner_Ct_CurrentRow;
                    Int32 Inner_Row_End = Inner_Row_Start + (ItemCount - 1);

                    String Inner_Location_Source = "";
                    String Inner_Location_Target = "";

                    //Table Formats and Borders
                    if (ItemCount == 1)
                    {
                        Inner_Location_Source =
                       ER_Common.GenerateChr(Inner_Pivot_Range.X1)
                       + (Inner_Pivot_Range.Y1 + 4).ToString()
                       + @":"
                       + ER_Common.GenerateChr(Inner_Pivot_Range.X2)
                       + (Inner_Pivot_Range.Y1 + 4).ToString();

                        Inner_Location_Target =
                            ER_Common.GenerateChr(Inner_Pivot_Range.X1)
                            + (Inner_Row_End).ToString()
                            + @":"
                            + ER_Common.GenerateChr(Inner_Pivot_Range.X1 + ((((Inner_Pivot_Range.X2 - Inner_Pivot_Range.X1) + 1) * Inner_PivotHeader_Count) - 1))
                            + (Inner_Row_End).ToString();

                        Ws_Template.Range[Inner_Location_Source].Copy(Ws_Document.Range[Inner_Location_Target], XlPasteType.xlPasteFormats);
                    }
                    else
                    {
                        // - Top
                        Inner_Location_Source =
                            ER_Common.GenerateChr(Inner_Pivot_Range.X1)
                            + Inner_Pivot_Range.Y1.ToString()
                            + @":"
                            + ER_Common.GenerateChr(Inner_Pivot_Range.X2)
                            + Inner_Pivot_Range.Y1.ToString();

                        Inner_Location_Target =
                            ER_Common.GenerateChr(Inner_Pivot_Range.X1)
                            + Inner_Row_Start.ToString()
                            + @":"
                            + ER_Common.GenerateChr(Inner_Pivot_Range.X1 + ((((Inner_Pivot_Range.X2 - Inner_Pivot_Range.X1) + 1) * Inner_PivotHeader_Count) - 1))
                            + (Inner_Row_Start).ToString();

                        Ws_Template.Range[Inner_Location_Source].Copy(Ws_Document.Range[Inner_Location_Target], XlPasteType.xlPasteFormats);

                        // - Middle
                        if (ItemCount > 2)
                        {
                            Inner_Location_Source =
                            ER_Common.GenerateChr(Inner_Pivot_Range.X1)
                            + (Inner_Pivot_Range.Y1 + 1).ToString()
                            + @":"
                            + ER_Common.GenerateChr(Inner_Pivot_Range.X2)
                            + (Inner_Pivot_Range.Y1 + 1).ToString();

                            Inner_Location_Target =
                                ER_Common.GenerateChr(Inner_Pivot_Range.X1)
                                + (Inner_Row_Start + 1).ToString()
                                + @":"
                                + ER_Common.GenerateChr(Inner_Pivot_Range.X1 + ((((Inner_Pivot_Range.X2 - Inner_Pivot_Range.X1) + 1) * Inner_PivotHeader_Count) - 1))
                                + (Inner_Row_End - 1).ToString();

                            Ws_Template.Range[Inner_Location_Source].Copy(Ws_Document.Range[Inner_Location_Target], XlPasteType.xlPasteFormats);
                        }

                        // - Bottom
                        Inner_Location_Source =
                        ER_Common.GenerateChr(Inner_Pivot_Range.X1)
                        + (Inner_Pivot_Range.Y1 + 2).ToString()
                        + @":"
                        + ER_Common.GenerateChr(Inner_Pivot_Range.X2)
                        + (Inner_Pivot_Range.Y1 + 2).ToString();

                        Inner_Location_Target =
                            ER_Common.GenerateChr(Inner_Pivot_Range.X1)
                            + (Inner_Row_End).ToString()
                            + @":"
                            + ER_Common.GenerateChr(Inner_Pivot_Range.X1 + ((((Inner_Pivot_Range.X2 - Inner_Pivot_Range.X1) + 1) * Inner_PivotHeader_Count) - 1))
                            + (Inner_Row_End).ToString();

                        Ws_Template.Range[Inner_Location_Source].Copy(Ws_Document.Range[Inner_Location_Target], XlPasteType.xlPasteFormats);
                    }

                    //Populate Pivot Items
                    // - Prepare DataTable Inner_Dt_Pivot_Items to be data container to input in excel
                    DataTable Inner_Dt_Pivot_Items = new DataTable();
                    var List_PivotTable_Fields =
                        (
                        from O in Document_PivotTables_Fields
                        where O.HasValue && O.Value.DataTable_Ct == Item_PivotTable.Ct
                        select O.Value);

                    foreach (Str_DataTable_Field Inner_Item_Pivot_Field in List_PivotTable_Fields)
                    {
                        foreach (DataRow Inner_Dr_Pivot_Desc in Ds_Source_Pivot_Desc.Tables[Ct_PivotTable].Rows)
                        {
                            Type Inner_Type = typeof(String);
                            foreach (DataColumn Inner_Dc in Ds_Source_Pivot.Tables[Ct_PivotTable].Columns)
                            {
                                if (Inner_Dc.ColumnName == Inner_Item_Pivot_Field.Name)
                                {
                                    Inner_Type = Inner_Dc.DataType;
                                    break;
                                }
                            }
                            Inner_Dt_Pivot_Items.Columns.Add(Inner_Item_Pivot_Field.Name + @"_" + Inner_Dr_Pivot_Desc["ID"].ToString(), Inner_Type);
                        }
                    }

                    foreach (DataRow Inner_Dr in Ds_Source.Tables[Ct_Table].Select())
                    {
                        DataRow Inner_Dr_New = Inner_Dt_Pivot_Items.NewRow();
                        Inner_Dt_Pivot_Items.Rows.Add(Inner_Dr_New);

                        var Inner_List_Dtpf =
                            (
                            from O in Document_PivotTables_Fields
                            where O.HasValue && O.Value.DataTable_Ct == Item_PivotTable.Ct
                            select O.Value);

                        foreach (Str_DataTable_Field Inner2_Dtf in Inner_List_Dtpf)
                        {
                            foreach (DataRow Inner2_Dr_Pivot_Desc in Ds_Source_Pivot_Desc.Tables[Ct_PivotTable].Rows)
                            {
                                DataRow[] Inner3_Arr_Dr =
                                    Ds_Source_Pivot.Tables[Ct_PivotTable].Select(
                                        @"ID = "
                                        + ER_Common.Convert_String(Inner2_Dr_Pivot_Desc["ID"], "0")
                                        + @" And "
                                        + Item_PivotTable.SourceKey
                                        + @" = "
                                        + Inner_Dr[Item_PivotTable.TargetKey]);
                                if (Inner3_Arr_Dr.Any())
                                { Inner_Dr_New[Inner2_Dtf.Name + @"_" + Inner2_Dr_Pivot_Desc["ID"].ToString()] = Inner3_Arr_Dr[0][Inner2_Dtf.Name]; }
                            }
                        }

                        Int32 Inner2_ItemCount = CreateExcelDocument_V3_CountItem(Item_Table, List_Tables, Ds_Source, Inner_Dr, false);
                        for (Int32 Inner3_Ct = 0; Inner3_Ct < Inner2_ItemCount; Inner3_Ct++)
                        {
                            DataRow Inner3_Dr_New = Inner_Dt_Pivot_Items.NewRow();
                            Inner_Dt_Pivot_Items.Rows.Add(Inner3_Dr_New);
                        }
                    }

                    // - Prepare Pivot Fields Definition

                    Int32 Inner_PivotTable_Length = (Inner_PivotHeader_Range.X2 - Inner_PivotHeader_Range.X1) + 1;
                    String[] Inner_Arr_PivotTable_Fields = new String[(Inner_PivotHeader_Count * Inner_PivotTable_Length) - 1];

                    for (Int32 Inner_Ct = 0; Inner_Ct < Inner_Arr_PivotTable_Fields.Length; Inner_Ct++)
                    { Inner_Arr_PivotTable_Fields[Inner_Ct] = ""; }

                    Int32 Outer_Ct2 = 0;
                    foreach (DataRow Inner_Dr_Pivot_Desc in Ds_Source_Pivot_Desc.Tables[Ct_PivotTable].Rows)
                    {
                        for (Int32 Inner_Ct = 0; Inner_Ct < Inner_PivotTable_Length; Inner_Ct++)
                        {
                            Str_DataTable_Field? Inner3_Dtf = Document_PivotTables_Fields.FirstOrDefault(O => O.Value.DataTable_Ct == Item_PivotTable.Ct && O.Value.Position == Inner_Ct);
                            if (Inner3_Dtf != null)
                            { Inner_Arr_PivotTable_Fields[Inner_Ct + (Outer_Ct2 * Inner_PivotTable_Length)] = Inner3_Dtf.Value.Name + @"_" + Inner_Dr_Pivot_Desc["ID"]; }
                        }
                        Outer_Ct2++;
                    }

                    // - Insert Data to Ws_Document

                    String Inner_Location_Pivot =
                        ER_Common.GenerateChr(Inner_Pivot_Range.X1)
                        + Inner_Ct_CurrentRow.ToString()
                        + @":"
                        + ER_Common.GenerateChr(Inner_Pivot_Range.X1 + ((((Inner_Pivot_Range.X2 - Inner_Pivot_Range.X1) + 1) * (Inner_PivotHeader_Count)) - 1))
                        + Inner_Ct_CurrentRow + (ItemCount - 1);

                    Ws_Document.Range[Inner_Location_Pivot].Value = ER_Common.ConvertDataTo2DimArray(Inner_Dt_Pivot_Items, Inner_Arr_PivotTable_Fields);

                    //[-]

                    Inner_Ct_CurrentRow = (Inner_Ct_CurrentRow + (ItemCount - 1)) + 1;

                    //Pivot Totals

                    var List_PivotTotals =
                        (from O in Document_PivotTotals
                         where O.HasValue && O.Value.Name == Item_PivotTable.Name
                         select O.Value).ToList();
                    foreach (Str_DataTable_Section Item_PivotTotal in List_PivotTotals)
                    {
                        String Inner_PivotTotal_Location = Item_PivotTotal.Location;
                        ER_Common.Str_ParsedExcelRange Inner_PivotTotal_Range = ER_Common.ParseExcelRange(Inner_PivotTotal_Location);
                        Int32 Inner_Ct_PivotTotal = Item_PivotTotal.Ct - 1;
                        Int32 Inner_PivotTotal_RowCount = Ds_Source_Pivot_Totals.Tables[Inner_Ct_PivotTotal].Rows.Count;
                        Int32 Inner_PivotTotal_Length = (Inner_PivotTotal_Range.X2 - Inner_PivotTotal_Range.X1) + 1;

                        String[] Inner_Arr_PivotTotal_Fields = new String[(Inner_PivotTotal_RowCount * Inner_PivotTotal_Length) - 1];
                        for (Int32 Inner_Ct = 0; Inner_Ct < Inner_Arr_PivotTotal_Fields.Count(); Inner_Ct++)
                        { Inner_Arr_PivotTotal_Fields[Inner_Ct] = ""; }

                        Int32 Inner_PivotTotal_Outer_Ct = 0;
                        foreach (DataRow Inner_Dr_PivotTotal in Ds_Source_Pivot_Totals.Tables[Inner_Ct_PivotTotal].Rows)
                        {
                            for (Int32 Inner_Ct = 0; Inner_Ct < Inner_PivotTotal_Length; Inner_Ct++)
                            {
                                Str_DataTable_Field? Inner2_Ptf = Document_PivotTotals_Fields.FirstOrDefault(O => O.Value.Ct == Item_PivotTotal.Ct && O.Value.Position == Inner_Ct);
                                if (Inner2_Ptf != null)
                                { Inner_Arr_PivotTotal_Fields[Inner_Ct + (Inner_PivotTotal_Outer_Ct * Inner_PivotTotal_Length)] = Inner2_Ptf.Value.Name + @"_" + ER_Common.Convert_String(Inner_Dr_PivotTotal["ID"]); }
                            }
                            Inner_PivotTotal_Outer_Ct++;
                        }

                        DataTable Inner_Dt_PivotTotal_Item = new DataTable();
                        var Inner_List_PivotTotal_Fields =
                            (from O in Document_PivotTotals_Fields
                             where O.HasValue && O.Value.DataTable_Ct == Item_PivotTotal.Ct
                             select O.Value).ToList();
                        foreach (Str_DataTable_Field Inner2_Ptf in Inner_List_PivotTotal_Fields)
                        {
                            foreach (DataRow Inner_Dr_PivotTotal in Ds_Source_Pivot_Totals.Tables[Inner_Ct_PivotTotal].Rows)
                            {
                                Type Inner_Type = typeof(String);
                                foreach (DataColumn Inner_Dc in Ds_Source_Pivot_Totals.Tables[Inner_Ct_PivotTotal].Columns)
                                {
                                    if (Inner_Dc.ColumnName == Inner2_Ptf.Name)
                                    {
                                        Inner_Type = Inner_Dc.DataType;
                                        break;
                                    }
                                }
                                Inner_Dt_PivotTotal_Item.Columns.Add(Inner2_Ptf.Name + @"_" + ER_Common.Convert_String(Inner_Dr_PivotTotal["ID"]), Inner_Type);
                            }
                        }

                        DataRow Inner_Dr_New = Inner_Dt_PivotTotal_Item.NewRow();
                        Inner_Dt_PivotTotal_Item.Rows.Add(Inner_Dr_New);

                        foreach (Str_DataTable_Field Inner2_Ptf in Inner_List_PivotTotal_Fields)
                        {
                            foreach (DataRow Inner_Dr_PivotTotal in Ds_Source_Pivot_Totals.Tables[Inner_Ct_PivotTotal].Rows)
                            {
                                DataRow[] Inner_Arr_Dr = Ds_Source_Pivot_Totals.Tables[Inner_Ct_PivotTotal].Select("ID = " + ER_Common.Convert_String(Inner_Dr_PivotTotal["ID"]));
                                if (Inner_Arr_Dr.Any())
                                { Inner_Dr_New[Inner2_Ptf.Name + @"_" + ER_Common.Convert_String(Inner_Dr_PivotTotal["ID"])] = Inner_Arr_Dr[0][Inner2_Ptf.Name]; }
                            }
                        }

                        //Table Formats and Borders
                        Inner_Location_Source =
                            ER_Common.GenerateChr(Inner_PivotTotal_Range.X1)
                            + Inner_PivotTotal_Range.Y1.ToString()
                            + @":"
                            + ER_Common.GenerateChr(Inner_PivotTotal_Range.X2)
                            + Inner_PivotTotal_Range.Y1.ToString();

                        Inner_Location_Target =
                            ER_Common.GenerateChr(Inner_PivotTotal_Range.X1)
                            + Inner_Ct_CurrentRow
                            + @":"
                            + ER_Common.GenerateChr(Inner_PivotTotal_Range.X1 + (((Inner_PivotTotal_Range.X2 - Inner_PivotTotal_Range.X1) + 1) * Inner_PivotTotal_RowCount) - 1)
                            + Inner_Ct_CurrentRow;

                        Ws_Template.Range[Inner_Location_Source].Copy(Ws_Document.Range[Inner_Location_Target], XlPasteType.xlPasteFormats);
                        Ws_Document.Range[Inner_Location_Target].Value = ER_Common.ConvertDataTo2DimArray(Inner_Dr_New, Inner_Arr_PivotTotal_Fields);
                    }
                }

                //[-]

                Item_Table.Items = Inner_ItemCount + ItemCount;
            }

            //Get the Table with Highest Item Count and add it to Ct_CurrentRow
            List_Tables = (
                from O in Document_Tables
                where O.HasValue && O.Value.IsSubTable == false
                orderby O.Value.Items ascending
                select O.Value).ToList();

            if (List_Tables.Any())
            { Ct_CurrentRow = Ct_CurrentRow + List_Tables.First().Items; }

            //Set Footer
            Str_Sections? Section_Footer = Document_Sections.FirstOrDefault(O => O.Value.Type.ToUpper() == "FOOTER");
            if (Section_Footer != null)
            {
                Str_Sections Inner_Section_Footer = Section_Header.Value;
                ER_Common.Str_ParsedExcelRange Inner_PR;
                String Inner_Location_Template;
                String Inner_Location_Document;
                Int32 Length;

                Inner_PR = ER_Common.ParseExcelRange(Inner_Section_Footer.Location);
                Length = (Inner_PR.Y2 - Inner_PR.Y1) + 1;
                Inner_Location_Template = @"A" + Inner_PR.Y1.ToString() + @":" + ER_Common.GenerateChr(Document_Settings.DocumentWidth) + Inner_PR.Y2.ToString();
                Inner_Location_Document = @"A" + Ct_CurrentRow.ToString() + @":" + ER_Common.GenerateChr(Document_Settings.DocumentWidth) + (Ct_CurrentRow + Length).ToString();

                Ws_Template.Range[Inner_Location_Template].Copy(Ws_Document.Range[Inner_Location_Document]);

                Ct_CurrentRow = Ct_CurrentRow + Length;
            }

            Ws_Document.Activate();
            Ws_Document.Range["A2"].Select();

            //Save the Document
            if (IsProtected)
            {
                String RandomPassword = Guid.NewGuid().ToString();
                Ws_Document.EnableSelection = XlEnableSelection.xlNoSelection;
                Ws_Document.Protect(RandomPassword);
                Wb_Document.Protect(RandomPassword);
            }

            if (SaveFileName == "")
            { SaveFileName = "Excel_File"; }

            XlFileFormat NxlFileFormat = ER_Common.ParseEnum<XlFileFormat>(FileFormat.ToString());
            Boolean Result = Ws_Document.SaveAs(SaveFileName, NxlFileFormat);
            return Result;

            //MemoryStream S_Output = new MemoryStream();
            //Ws_Document.SaveAs(S_Output, NxlFileFormat);

            //ER_Common.WriteFile(S_Output, SaveFileName);
        }

        //[-]

        static Int32 CreateExcelDocument_V3_CountItem(
            Str_DataTable Table
            , List<Str_DataTable> List_Tables
            , DataSet Ds_Source
            , DataRow Dr_SourceKey
            , Boolean IsIncludeCurrent = true)
        {
            String SourceKey = Table.SourceKey;
            String TargetKey = Table.TargetKey;

            String Condition = "";
            String SourceKey_ID = "0";

            if (SourceKey != "")
            {
                if (Dr_SourceKey != null)
                {
                    if (Dr_SourceKey.Table.Columns.Contains(SourceKey))
                    {
                        SourceKey_ID = ER_Common.Convert_String(Dr_SourceKey[SourceKey], "0");
                        Condition = TargetKey + @" = " + @"'" + SourceKey_ID + @"'";
                    }
                }
            }

            Int32 Rv = 0;
            Int32 Table_Ct = Table.Ct - 1;
            DataRow[] Arr_Data = Ds_Source.Tables[Table_Ct].Select(Condition);

            if (IsIncludeCurrent)
            { Rv = Rv + Arr_Data.Length; }

            var List = from O in List_Tables where O.GroupName == Table.Name select O;
            foreach (Str_DataTable Inner_Table in List)
            {
                foreach (DataRow Dr in Arr_Data)
                {
                    Int32 Inner_Rv = CreateExcelDocument_V3_CountItem(Inner_Table, List_Tables, Ds_Source, Dr, true);
                    Rv = Rv + Inner_Rv;
                }
            }

            return Rv;
        }

        static void CreateExcelDocument_V3_SubTable(
            DataSet Ds_Source
            , List<Str_DataTable?> List_Tables
            , List<Str_DataTable_Field?> List_Table_Fields
            , IWorksheet Ws_Template
            , IWorksheet Ws_Document
            , ref Int32 Ct_CurrentRow
            , List<Str_DataTable?> List_Tables_Group
            , DataRow Dr_SourceKey)
        {
            foreach (Str_DataTable Var_Table in List_Tables_Group)
            {
                Str_DataTable Table = Var_Table;

                ER_Common.Str_ParsedExcelRange PR = ER_Common.ParseExcelRange(Table.Location);
                Int32 Table_Ct = Table.Ct - 1;
                Int32 Ct_Items = CreateExcelDocument_V3_CountItem(
                    Table
                    , (from O in List_Tables select O.Value).ToList()
                    , Ds_Source
                    , Dr_SourceKey);

                if (Ct_Items > 0)
                {
                    //Set Formatting

                    Int32 Row_Start = Ct_CurrentRow;
                    Int32 Row_End = Row_Start + (Ct_Items - 1);

                    String Location_Source;
                    String Location_Target;

                    //Table Formats and Borders
                    if (Ct_Items == 1)
                    {
                        Location_Source = ER_Common.GenerateChr(PR.X1) + (PR.Y1 + 4).ToString() + @":" + ER_Common.GenerateChr(PR.X2) + (PR.Y1 + 4).ToString();
                        Location_Target = ER_Common.GenerateChr(PR.X1) + Row_End.ToString() + @":" + ER_Common.GenerateChr(PR.X2) + Row_End.ToString();
                        Ws_Template.Range[Location_Source].Copy(Ws_Document.Range[Location_Target], XlPasteType.xlPasteFormats);
                    }
                    else
                    {
                        // - Top
                        Location_Source = ER_Common.GenerateChr(PR.X1) + PR.Y1.ToString() + @":" + ER_Common.GenerateChr(PR.X2) + PR.Y1.ToString();
                        Location_Target = ER_Common.GenerateChr(PR.X1) + Row_Start.ToString() + @":" + ER_Common.GenerateChr(PR.X2) + Row_Start.ToString();
                        Ws_Template.Range[Location_Source].Copy(Ws_Document.Range[Location_Target], XlPasteType.xlPasteFormats);

                        // - Middle
                        if (Ct_Items > 2)
                        {
                            Location_Source = ER_Common.GenerateChr(PR.X1) + (PR.Y1 + 1).ToString() + @":" + ER_Common.GenerateChr(PR.X2) + (PR.Y1 + 1).ToString();
                            Location_Target = ER_Common.GenerateChr(PR.X1) + (Row_Start + 1).ToString() + @":" + ER_Common.GenerateChr(PR.X2) + (Row_End - 1).ToString();
                            Ws_Template.Range[Location_Source].Copy(Ws_Document.Range[Location_Target], XlPasteType.xlPasteFormats);
                        }

                        // - Bottom
                        Location_Source = ER_Common.GenerateChr(PR.X1) + (PR.Y1 + 2).ToString() + @":" + ER_Common.GenerateChr(PR.X2) + (PR.Y1 + 2).ToString();
                        Location_Target = ER_Common.GenerateChr(PR.X1) + Row_End.ToString() + @":" + ER_Common.GenerateChr(PR.X2) + Row_End.ToString();
                        Ws_Template.Range[Location_Source].Copy(Ws_Document.Range[Location_Target], XlPasteType.xlPasteFormats);
                    }


                    //Table Data

                    String[] Arr_Fields = new String[PR.X2 - PR.X1];
                    for (Int32 Inner_Ct = 0; Inner_Ct < Arr_Fields.Length; Inner_Ct++)
                    {
                        Str_DataTable_Field? Dtf = List_Table_Fields.First(O => O.Value.DataTable_Ct == Table.Ct && O.Value.Position == Inner_Ct);
                        if (Dtf != null)
                        { Arr_Fields[Inner_Ct] = Dtf.Value.Name; }
                        else
                        { Arr_Fields[Inner_Ct] = ""; }
                    }

                    String SourceKey = Table.SourceKey;
                    String TargetKey = Table.TargetKey;

                    String Condition = "";
                    String SourceKey_ID;

                    if (SourceKey != "")
                    {
                        if (Dr_SourceKey != null)
                        {
                            if (Dr_SourceKey.Table.Columns.Contains(SourceKey))
                            {
                                SourceKey_ID = ER_Common.Convert_String(Dr_SourceKey[SourceKey], "0");
                                Condition = TargetKey + @" = " + @"'" + SourceKey_ID + @"'";
                            }
                        }
                    }

                    DataRow[] Arr_Data = Ds_Source.Tables[Table_Ct].Select(Condition);
                    foreach (DataRow Dr in Arr_Data)
                    {
                        IRange R = Ws_Document.Range[ER_Common.GenerateChr(PR.X1) + Ct_CurrentRow.ToString() + @":" + ER_Common.GenerateChr(PR.X2) + Ct_CurrentRow.ToString()];
                        R.Value = ER_Common.ConvertDataTo2DimArray(Dr, Arr_Fields);
                        Ct_CurrentRow++;

                        var Inner_List_Table_Group =
                            (from O in List_Tables_Group
                             where O.Value.IsSubTable == true && O.Value.GroupName == Table.Name
                             select O).ToList();
                        if (Inner_List_Table_Group.Any())
                        {
                            CreateExcelDocument_V3_SubTable(
                                Ds_Source
                                , List_Tables
                                , List_Table_Fields
                                , Ws_Template
                                , Ws_Document
                                , ref Ct_CurrentRow
                                , Inner_List_Table_Group
                                , Dr);
                        }
                    }
                }
            }
        }

        //[-]

        static Str_Settings CreateExcelDocument_GetSettings(IWorksheet Ws_Parameters)
        {
            Str_Settings Rv_Settings = new Str_Settings();

            Int32[] LineInfo = CreateExcelDocument_ReadLineInfo(Ws_Parameters, CnsExcelKeyword_Settings, CnsExcelKeyword_Settings_End);
            Int32 Ct_Start = LineInfo[0];
            Int32 Ct_End = LineInfo[1];

            for (Int32 Ct = Ct_Start; Ct <= Ct_End; Ct++)
            {
                String ExcelText = Ws_Parameters.Range["A" + Ct.ToString()].Characters.Text;

                if ((Strings.InStr(ExcelText, "@DocumentLimit") > 0))
                { Rv_Settings.DocumentLimit = ER_Common.Convert_Int32(Strings.Mid(ExcelText, Strings.Len("@DocumentLimit") + 1)); }
                else if ((Strings.InStr(ExcelText, "@DocumentWidth") > 0))
                { Rv_Settings.DocumentWidth = ER_Common.Convert_Int32(Strings.Mid(ExcelText, Strings.Len("@DocumentWidth") + 1)); }
                else if ((Strings.InStr(ExcelText, "@IsRepeatHeader") > 0))
                { Rv_Settings.IsRepeatHeader = ER_Common.Convert_Boolean(Strings.Mid(ExcelText, Strings.Len("@IsRepeatHeader") + 1)); }
            }

            return Rv_Settings;
        }

        static List<Str_Sections?> CreateExcelDocument_GetSections(IWorksheet Ws_Parameters)
        {
            List<Str_Sections?> List_Sections = new List<Str_Sections?>();

            Int32[] LineInfo = CreateExcelDocument_ReadLineInfo(Ws_Parameters, CnsExcelKeyword_Sections, CnsExcelKeyword_Sections_End);
            Int32 Ct_Start = LineInfo[0];
            Int32 Ct_End = LineInfo[1];

            Int32 Ct_Sections = 0;

            for (Int32 Ct = Ct_Start; Ct <= Ct_End; Ct++)
            {
                String ExcelText = Ws_Parameters.Range["A" + Ct.ToString()].Characters.Text;
                String Type = "";
                String Location = "";

                if (!(Strings.InStr(ExcelText, "[") > 0))
                {
                    try
                    {
                        Ct_Sections++;

                        Type = Strings.Mid(ExcelText, 1, (Strings.InStr(ExcelText, " ")) - 1);
                        Location = Strings.Mid(ExcelText, (Strings.InStr(ExcelText, " ")) + 1);

                        List_Sections.Add(
                            new Str_Sections()
                            {
                                Ct = Ct_Sections,
                                Type = Type.ToUpper(),
                                Location = Location.ToUpper()
                            });
                    }
                    catch
                    { throw new Exception(@"Invalid Syntax in [#]Sections."); }
                }
            }

            return List_Sections;
        }

        static List<Str_Parameters?> CreateExcelDocument_GetParameters(IWorksheet Ws_Parameters, List<ER_Common.Str_Parameter> List_ErcSp)
        {
            List<Str_Parameters?> List_Sp = new List<Str_Parameters?>();

            Int32[] LineInfo = CreateExcelDocument_ReadLineInfo(Ws_Parameters, CnsExcelKeyword_Parameters, CnsExcelKeyword_Parameters_End);
            Int32 Ct_Start = LineInfo[0];
            Int32 Ct_End = LineInfo[1];

            for (Int32 Ct = Ct_Start; Ct <= Ct_End; Ct++)
            {
                String ExcelText = Ws_Parameters.Range["A" + Ct.ToString()].Characters.Text;
                String Parameter_Name;
                String Parameter_Type;

                if ((Strings.InStr(ExcelText, "@") > 0))
                {
                    try
                    {
                        Parameter_Name = Strings.Mid(ExcelText, Strings.Len("@") + 1, (Strings.InStr(ExcelText, " ") - Strings.Len("@")) - 1);
                        Parameter_Type = Strings.Mid(ExcelText, Strings.InStr(ExcelText, " ") + 1);

                        Str_Parameters Sp = new Str_Parameters();
                        Sp.Name = Parameter_Name;
                        Sp.Type = Parameter_Type;
                        Sp.Value = (from O in List_ErcSp where O.Name == Sp.Name select O.Value).FirstOrDefault().ToString();

                        List_Sp.Add(Sp);
                    }
                    catch { }
                }
            }

            return List_Sp;
        }

        static List<Str_DataTable?> CreateExcelDocument_GetDataTables(IWorksheet Ws_Parameters)
        {
            List<Str_DataTable?> List_Dt = new List<Str_DataTable?>();

            Int32[] LineInfo = CreateExcelDocument_ReadLineInfo(Ws_Parameters, CnsExcelKeyword_DataTable, CnsExcelKeyword_DataTable_End);
            Int32 Ct_Start = LineInfo[0];
            Int32 Ct_End = LineInfo[1];

            if (Ct_Start == 0 && Ct_End == 0)
            { throw new Exception("Invalid Syntax in [#]DataTable"); }

            Int32 Ct_DataTable = 0;
            for (Int32 Ct = Ct_Start; Ct <= Ct_End; Ct++)
            {
                String ExcelText = "";

                try { ExcelText = Ws_Parameters.Range["A" + Ct.ToString()].Characters.Text; }
                catch { }

                String DataTable_Name = "";
                String DataTable_GroupName = "";
                String DataTable_SourceKey = "";
                String DataTable_TargetKey = "";
                String DataTable_Location = "";

                if (!(Strings.InStr(ExcelText, "[") > 0))
                {
                    try
                    {
                        Ct_DataTable++;

                        String[] Arr_ExcelText = Strings.Split(ExcelText, " ");

                        DataTable_Name = Arr_ExcelText[0];
                        DataTable_Location = Arr_ExcelText[1];

                        try
                        {
                            if (Arr_ExcelText.Length > 2)
                            {
                                DataTable_GroupName = Arr_ExcelText[3];
                                DataTable_SourceKey = Arr_ExcelText[4];
                                DataTable_TargetKey = Arr_ExcelText[5];
                            }
                        }
                        catch { }

                        Str_DataTable Dt_New = new Str_DataTable();
                        Dt_New.Ct = Ct_DataTable;
                        Dt_New.Name = DataTable_Name;
                        Dt_New.Location = DataTable_Location;

                        if (DataTable_GroupName.Trim() != "")
                        {
                            Dt_New.GroupName = DataTable_GroupName;
                            Dt_New.SourceKey = DataTable_SourceKey;
                            Dt_New.TargetKey = DataTable_TargetKey;
                            Dt_New.IsSubTable = true;
                        }

                        List_Dt.Add(Dt_New);
                    }
                    catch
                    { throw new Exception(@"Invalid Syntax in [#]DataTable."); }
                }
            }

            return List_Dt;
        }

        static List<Str_DataTable_Section?> CreateExcelDocument_GetDataTables_Headers(IWorksheet Ws_Parameters)
        { return CreateExcelDocument_GetSections(Ws_Parameters, CnsExcelKeyword_DataTable_Header, CnsExcelKeyword_DataTable_Header_End); }

        static List<Str_DataTable_Section?> CreateExcelDocument_GetDataTables_Footers(IWorksheet Ws_Parameters)
        { return CreateExcelDocument_GetSections(Ws_Parameters, CnsExcelKeyword_DataTable_Footer, CnsExcelKeyword_DataTable_Footer_End); }

        static List<Str_DataTable?> CreateExcelDocument_GetDataTables_PivotTables(IWorksheet Ws_Parameters)
        {
            List<Str_DataTable?> List_Dtp = new List<Str_DataTable?>();

            Int32[] LineInfo = CreateExcelDocument_ReadLineInfo(Ws_Parameters, CnsExcelKeyword_DataTable_Pivot, CnsExcelKeyword_DataTable_Pivot_End);
            Int32 Ct_Start = LineInfo[0];
            Int32 Ct_End = LineInfo[1];

            if (Ct_Start == 0 && Ct_End == 0)
            { return List_Dtp; }

            Int32 DataTable_Ct = 0;
            for (Int32 Ct = Ct_Start; Ct <= Ct_End; Ct++)
            {
                String ExcelText = "";

                try { ExcelText = Ws_Parameters.Range["A" + Ct.ToString()].Characters.Text; }
                catch { }

                String DataTable_Name = "";
                String DataTable_ParentName = "";
                String DataTable_SourceKey = "";
                String DataTable_TargetKey = "";
                String DataTable_Location = "";

                if (!(Strings.InStr(ExcelText, "[") > 0))
                {
                    try
                    {
                        DataTable_Ct++;

                        String[] Arr_ExcelText = Strings.Split(ExcelText, " ");

                        DataTable_Name = Arr_ExcelText[0];
                        DataTable_Location = Arr_ExcelText[1];

                        try
                        {
                            if (Arr_ExcelText.Length > 2)
                            {
                                DataTable_ParentName = Arr_ExcelText[3];
                                DataTable_SourceKey = Arr_ExcelText[4];
                                DataTable_TargetKey = Arr_ExcelText[5];
                            }
                        }
                        catch { }

                        Str_DataTable Dt_New = new Str_DataTable();
                        Dt_New.Ct = DataTable_Ct;
                        Dt_New.Name = DataTable_Name;
                        Dt_New.Location = DataTable_Location;
                        Dt_New.ParentName = DataTable_ParentName;
                        Dt_New.SourceKey = DataTable_SourceKey;
                        Dt_New.TargetKey = DataTable_TargetKey;

                        List_Dtp.Add(Dt_New);
                    }
                    catch
                    { throw new Exception(@"Invalid Syntax in [#]DataTable_Pivot."); }
                }
            }

            return List_Dtp;
        }

        static List<Str_DataTable_Section?> CreateExcelDocument_GetDataTables_PivotTables_Headers(IWorksheet Ws_Parameters)
        { return CreateExcelDocument_GetSections(Ws_Parameters, CnsExcelKeyword_DataTable_Pivot_Header, CnsExcelKeyword_DataTable_Pivot_Header_End); }

        static List<Str_DataTable_Section?> CreateExcelDocument_GetDataTables_PivotTables_Totals(IWorksheet Ws_Parameters)
        { return CreateExcelDocument_GetSections(Ws_Parameters, CnsExcelKeyword_DataTable_Pivot_Totals, CnsExcelKeyword_DataTable_Pivot_Totals_End); }

        static Int32[] CreateExcelDocument_ReadLineInfo(IWorksheet Ws_Source, String Line_Start, String Line_End)
        {
            Int32 Ct_Start = 0;
            Int32 Ct_End = 0;
            Int32 Ct = 0;

            for (Ct = 1; Ct <= CnsExcelMaxHeight; Ct++)
            {
                String ExcelLine = Ws_Source.Range["A" + Ct.ToString()].Characters.Text;
                if (ExcelLine == Line_Start)
                { Ct_Start = Ct; }
                else if (ExcelLine == Line_End)
                {
                    Ct_End = Ct;
                    break;
                }
            }

            Int32[] Rv = new Int32[2];
            Rv[0] = Ct_Start;
            Rv[1] = Ct_End;

            return Rv;
        }

        static List<Str_DataTable_Section?> CreateExcelDocument_GetSections(IWorksheet Ws_Parameters, String Line_Start, String Line_End)
        {
            List<Str_DataTable_Section?> List_Dts = new List<Str_DataTable_Section?>();

            Int32[] LineInfo = CreateExcelDocument_ReadLineInfo(Ws_Parameters, Line_Start, Line_End);
            Int32 Ct_Start = LineInfo[0];
            Int32 Ct_End = LineInfo[1];

            if (Ct_Start == 0 && Ct_End == 0)
            { throw new Exception(@"Invalid Syntax in " + Line_Start + @"."); }

            Int32 DataTable_Ct = 0;
            for (Int32 Ct = Ct_Start; Ct <= Ct_End; Ct++)
            {
                String ExcelText = "";

                try { ExcelText = Ws_Parameters.Range["A" + Ct.ToString()].Characters.Text; }
                catch { }

                String Name = "";
                String Location = "";

                if (!(Strings.InStr(ExcelText, "[") > 0))
                {
                    try
                    {
                        DataTable_Ct++;

                        Name = Strings.Mid(ExcelText, 1, Strings.InStr(ExcelText, " ") - 1);
                        Location = Strings.Mid(ExcelText, Strings.InStr(ExcelText, " ") + 1);

                        List_Dts.Add(
                            new Str_DataTable_Section()
                            {
                                Ct = DataTable_Ct,
                                Name = Name,
                                Location = Location
                            });
                    }
                    catch
                    { throw new Exception(@"Invalid Syntax in " + Line_Start + "."); }
                }
            }

            return List_Dts;
        }

        static List<Str_DataTable_Field?> CreateExcelDocument_GetFields(List<Str_DataTable_Section?> List_Section, IWorksheet Ws_Template)
        {
            var List_Table = (
                from O in
                    (from O in List_Section
                     where O.HasValue
                     select O.Value)
                select new Str_DataTable?(
                    new Str_DataTable() { Ct = O.Ct, Name = O.Name, Location = O.Location })).ToList();

            return CreateExcelDocument_GetFields(List_Table, Ws_Template);
        }

        static List<Str_DataTable_Field?> CreateExcelDocument_GetFields(List<Str_DataTable?> List_Table, IWorksheet Ws_Template)
        {
            List<Str_DataTable_Field?> List_Dtf = new List<Str_DataTable_Field?>();

            var List = from O in List_Table where O.HasValue orderby O.Value.Ct select O.Value;
            foreach (var Item in List)
            {
                ER_Common.Str_ParsedExcelRange PR = ER_Common.ParseExcelRange(Item.Location);
                Int32 Table_Width = PR.X2 - PR.X1;
                Int32 Table_Pivot_Field_Ct = 0;

                for (Int32 Ct = 0; Ct <= Table_Width; Ct++)
                {
                    String Excel_Text = Ws_Template.Range[ER_Common.GenerateChr(PR.X1 + Ct) + PR.Y1.ToString()].Characters.Text;

                    if (Strings.InStr(Excel_Text, "[") > 0)
                    {
                        Table_Pivot_Field_Ct++;

                        String FieldName =
                            Strings.Mid(
                                Excel_Text
                                , Strings.InStr(Excel_Text, "[") + 1
                                , (Strings.InStrRev(Excel_Text, "]") - Strings.Len("]")) - 1);

                        Str_DataTable_Field Dtf = new Str_DataTable_Field();
                        Dtf.Ct = Table_Pivot_Field_Ct;
                        Dtf.DataTable_Ct = Item.Ct;
                        Dtf.Name = FieldName;
                        Dtf.Position = Ct;

                        List_Dtf.Add(Dtf);
                    }
                }
            }

            return List_Dtf;
        }

        static void CreateExcelDocument_CheckStringFields(DataSet Ds_Source)
        {
            foreach (DataTable Inner_Dt in Ds_Source.Tables)
            {
                foreach (DataRow Inner_Dr in Inner_Dt.Rows)
                {
                    foreach (DataColumn Inner_Dc in Inner_Dt.Columns)
                    {
                        if (Inner_Dc.DataType == typeof(String))
                        {
                            Inner_Dr[Inner_Dc.ColumnName] = Strings.Replace(ER_Common.Convert_String(Inner_Dr[Inner_Dc.ColumnName]), Constants.vbTab, "");
                            Inner_Dr[Inner_Dc.ColumnName] = Strings.Replace(ER_Common.Convert_String(Inner_Dr[Inner_Dc.ColumnName]), Constants.vbCrLf, "");
                            Inner_Dr[Inner_Dc.ColumnName] = Strings.Replace(ER_Common.Convert_String(Inner_Dr[Inner_Dc.ColumnName]), Constants.vbCr, "");
                            Inner_Dr[Inner_Dc.ColumnName] = Strings.Replace(ER_Common.Convert_String(Inner_Dr[Inner_Dc.ColumnName]), Constants.vbLf, "");
                        }
                    }
                }
            }
        }

        #endregion
    }
}
