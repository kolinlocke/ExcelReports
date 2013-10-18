#Region " Imports "

Imports System
Imports System.Data
Imports System.Diagnostics
Imports ExcelReports.Methods
Imports NativeExcel

#End Region

Friend Class Methods_NativeExcel

    Shared Sub NativeExcel_CreateExcel(ByVal Dt As DataTable, ByVal Columns_Obj As ClsNativeExcel_Columns, ByVal SaveFileName As String, Optional ByVal FileFormat As NativeExcel.XlFileFormat = XlFileFormat.xlNormal)
        Dim owbook As IWorkbook = NativeExcel.Factory.CreateWorkbook
        Dim owsheet As IWorksheet = owbook.Worksheets.Add

        '[Setup Header]
        Dim RowCt As Int32 = 1
        Dim ColCt As Int32 = 1

        For Each Obj As ClsNativeExcel_Columns_Obj In Columns_Obj.pObj
            owsheet.Cells(RowCt, ColCt).Value = Obj.mFieldDesc
            owsheet.Cells(RowCt, ColCt).Font.Bold = True
            Dim Inner_ExRange As IRange = owsheet.Range(GenerateChr(ColCt) & RowCt & ":" & GenerateChr(ColCt) & (RowCt + Dt.Rows.Count))
            Inner_ExRange.NumberFormat = Obj.mNumberFormat
            ColCt = ColCt + 1
        Next

        RowCt = RowCt + 1
        ColCt = 1

        Dim ItemCt As Int32 = 0
        ItemCt = Dt.Rows.Count
        If ItemCt > 0 Then
            ItemCt = ItemCt - 1
        End If

        Dim ExRange As IRange = owsheet.Range(GenerateChr(ColCt) & RowCt & ":" & GenerateChr(ColCt + Columns_Obj.pObj.Count) & (RowCt + (ItemCt)))
        ExRange.Value = ConvertDataTo2DimArray(Dt, "", Columns_Obj.pFieldName)
        owsheet.Range("A1:IV65536").Autofit()

        If SaveFileName = "" Then
            SaveFileName = "Excel_File"
        End If

        owsheet.Cells("A1:IV1").Insert(XlInsertShiftDirection.xlShiftDown)
        owsheet.Cells("A1").RowHeight = 0
        owsheet.Range("A2:A2").Select()

        '[-]

        owbook.SaveAs(SaveFileName, FileFormat)

    End Sub

    Shared Sub NativeExcel_CreateExcelDocument(ByVal ExcelTemplateFileName As String, ByVal pParameters() As String, ByVal pParametersValue() As Object, ByVal pDs As DataSet, ByVal pDs_Pivot As DataSet, ByVal pDs_Pivot_Desc As DataSet, ByVal pDs_Pivot_Totals As DataSet, ByVal SaveFileName As String, Optional ByVal IsProtected As Boolean = False, Optional ByVal FileFormat As NativeExcel.XlFileFormat = XlFileFormat.xlNormal)
        Try

            Dim TemplateVersion As Int32 = 0
            Dim owbook_Template As IWorkbook = NativeExcel.Factory.OpenWorkbook(ExcelTemplateFileName)
            Dim owsheet_Parameters As IWorksheet = owbook_Template.Worksheets.Item("Parameters")

            Dim CtStart As Int32 = 0
            Dim CtEnd As Int32 = 0
            Dim Ct As Int32 = 0

            '[Get Settings]
            For Ct = 1 To 65536
                Select Case owsheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                    Case "[#]Settings"
                        CtStart = Ct
                    Case "[#]End_Settings"
                        CtEnd = Ct
                        Exit For
                End Select
            Next

            For Ct = CtStart To CtEnd
                Dim ExcelText As String = ""
                ExcelText = owsheet_Parameters.Range("A" & Ct.ToString).Characters.Text

                Select Case True
                    Case InStr(ExcelText, "@TemplateVersion")
                        TemplateVersion = CType(Mid(ExcelText, Len("@TemplateVersion") + 1).Trim, Int32)
                End Select
            Next

            Select Case TemplateVersion
                Case 1
                    'Under Construction
                Case 2
                    NativeExcel_CreateExcelDocument_V2(ExcelTemplateFileName, pParameters, pParametersValue, pDs, SaveFileName, IsProtected, FileFormat)
                Case 3
                    NativeExcel_CreateExcelDocument_V3(ExcelTemplateFileName, pParameters, pParametersValue, pDs, pDs_Pivot, pDs_Pivot_Desc, pDs_Pivot_Totals, SaveFileName, IsProtected, FileFormat)
            End Select

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Shared Sub NativeExcel_CreateExcelDocument(ByVal ExcelTemplateFileName As String, ByVal Sp As ClsParameters, ByVal pDs As DataSet, ByVal pDs_Pivot As DataSet, ByVal pDs_Pivot_Desc As DataSet, ByVal pDs_Pivot_Totals As DataSet, ByVal SaveFileName As String, Optional ByVal IsProtected As Boolean = False, Optional ByVal FileFormat As NativeExcel.XlFileFormat = XlFileFormat.xlNormal)
        NativeExcel_CreateExcelDocument(ExcelTemplateFileName, Sp.pName, Sp.pValue, pDs, pDs_Pivot, pDs_Pivot_Desc, pDs_Pivot_Totals, SaveFileName, IsProtected, FileFormat)
    End Sub

    Shared Sub NativeExcel_CreateExcelDocument_V2(ByVal ExcelTemplateFileName As String, ByVal pParameters() As String, ByVal pParametersValue() As Object, ByVal pDs As DataSet, ByVal SaveFileName As String, Optional ByVal IsProtected As Boolean = False, Optional ByVal FileFormat As NativeExcel.XlFileFormat = XlFileFormat.xlNormal)
        Try

            'Dim SbExcelData As New System.Text.StringBuilder
            Dim ExRange As NativeExcel.IRange = Nothing

            Dim owbook_Document As IWorkbook  '--#Workbook for Document
            Dim owbook_Template As IWorkbook  '--#Workbook for Template Data

            Dim owsheet_Document As IWorksheet
            Dim owsheet_Parameters As IWorksheet
            Dim owsheet_Template As IWorksheet

            Try

                owbook_Template = NativeExcel.Factory.OpenWorkbook(ExcelTemplateFileName)
                owsheet_Parameters = owbook_Template.Worksheets.Item("Parameters")
                owsheet_Template = owbook_Template.Worksheets.Item("Template")

                'Clear all String Fields of all TAB and RETURN chars
                For Each Dt As DataTable In pDs.Tables
                    For Each InnerDr As DataRow In Dt.Rows
                        For Each InnerDc As DataColumn In Dt.Columns
                            If InnerDc.DataType.Name = GetType(System.String).Name Then
                                InnerDr.Item(InnerDc) = Replace(IsNull(InnerDr.Item(InnerDc), ""), vbTab, "")
                                InnerDr.Item(InnerDc) = Replace(IsNull(InnerDr.Item(InnerDc), ""), vbCrLf, "")
                            End If
                        Next
                    Next
                Next

                Dim CtStart As Int32 = 0
                Dim CtEnd As Int32 = 0
                Dim Ct As Int32 = 0

                '--#Get Settings
                Dim DocumentLimit As Int32 = 0
                Dim DocumentWidth As Int32 = 0

                For Ct = 1 To 65536
                    Select Case owsheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                        Case "[#]Settings"
                            CtStart = Ct
                        Case "[#]End_Settings"
                            CtEnd = Ct
                            Exit For
                    End Select
                Next

                For Ct = CtStart To CtEnd
                    Dim ExcelText As String = ""
                    ExcelText = owsheet_Parameters.Range("A" & Ct.ToString).Characters.Text

                    Select Case True
                        Case InStr(ExcelText, "@DocumentLimit")
                            DocumentLimit = CType(Mid(ExcelText, Len("@DocumentLimit") + 1).Trim, Int32)
                        Case InStr(ExcelText, "@DocumentWidth")
                            DocumentWidth = CType(Mid(ExcelText, Len("@DocumentWidth") + 1).Trim, Int32)
                    End Select
                Next
                '--#End Get Settings

                '--#Get Parameters
                CtStart = 0
                CtEnd = 0
                Ct = 0

                For Ct = 1 To 65536
                    Select Case owsheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                        Case "[#]Parameters"
                            CtStart = Ct
                        Case "[#]End_Parameters"
                            CtEnd = Ct
                            Exit For
                    End Select
                Next

                Dim DtParameters As New DataTable
                DtParameters.Columns.Add("ParameterName", GetType(String))
                DtParameters.Columns.Add("ParameterType", GetType(String))
                DtParameters.Columns.Add("ParameterValue", GetType(String))

                For Ct = CtStart To CtEnd
                    Dim ExcelText As String = ""
                    ExcelText = owsheet_Parameters.Range("A" & Ct.ToString).Characters.Text

                    Dim ParameterName As String = ""
                    Dim ParameterType As String = ""

                    Select Case True
                        Case InStr(ExcelText, "@") > 0
                            ParameterName = Mid(ExcelText, Len("@") + 1, (InStr(ExcelText, " ") - Len("@")) - 1)
                            ParameterType = Mid(ExcelText, (InStr(ExcelText, " ") + 1))

                            Dim NewRow As DataRow
                            NewRow = DtParameters.NewRow
                            NewRow.Item("ParameterName") = ParameterName
                            NewRow.Item("ParameterType") = ParameterType

                            DtParameters.Rows.Add(NewRow)
                    End Select
                Next

                Ct = 0
                For Ct = 0 To pParameters.Length - 1
                    Dim Adr() As DataRow
                    Adr = DtParameters.Select("ParameterName = '" & pParameters(Ct) & "'")
                    If Adr.Length > 0 Then
                        Adr(0).Item("ParameterValue") = pParametersValue(Ct)
                    End If
                Next
                '--#End Get Parameters

                '--#Get DataTable Parameters
                CtStart = 0
                CtEnd = 0
                Ct = 0

                For Ct = 1 To 65536
                    Select Case owsheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                        Case "[#]DataTable"
                            CtStart = Ct
                        Case "[#]End_DataTable"
                            CtEnd = Ct
                            Exit For
                    End Select
                Next

                Dim DtDataTable As New DataTable
                DtDataTable.Columns.Add("Ct", GetType(System.Int32))
                DtDataTable.Columns.Add("Name", GetType(System.String))
                DtDataTable.Columns.Add("Location", GetType(System.String))
                DtDataTable.Columns.Add("Items", GetType(System.Int32))

                Dim DataTable_Ct As Int32 = 0

                For Ct = CtStart To CtEnd
                    Dim ExcelText As String = ""
                    ExcelText = owsheet_Parameters.Range("A" & Ct.ToString).Characters.Text

                    Dim DataTable_Name As String = ""
                    Dim DataTable_Location As String = ""

                    Select Case True
                        Case InStr(ExcelText, "[") > 0
                        Case Else
                            Try
                                DataTable_Name = Mid(ExcelText, 1, (InStr(ExcelText, " ")) - 1)
                                DataTable_Location = Mid(ExcelText, (InStr(ExcelText, " ") + 1))

                                DataTable_Ct = DataTable_Ct + 1

                                Dim NewRow As DataRow
                                NewRow = DtDataTable.NewRow
                                NewRow.Item("Ct") = DataTable_Ct
                                NewRow.Item("Name") = DataTable_Name
                                NewRow.Item("Location") = DataTable_Location

                                Try
                                    NewRow.Item("Items") = pDs.Tables(DataTable_Ct - 1).Rows.Count
                                Catch ex As Exception
                                    Debug.Print(ex.ToString)
                                    Throw New Exception("Exception Code 1010. DataTable Supplied is less than specfied in the template.")
                                End Try

                                DtDataTable.Rows.Add(NewRow)

                            Catch ex As Exception
                                Debug.Print(ex.ToString)
                                Throw New Exception(ex.Message & vbCrLf & "Exception Code 1003. Invalid Syntax in [#]DataTable.")
                            End Try

                    End Select
                Next
                '--#End Get DataTable Parameters

                '--#Prepare DataTable Insertion
                Dim DtDataTable_Fields As New DataTable
                DtDataTable_Fields.Columns.Add("Ct", GetType(System.Int32))
                DtDataTable_Fields.Columns.Add("DtDataTable_Ct", GetType(System.Int32))
                DtDataTable_Fields.Columns.Add("Name", GetType(System.String))
                DtDataTable_Fields.Columns.Add("Position", GetType(System.Int32))

                Dim StFields As String = ""
                Dim ArStFields() As String = Nothing
                ReDim ArStFields(-1)

                Dim AdrDtDataTable() As DataRow
                AdrDtDataTable = DtDataTable.Select("", "Ct")
                If AdrDtDataTable.Length > 0 Then
                    For Each Dr As DataRow In AdrDtDataTable

                        Ct = 0
                        StFields = ""
                        Dim Delimiter As String = ""

                        Dim DataTable_Width As Int32 = 0
                        Dim Excel_Range() As Int32 = ParseExcelRange(Dr.Item("Location"))

                        DataTable_Width = Excel_Range(2) - Excel_Range(0)

                        Dim DtDataTable_Fields_Ct As Int32 = 0

                        For Ct = 0 To DataTable_Width
                            Dim ExcelText As String = ""
                            Try
                                ExcelText = owsheet_Template.Range(GenerateChr(Excel_Range(0) + Ct) & Excel_Range(1)).Characters.Text
                            Catch
                            End Try

                            Dim Field As String = ""
                            Select Case True
                                Case InStr(ExcelText, "[") > 0
                                    Field = Mid(ExcelText, InStr(ExcelText, "[") + 1, (InStrRev(ExcelText, "]") - Len("]")) - 1)

                                    DtDataTable_Fields_Ct = DtDataTable_Fields_Ct + 1

                                    Dim NewRow As DataRow
                                    NewRow = DtDataTable_Fields.NewRow
                                    NewRow.Item("Ct") = DtDataTable_Fields_Ct
                                    NewRow.Item("DtDataTable_Ct") = Dr.Item("Ct")
                                    NewRow.Item("Name") = Field
                                    NewRow.Item("Position") = Ct

                                    DtDataTable_Fields.Rows.Add(NewRow)

                            End Select
                        Next
                    Next
                End If
                '--#End Prepare DataTable Insertion

                '--#Input Data From Parameters
                Dim Inner0Ct As Int32 = 0
                Dim Inner1Ct As Int32 = 0

                For Inner0Ct = 0 To DocumentLimit - 1
                    For Inner1Ct = 0 To DocumentWidth - 1
                        Dim ExcelText As String = ""
                        Try
                            ExcelText = owsheet_Template.Range(GenerateChr(Inner1Ct + 1) & (Inner0Ct + 1)).Characters.Text
                        Catch
                        End Try

                        Dim Parameter As String = ""
                        Select Case True
                            Case InStr(ExcelText, "[@") > 0
                                Parameter = Mid(ExcelText, InStr(ExcelText, "[@") + Len("[@"), (InStrRev(ExcelText, "]") - Len("]")) - Len("[@"))
                            Case Else
                                Parameter = ""
                        End Select

                        Dim InnerAdr() As DataRow
                        InnerAdr = DtParameters.Select("ParameterName = '" & Parameter & "'")
                        If InnerAdr.Length > 0 Then
                            Dim ParameterValue As Object
                            ParameterValue = InnerAdr(0).Item("ParameterValue")

                            Dim Ar_TypeCode() As Integer
                            Ar_TypeCode = [Enum].GetValues(GetType(TypeCode))

                            For Each Tc As Integer In Ar_TypeCode
                                Dim TcName As String = ""
                                TcName = [Enum].GetName(GetType(TypeCode), Tc)
                                If TcName = InnerAdr(0).Item("ParameterType") Then
                                    Try
                                        ParameterValue = Convert.ChangeType(ParameterValue, Tc)
                                    Catch
                                    End Try
                                    Exit For
                                End If
                            Next

                            owsheet_Template.Cells(Inner0Ct + 1, Inner1Ct + 1).Value = ParameterValue

                        End If

                    Next
                Next
                '--#End Input Data From Parameters

                '--#Document Generation
                Dim DtDocumentParameters As New DataTable
                DtDocumentParameters.Columns.Add("Sheet", GetType(System.Int32))
                DtDocumentParameters.Columns.Add("Page", GetType(System.Int32))

                For Each Row As DataRow In DtDataTable.Rows
                    DtDocumentParameters.Columns.Add("Dt_" & Row.Item("Ct"), GetType(System.Int32))
                Next

                Dim IsEndReached As Boolean = False
                Dim CtPages As Int32 = 0
                Dim CtSheets As Int32 = 0

                While Not IsEndReached
                    Dim IsItem As Boolean = False

                    For Each Row As DataRow In DtDataTable.Rows
                        If Row.Item("Items") > 0 Then
                            IsItem = True
                            Exit For
                        End If
                    Next

                    If IsItem Then
                        Dim NewRow As DataRow
                        NewRow = DtDocumentParameters.NewRow

                        For Each Row As DataRow In DtDataTable.Rows
                            Dim CtTemp As Int32 = 0
                            Dim Excel_Range() As Int32
                            Excel_Range = ParseExcelRange(Row.Item("Location"))

                            Dim DataTable_Limit As Int32 = 0
                            DataTable_Limit = Math.Abs(Excel_Range(3) - Excel_Range(1)) + 1

                            If Row.Item("Items") > DataTable_Limit Then
                                CtTemp = DataTable_Limit
                            Else
                                CtTemp = Row.Item("Items")
                            End If

                            Row.Item("Items") = Row.Item("Items") - CtTemp

                            NewRow.Item("Dt_" & Row.Item("Ct")) = CtTemp

                        Next

                        NewRow.Item("Page") = CtPages
                        NewRow.Item("Sheet") = CtSheets
                        DtDocumentParameters.Rows.Add(NewRow)

                    Else
                        IsEndReached = True
                        Exit While
                    End If

                    CtPages = CtPages + 1

                    'If (CtPages * DocumentLimit) > cnsExcelHeightLimit Then
                    '    CtSheets = CtSheets + 1
                    '    CtPages = 0
                    'End If

                End While

                '[-]

                owbook_Document = NativeExcel.Factory.OpenWorkbook(ExcelTemplateFileName)
                For Each Ws As IWorksheet In owbook_Document.Worksheets
                    Select Case Ws.Name
                        Case "Template"
                        Case Else
                            Ws.Delete()
                    End Select
                Next

                owsheet_Document = owbook_Document.Worksheets.Item("Template")
                owsheet_Document.Name = "Document"
                'Dim ExRange As Excel.Range = Nothing
                ExRange = Nothing
                Dim AdrPage() As DataRow
                AdrPage = DtDocumentParameters.Select("", "Page")

                For Each DrPage As DataRow In AdrPage
                    Dim Page_TopLimit As Int32 = 0
                    Dim Page_BottomLimit As Int32 = 0

                    Page_TopLimit = DrPage.Item("Page") * DocumentLimit
                    Page_BottomLimit = (DrPage.Item("Page") + 1) * DocumentLimit

                    'Copy From Template
                    Dim Location As String = ""
                    Location = "A1:" & GenerateChr(DocumentWidth) & DocumentLimit.ToString

                    ExRange = owsheet_Template.Range(Location)
                    ExRange.Copy(owsheet_Document.Range("A" & (Page_TopLimit + 1).ToString & ":" & GenerateChr(DocumentWidth) & (Page_BottomLimit).ToString), XlPasteType.xlPasteAll)
                    'ExRange.Copy(owsheet_Document.Range("A" & (Page_TopLimit + 1).ToString & ":" & GenerateChr(DocumentWidth) & (Page_BottomLimit).ToString))

                    'End Copy From Template

                    Dim cAdr() As DataRow
                    cAdr = DtDataTable.Select("", "Ct")

                    For Each cDr As DataRow In cAdr
                        DataTable_Ct = 0
                        For Each Dc As DataColumn In DtDocumentParameters.Columns
                            If Dc.ColumnName = "Dt_" & cDr.Item("Ct") Then
                                If IsNull(DrPage.Item(Dc), 0) > 0 Then

                                    If pDs.Tables(DataTable_Ct).Rows.Count > 0 Then
                                        DataTable_Ct = cDr.Item("Ct")

                                        Dim Excel_Range() As Int32
                                        Excel_Range = ParseExcelRange(cDr.Item("Location"))

                                        Dim RowStart As Int32 = 0
                                        Dim RowEnd As Int32 = 0

                                        Dim InnerAdr() As DataRow
                                        Dim InnerItemsTotal As Int32 = 0

                                        InnerAdr = DtDocumentParameters.Select("Page < " & DrPage.Item("Page"))
                                        For Each InnerDr As DataRow In InnerAdr
                                            InnerItemsTotal = InnerItemsTotal + IsNull(InnerDr.Item(Dc), 0)
                                        Next

                                        RowStart = InnerItemsTotal
                                        RowEnd = (InnerItemsTotal + DrPage.Item(Dc)) - 1

                                        'Dim RowCt As Int32 = 0
                                        'Dim RowCtEnd As Int32 = 0

                                        For RowCt As Int32 = 0 To RowEnd - RowStart
                                            Dim Adr_Fields() As DataRow
                                            Adr_Fields = DtDataTable_Fields.Select("DtDataTable_Ct = " & DataTable_Ct, "Ct")
                                            For Each Dr_Fields As DataRow In Adr_Fields
                                                owsheet_Document.Cells((Excel_Range(1) + Page_TopLimit) + RowCt, Excel_Range(0) + IsNull(Dr_Fields.Item("Position"), 0)).Value = IsNull(pDs.Tables(DataTable_Ct - 1).Rows(RowCt + RowStart).Item(Dr_Fields.Item("Name")), "")
                                            Next
                                        Next
                                    End If
                                Else
                                    'Clear Field Place Holders
                                    Dim Excel_Range() As Int32
                                    Excel_Range = ParseExcelRange(cDr.Item("Location"))

                                    ExRange = owsheet_Document.Range(GenerateChr(Excel_Range(0)) & (Excel_Range(1) + Page_TopLimit).ToString & ":" & GenerateChr(Excel_Range(2)) & (Excel_Range(1) + Page_TopLimit).ToString)
                                    ExRange.ClearContents()

                                End If
                            End If
                        Next
                    Next

                    ExRange = owsheet_Document.Range("A" & Page_BottomLimit + 1)
                    owsheet_Document.HPageBreaks.Add(ExRange)

                    If IsProtected Then
                        Dim cPassword As String
                        cPassword = Generate_RandomPassword()

                        owsheet_Document.EnableSelection = XlEnableSelection.xlNoSelection
                        owsheet_Document.Protect(cPassword)
                    End If
                Next

                '[-]

                owsheet_Document.Range("A1:A1").Select()
                '--#End Document Generation

                '--#Show the Document Excel Output

                If IsProtected Then
                    Dim cPassword As String
                    cPassword = Generate_RandomPassword()
                    owbook_Document.Protect(cPassword)
                End If

                If SaveFileName = "" Then
                    SaveFileName = "Excel_File"
                End If

                owbook_Document.SaveAs(SaveFileName, FileFormat)

            Catch ex As Exception
                Throw ex
            End Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Shared Sub NativeExcel_CreateExcelDocument_V3(ByVal ExcelTemplateFileName As String, ByVal pParameters() As String, ByVal pParametersValue() As Object, ByVal pDs As DataSet, ByVal pDs_Pivot As DataSet, ByVal pDs_Pivot_Desc As DataSet, ByVal pDs_Pivot_Totals As DataSet, ByVal SaveFileName As String, Optional ByVal IsProtected As Boolean = False, Optional ByVal FileFormat As NativeExcel.XlFileFormat = XlFileFormat.xlNormal)
        Try

            'Dim mExcel As New Excel.Application
            Dim SbExcelData As New System.Text.StringBuilder
            Dim ExRange As NativeExcel.IRange = Nothing

            'Dim owbooks As Excel.Workbooks = mExcel.Workbooks
            Dim owbook_Document As NativeExcel.IWorkbook = Nothing   '--#Workbook for Document
            Dim owbook_Template As NativeExcel.IWorkbook = Nothing     '--#Workbook for Template Data

            Dim owsheet_Document As NativeExcel.IWorksheet = Nothing
            Dim owsheet_Parameters As NativeExcel.IWorksheet = Nothing
            Dim owsheet_Template As NativeExcel.IWorksheet = Nothing

            'Clear all String Fields of all TAB and RETURN chars.
            For Each Inner_Dt As DataTable In pDs.Tables
                For Each Inner_Dr As DataRow In Inner_Dt.Rows
                    For Each Inner_Dc As DataColumn In Inner_Dt.Columns
                        If Inner_Dc.DataType.Name = GetType(System.String).Name Then
                            Inner_Dr.Item(Inner_Dc) = Replace(IsNull(Inner_Dr.Item(Inner_Dc), ""), vbTab, "")
                            Inner_Dr.Item(Inner_Dc) = Replace(IsNull(Inner_Dr.Item(Inner_Dc), ""), vbCrLf, "")
                            Inner_Dr.Item(Inner_Dc) = Replace(IsNull(Inner_Dr.Item(Inner_Dc), ""), vbCr, "")
                            Inner_Dr.Item(Inner_Dc) = Replace(IsNull(Inner_Dr.Item(Inner_Dc), ""), vbLf, "")
                        End If
                    Next
                Next
            Next

            Try

                owbook_Template = NativeExcel.Factory.OpenWorkbook(ExcelTemplateFileName)
                owsheet_Parameters = owbook_Template.Worksheets.Item("Parameters")
                owsheet_Template = owbook_Template.Worksheets.Item("Template")

                owbook_Document = NativeExcel.Factory.OpenWorkbook(ExcelTemplateFileName)

                For Each Ws As NativeExcel.IWorksheet In owbook_Document.Worksheets
                    Ws.Delete()
                Next

                owsheet_Document = owbook_Document.Worksheets.Item("Template")
                owsheet_Document.Name = "Document"

                Dim CtStart As Int32 = 0
                Dim CtEnd As Int32 = 0
                Dim Ct As Int32 = 0

                'Clear the Contents of Document
                'owsheet_Document.Range("A1:IV65536").Clear()
                owsheet_Document.Range("A1:" & GenerateChr(256) & "65536").Clear()

                '[Get Settings]
                Dim DocumentLimit As Int32 = 0
                Dim DocumentWidth As Int32 = 0
                Dim IsRepeatHeader As Boolean = False

                For Ct = 1 To 65536
                    Select Case owsheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                        Case "[#]Settings"
                            CtStart = Ct
                        Case "[#]End_Settings"
                            CtEnd = Ct
                            Exit For
                    End Select
                Next

                For Ct = CtStart To CtEnd
                    Dim ExcelText As String = ""
                    ExcelText = owsheet_Parameters.Range("A" & Ct.ToString).Characters.Text

                    Select Case True
                        Case InStr(ExcelText, "@DocumentLimit")
                            DocumentLimit = CType(Mid(ExcelText, Len("@DocumentLimit") + 1).Trim, Int32)

                        Case InStr(ExcelText, "@DocumentWidth")
                            DocumentWidth = CType(Mid(ExcelText, Len("@DocumentWidth") + 1).Trim, Int32)

                        Case InStr(ExcelText, "@IsRepeatHeader")
                            Try
                                IsRepeatHeader = CType(Mid(ExcelText, Len("@IsRepeatHeader") + 1).Trim, Boolean)
                            Catch
                            End Try

                    End Select
                Next
                '[End Get Settings]

                '[Get Parameters]
                Dim Dt_Parameters As DataTable = NativeExcel_CreateExcelDocument_GetParameters(owsheet_Parameters, pParameters, pParametersValue)
                '[End Get Parameters]

                '[Get Sections]
                Dim Dt_Sections As DataTable = NativeExcel_CreateExcelDocument_GetSections(owsheet_Parameters)
                '[End Get Sections]

                '[Get DataTable]
                Dim Dt_Tables As DataTable = NativeExcel_CreateExcelDocument_GetDataTables(owsheet_Parameters)
                '[End Get DataTable]

                '[Get DataTable_Headers]
                Dim Dt_Tables_Header As DataTable = NativeExcel_CreateExcelDocument_GetDataTable_Headers(owsheet_Parameters)
                '[End Get DataTable_Headers]

                '[Get DataTable_Footers]
                Dim Dt_Tables_Footer As DataTable = NativeExcel_CreateExcelDocument_GetDataTable_Footers(owsheet_Parameters)
                '[End Get DataTable_Footers]

                '[Get Table Fields]
                Dim Dt_Tables_Fields As DataTable = NativeExcel_CreateExcelDocument_GetDataTableFields(Dt_Tables, owsheet_Template)
                '[End Get Table Fields]

                '[Get DataTable_Pivot]
                Dim Dt_Pivot As DataTable = NativeExcel_CreateExcelDocument_GetDataTable_Pivot(owsheet_Parameters)
                '[End Get DataTable_Pivot]

                '[Get DataTable_Pivot_Fields]
                Dim Dt_Pivot_Fields As DataTable = NativeExcel_CreateExcelDocument_GetDataTable_Pivot_Fields(Dt_Pivot, owsheet_Template)
                '[End Get DataTable_Pivot_Fields]

                '[Get DataTable_Pivot_Header]
                Dim Dt_Pivot_Header As DataTable = NativeExcel_CreateExcelDocument_GetDataTable_Pivot_Header(owsheet_Parameters)
                '[End Get DataTable_Pivot_Header]

                '[Get DataTable_Pivot_Header_Fields]
                Dim Dt_Pivot_Header_Fields As DataTable = NativeExcel_CreateExcelDocument_GetDataTable_Pivot_Header_Fields(Dt_Pivot_Header, owsheet_Template)
                '[End Get DataTable_Pivot_Header_Fields]

                '[Get DataTable_Pivot_Totals]
                Dim Dt_Pivot_Totals As DataTable = NativeExcel_CreateExcelDocument_GetDataTable_Pivot_Totals(owsheet_Parameters)
                '[End Get DataTable_Pivot_Totals]

                '[Get DataTable_Pivot_Totals_Fields]
                Dim Dt_Pivot_Totals_Fields As DataTable = NativeExcel_CreateExcelDocument_GetDataTable_Pivot_Totals_Fields(Dt_Pivot_Totals, owsheet_Template)
                '[End Get DataTable_Pivot_Totals_Fields]

                '[Input Data From Parameters]
                Dim Inner0Ct As Int32 = 0
                Dim Inner1Ct As Int32 = 0

                For Inner0Ct = 0 To DocumentLimit - 1
                    For Inner1Ct = 0 To DocumentWidth - 1
                        Dim ExcelText As String = ""
                        Try
                            ExcelText = owsheet_Template.Range(GenerateChr(Inner1Ct + 1) & (Inner0Ct + 1)).Characters.Text
                        Catch
                        End Try

                        Dim Parameter As String = ""
                        Select Case True
                            Case InStr(ExcelText, "[@") > 0
                                Parameter = Mid(ExcelText, InStr(ExcelText, "[@") + Len("[@"), (InStrRev(ExcelText, "]") - Len("]")) - Len("[@"))
                            Case Else
                                Parameter = ""
                        End Select

                        Dim InnerAdr() As DataRow
                        InnerAdr = Dt_Parameters.Select("ParameterName = '" & Parameter & "'")
                        If InnerAdr.Length > 0 Then
                            Dim ParameterValue As Object
                            ParameterValue = InnerAdr(0).Item("ParameterValue")

                            Dim Ar_TypeCode() As Integer
                            Ar_TypeCode = [Enum].GetValues(GetType(TypeCode))

                            For Each Tc As Integer In Ar_TypeCode
                                Dim TcName As String = ""
                                TcName = [Enum].GetName(GetType(TypeCode), Tc)
                                If TcName = InnerAdr(0).Item("ParameterType") Then
                                    Try
                                        ParameterValue = Convert.ChangeType(ParameterValue, Tc)

                                        Select Case CType(Tc, TypeCode)
                                            Case TypeCode.String
                                                ParameterValue = Replace(ParameterValue, vbTab, " ")
                                                ParameterValue = Replace(ParameterValue, vbCrLf, " ")
                                                ParameterValue = Replace(ParameterValue, vbCr, " ")
                                                ParameterValue = Replace(ParameterValue, vbLf, " ")
                                        End Select
                                    Catch
                                    End Try
                                    Exit For
                                End If
                            Next

                            owsheet_Template.Cells(Inner0Ct + 1, Inner1Ct + 1).Value = ParameterValue

                        End If

                    Next
                Next
                '[End Input Data From Parameters]

                '[Input to Document Sheet]
                Dim CtCurrentRow As Int32 = 1

                'Put Header if it exists
                Dim Arr_Dr() As DataRow
                Arr_Dr = Dt_Sections.Select("Type = 'Header'")
                If Arr_Dr.Length > 0 Then
                    Dim Inner_Range() As Int32
                    Dim Location As String = ""
                    Dim Length As Int32

                    Try
                        Inner_Range = ParseExcelRange(Arr_Dr(0).Item("Location"))
                    Catch ex As Exception
                        Throw New Exception(ex.Message & vbCrLf & "Exception Code 1104. Invalid Location in [#]Section: Header.")
                    End Try

                    Length = (Inner_Range(3) - Inner_Range(1)) + 1

                    Location = "A" & Inner_Range(1) & ":" & GenerateChr(DocumentWidth) & Inner_Range(3)

                    owsheet_Template.Range(Location).Copy(owsheet_Document.Range("A" & (CtCurrentRow) & ":" & GenerateChr(DocumentWidth) & ((CtCurrentRow + Length) - 1)))

                    If IsRepeatHeader Then
                        Dim Inner2_Range() As Int32
                        Dim Arr2_Dr() As DataRow = Dt_Sections.Select("Type = 'Repeat'")
                        If Arr2_Dr.Length > 0 Then
                            Inner2_Range = ParseExcelRange(Arr2_Dr(0)("Location"))
                        Else
                            Inner2_Range = Inner_Range
                        End If

                        'Set Rows to Repeat in Top (Page Setup)
                        owsheet_Document.PageSetup.PrintTitleRows = "$" & Inner2_Range(1) & ":" & "$" & Inner2_Range(3)
                    End If

                    CtCurrentRow = CtCurrentRow + Length

                End If

                'Put Tables
                Dim Arr_Dr_Dt_Tables() As DataRow
                Arr_Dr_Dt_Tables = Dt_Tables.Select("ISNULL(IsSubTable,0) = 0")
                For Each Dr_Dt_Tables As DataRow In Arr_Dr_Dt_Tables

                    Dim Inner_Range() As Int32
                    Dim Location As String = Dr_Dt_Tables.Item("Location")
                    Dim CtTable As Int32 = (Dr_Dt_Tables.Item("Ct") - 1)
                    'Dim CtItems As Int32 = pDs.Tables(CtTable).Rows.Count
                    Dim CtItems As Int32 = NativeExcel_CreateExcelDocument_V3_CountItem(Dr_Dt_Tables, Dt_Tables, pDs, Nothing)

                    'Check the position of this table to get CtCurrentRow to be used
                    Inner_Range = ParseExcelRange(Location)

                    Dim Inner_CtItems As Int32 = 0
                    Dim Inner_Arr_Dr() As DataRow = Dt_Tables.Select("ISNULL(IsSubTable,0) = 0 And Ct <> " & (CtTable + 1))
                    For Each Inner_Dr As DataRow In Inner_Arr_Dr
                        Dim Inner2_Range() As Int32 = ParseExcelRange(IsNull(Inner_Dr("Location"), ""))
                        If (((Inner2_Range(0) <= Inner_Range(0)) Or (Inner2_Range(0) >= Inner_Range(2))) Or ((Inner2_Range(2) <= Inner_Range(0)) Or (Inner2_Range(2) >= Inner_Range(2)))) And (Inner_Range(1) >= Inner2_Range(3)) Then
                            Inner_CtItems = Inner_CtItems + NativeExcel_CreateExcelDocument_V3_CountItem(Inner_Dr, Dt_Tables, pDs, Nothing)
                            Dim Inner2_Arr_Dr() As DataRow
                            Inner2_Arr_Dr = Dt_Tables_Header.Select("Name = '" & IsNull(Inner_Dr("Name"), "") & "'")
                            If Inner2_Arr_Dr.Length > 0 Then
                                Inner_CtItems = Inner_CtItems + (ParseExcelRange_GetHeight(IsNull(Inner2_Arr_Dr(0)("Location"), "")) + 1)
                            End If
                            Inner2_Arr_Dr = Dt_Tables_Footer.Select("Name = '" & IsNull(Inner_Dr("Name"), "") & "'")
                            If Inner2_Arr_Dr.Length > 0 Then
                                Inner_CtItems = Inner_CtItems + (ParseExcelRange_GetHeight(IsNull(Inner2_Arr_Dr(0)("Location"), "")) + 1)
                            End If
                        End If
                    Next

                    Dim Inner_CtCurrentRow As Int32 = CtCurrentRow + Inner_CtItems

                    'Table Header
                    Arr_Dr = Dt_Tables_Header.Select("Name = '" & Dr_Dt_Tables.Item("Name") & "'")
                    If Arr_Dr.Length > 0 Then
                        Inner_Range = ParseExcelRange(Arr_Dr(0).Item("Location"))

                        Dim Length As Int32
                        Length = ParseExcelRange_GetHeight(IsNull(Arr_Dr(0).Item("Location"), ""))

                        Dim Inner_Source_Location As String = Arr_Dr(0).Item("Location")
                        Dim Inner_Target_Location As String = GenerateChr(Inner_Range(0)) & (Inner_CtCurrentRow) & ":" & GenerateChr(Inner_Range(2)) & ((Inner_CtCurrentRow) + Length)
                        owsheet_Template.Range(Inner_Source_Location).Copy(owsheet_Document.Range(Inner_Target_Location))

                        Inner_CtCurrentRow = (Inner_CtCurrentRow + Length) + 1
                        Inner_CtItems = (Inner_CtItems + Length) + 1
                    End If

                    'Table Items
                    If CtItems > 0 Then
                        NativeExcel_CreateExcelDocument_V3_SubTable(pDs, Dt_Tables, Dt_Tables_Fields, owsheet_Template, owsheet_Document, Inner_CtCurrentRow, New DataRow() {Dr_Dt_Tables}, Nothing)
                    End If

                    'Table Footer
                    Dim Inner_Tables_Footer_Length As Int32 = 0
                    Arr_Dr = Dt_Tables_Footer.Select("Name = '" & Dr_Dt_Tables.Item("Name") & "'")
                    If Arr_Dr.Length > 0 Then
                        Inner_Range = ParseExcelRange(Arr_Dr(0).Item("Location"))

                        Dim Length As Int32
                        Length = ParseExcelRange_GetHeight(IsNull(Arr_Dr(0).Item("Location"), ""))

                        Dim Inner_Source_Location As String = Arr_Dr(0).Item("Location")
                        Dim Inner_Target_Location As String = GenerateChr(Inner_Range(0)) & (Inner_CtCurrentRow) & ":" & GenerateChr(Inner_Range(2)) & ((Inner_CtCurrentRow) + Length)
                        owsheet_Template.Range(Inner_Source_Location).Copy(owsheet_Document.Range(Inner_Target_Location))

                        Inner_CtItems = (Inner_CtItems + Length) + 1
                        Inner_Tables_Footer_Length = Inner_Tables_Footer_Length + Length + 1
                    End If

                    '[Pivot Tables]

                    Dim Inner_OffsetColumn As Int32 = 0
                    Dim Inner_OffsetRow As Int32 = CtItems + Inner_Tables_Footer_Length

                    'Reuse Inner_CtCurrentRow, reassign with CtCurrentRow for the current iteration
                    Inner_CtCurrentRow = CtCurrentRow

                    Dim Arr_Dr_Dt_Pivot() As DataRow
                    Arr_Dr_Dt_Pivot = Dt_Pivot.Select("ParentTableName = '" & Dr_Dt_Tables("Name") & "'")
                    For Each Dr_Dt_Pivot As DataRow In Arr_Dr_Dt_Pivot

                        'Pivot Header
                        Dim Target_Location As String = ""

                        Dim Arr_Dr_Dt_Pivot_Header() As DataRow = Dt_Pivot_Header.Select("Name = '" & Dr_Dt_Pivot("Name") & "'")
                        Dim Dr_Dt_Pivot_Header As DataRow
                        If Arr_Dr_Dt_Pivot_Header.Length > 0 Then
                            Dr_Dt_Pivot_Header = Arr_Dr_Dt_Pivot_Header(0)
                        Else
                            Throw New Exception("Pivot Tables must have a Header")
                        End If

                        Dim Inner_PivotHeader_Location As String = Dr_Dt_Pivot_Header("Location")
                        Dim Inner_PivotHeader_Range() As Int32 = ParseExcelRange(Inner_PivotHeader_Location)
                        Dim Inner_PivotHeaderCount As Int32 = pDs_Pivot_Desc.Tables(CInt(Dr_Dt_Pivot_Header("Ct") - 1)).Rows.Count
                        Arr_Dr = Dt_Pivot_Header_Fields.Select("DtDataTable_Pivot_Header_Ct = " & Dr_Dt_Pivot_Header("Ct"))
                        Dim Inner_PivotHeaderFieldCount As Int32 = Arr_Dr.Length

                        Target_Location = GenerateChr(Inner_PivotHeader_Range(0) + Inner_OffsetColumn) & CtCurrentRow & ":" & GenerateChr((Inner_PivotHeader_Range(0) + ((((Inner_PivotHeader_Range(2) - Inner_PivotHeader_Range(0)) + 1) * (Inner_PivotHeaderCount - 1)) - 1) + Inner_OffsetColumn)) & (CtCurrentRow + Inner_OffsetRow)
                        owsheet_Document.Range(Target_Location).Insert(XlInsertShiftDirection.xlShiftToRight)

                        Dim Target_Location_Range() As Int32 = ParseExcelRange(Target_Location)
                        Target_Location = GenerateChr(Target_Location_Range(0)) & Target_Location_Range(1) & ":" & GenerateChr(Target_Location_Range(2)) & Target_Location_Range(1)

                        owsheet_Template.Range(Inner_PivotHeader_Location).Copy(owsheet_Document.Range(Target_Location), XlPasteType.xlPasteFormats)

                        Dim Inner2_Ct As Int32 = 0
                        For Each Inner_Dr As DataRow In pDs_Pivot_Desc.Tables(CInt(Dr_Dt_Pivot_Header("Ct")) - 1).Rows
                            Dim Inner2_OffsetColumn As Int32 = (Inner2_Ct * ((Inner_PivotHeader_Range(2) - Inner_PivotHeader_Range(0)) + 1))
                            For Each Inner2_Dr As DataRow In Dt_Pivot_Header_Fields.Select("DtDataTable_Pivot_Header_Ct = " & CInt(Dr_Dt_Pivot_Header("Ct")))
                                Dim Inner_Template_Position As Int32 = Inner2_Dr("Position")
                                Dim Inner_Document_Location As String = GenerateChr((Inner_PivotHeader_Range(0) + Inner_OffsetColumn) + Inner2_Dr("Position") + Inner2_OffsetColumn) & CtCurrentRow
                                owsheet_Document.Range(Inner_Document_Location).Characters.Text = Inner_Dr(Inner2_Dr("Name"))
                            Next
                            Inner2_Ct = Inner2_Ct + 1
                        Next

                        Inner_CtCurrentRow = Inner_CtCurrentRow + ParseExcelRange_GetHeight(Dr_Dt_Tables("Location")) + 1

                        '[-]

                        'Pivot Items

                        Dim Inner_Pivot_Location As String = Dr_Dt_Pivot("Location")
                        Dim Inner_Pivot_Location_Range() As Int32 = ParseExcelRange(Inner_Pivot_Location)

                        Dim Inner_StartRow As Int32 = Inner_CtCurrentRow
                        Dim Inner_EndRow As Int32 = Inner_StartRow + (CtItems - 1)

                        Dim Inner_Source_Location As String = ""
                        Dim Inner_Target_Location As String = ""

                        'Table Formats And Borders
                        ' - Top
                        Inner_Source_Location = GenerateChr(Inner_Pivot_Location_Range(0)) & (Inner_Pivot_Location_Range(1)) & ":" & GenerateChr(Inner_Pivot_Location_Range(2)) & (Inner_Pivot_Location_Range(1))
                        Inner_Target_Location = GenerateChr(Inner_Pivot_Location_Range(0)) & Inner_StartRow & ":" & GenerateChr((Inner_PivotHeader_Range(0) + ((((Inner_PivotHeader_Range(2) - Inner_PivotHeader_Range(0)) + 1) * (Inner_PivotHeaderCount)) - 1))) & (Inner_StartRow)
                        owsheet_Template.Range(Inner_Source_Location).Copy(owsheet_Document.Range(Inner_Target_Location), XlPasteType.xlPasteFormats)

                        ' - Middle
                        If CtItems > 2 Then
                            Inner_Source_Location = GenerateChr(Inner_Pivot_Location_Range(0)) & (Inner_Pivot_Location_Range(1) + 1) & ":" & GenerateChr(Inner_Pivot_Location_Range(2)) & (Inner_Pivot_Location_Range(1) + 1)
                            Inner_Target_Location = GenerateChr(Inner_Pivot_Location_Range(0)) & (Inner_StartRow + 1) & ":" & GenerateChr((Inner_PivotHeader_Range(0) + ((((Inner_PivotHeader_Range(2) - Inner_PivotHeader_Range(0)) + 1) * (Inner_PivotHeaderCount)) - 1))) & (Inner_EndRow - 1)
                            owsheet_Template.Range(Inner_Source_Location).Copy(owsheet_Document.Range(Inner_Target_Location), XlPasteType.xlPasteFormats)
                        End If

                        ' - Bottom
                        Inner_Source_Location = GenerateChr(Inner_Pivot_Location_Range(0)) & (Inner_Pivot_Location_Range(1) + 2) & ":" & GenerateChr(Inner_Pivot_Location_Range(2)) & (Inner_Pivot_Location_Range(1) + 2)
                        Inner_Target_Location = GenerateChr(Inner_Pivot_Location_Range(0)) & (Inner_EndRow) & ":" & GenerateChr((Inner_PivotHeader_Range(0) + ((((Inner_PivotHeader_Range(2) - Inner_PivotHeader_Range(0)) + 1) * (Inner_PivotHeaderCount)) - 1))) & (Inner_EndRow)
                        owsheet_Template.Range(Inner_Source_Location).Copy(owsheet_Document.Range(Inner_Target_Location), XlPasteType.xlPasteFormats)

                        If CtItems = 1 Then
                            Inner_Source_Location = GenerateChr(Inner_Pivot_Location_Range(0)) & (Inner_Pivot_Location_Range(1) + 4) & ":" & GenerateChr(Inner_Pivot_Location_Range(2)) & (Inner_Pivot_Location_Range(1) + 4)
                            Inner_Target_Location = GenerateChr(Inner_Pivot_Location_Range(0)) & (Inner_EndRow) & ":" & GenerateChr((Inner_PivotHeader_Range(0) + ((((Inner_PivotHeader_Range(2) - Inner_PivotHeader_Range(0)) + 1) * (Inner_PivotHeaderCount)) - 1))) & (Inner_EndRow)
                            owsheet_Template.Range(Inner_Source_Location).Copy(owsheet_Document.Range(Inner_Target_Location), XlPasteType.xlPasteFormats)
                        End If

                        'Input Pivot Items
                        ' - Prepare DataTable Inner_Dt_Pivot_Items to be data container to input in excel
                        Dim Inner_Dt_Pivot_Items As New DataTable
                        For Each Inner2_Dr_Dt_Pivot_Fields As DataRow In Dt_Pivot_Fields.Select("DtDataTable_Pivot_Ct = " & Dr_Dt_Pivot("Ct") & "")
                            For Each Inner2_Dr_Pivot_Desc As DataRow In pDs_Pivot_Desc.Tables(CInt(Dr_Dt_Pivot_Header("Ct") - 1)).Rows
                                Dim Inner_DataType As System.Type = GetType(String)
                                For Each Inner_Dc As DataColumn In pDs_Pivot.Tables(CInt(Dr_Dt_Pivot("Ct") - 1)).Columns
                                    If Inner_Dc.ColumnName = Inner2_Dr_Dt_Pivot_Fields("Name") Then
                                        Inner_DataType = Inner_Dc.DataType
                                        Continue For
                                    End If
                                Next
                                Inner_Dt_Pivot_Items.Columns.Add(Inner2_Dr_Dt_Pivot_Fields("Name") & "_" & Inner2_Dr_Pivot_Desc("ID"), Inner_DataType)
                            Next
                        Next

                        Dim Inner2_Arr_Dr() As DataRow = pDs.Tables(CtTable).Select
                        For Each Inner2_Dr As DataRow In Inner2_Arr_Dr
                            Dim Inner_Nr As DataRow = Inner_Dt_Pivot_Items.NewRow
                            Inner_Dt_Pivot_Items.Rows.Add(Inner_Nr)

                            For Each Inner2_Dr_Dt_Pivot_Fields As DataRow In Dt_Pivot_Fields.Select("DtDataTable_Pivot_Ct = " & Dr_Dt_Pivot("Ct") & "")
                                For Each Inner2_Dr_Pivot_Desc As DataRow In pDs_Pivot_Desc.Tables(CInt(Dr_Dt_Pivot_Header("Ct") - 1)).Rows
                                    Dim Inner3_Arr_Dr() As DataRow = pDs_Pivot.Tables(CInt(Dr_Dt_Pivot("Ct") - 1)).Select("ID = " & Inner2_Dr_Pivot_Desc("ID") & " And " & Dr_Dt_Pivot("SourceKey") & " = " & Inner2_Dr(Dr_Dt_Pivot("TargetKey")))
                                    If Inner3_Arr_Dr.Length > 0 Then
                                        Inner_Nr(Inner2_Dr_Dt_Pivot_Fields("Name") & "_" & Inner2_Dr_Pivot_Desc("ID")) = Inner3_Arr_Dr(0)(Inner2_Dr_Dt_Pivot_Fields("Name"))
                                    End If
                                Next
                            Next

                            Dim Inner2_CtItems As Int32 = 0
                            Inner2_CtItems = NativeExcel_CreateExcelDocument_V3_CountItem(Dr_Dt_Tables, Dt_Tables, pDs, Inner2_Dr, False)
                            For Inner3_Ct As Int32 = 0 To Inner2_CtItems - 1
                                Inner_Nr = Inner_Dt_Pivot_Items.NewRow
                                Inner_Dt_Pivot_Items.Rows.Add(Inner_Nr)
                            Next
                        Next

                        ' - Prepare Pivot Fields Definition
                        Dim Inner5_Ct As Int32 = 0
                        Dim Inner5_Pivot_Length As Int32 = (Inner_Pivot_Location_Range(2) - Inner_Pivot_Location_Range(0)) + 1

                        Dim Inner_Arr_St_Pivot_Fields() As String
                        ReDim Inner_Arr_St_Pivot_Fields((Inner_PivotHeaderCount * Inner5_Pivot_Length) - 1)
                        For Each Inner_St As String In Inner_Arr_St_Pivot_Fields
                            Inner_St = ""
                        Next

                        For Each Inner2_Dr_Pivot_Desc As DataRow In pDs_Pivot_Desc.Tables(CInt(Dr_Dt_Pivot_Header("Ct") - 1)).Rows
                            For Inner4_Ct As Int32 = 0 To Inner5_Pivot_Length - 1
                                Dim Inner3_Arr_Dr() As DataRow
                                Inner3_Arr_Dr = Dt_Pivot_Fields.Select("DtDataTable_Pivot_Ct = " & Dr_Dt_Pivot("Ct") & " And Position = " & Inner4_Ct)
                                If Inner3_Arr_Dr.Length > 0 Then
                                    Inner_Arr_St_Pivot_Fields(Inner4_Ct + (Inner5_Ct * Inner5_Pivot_Length)) = Inner3_Arr_Dr(0).Item("Name") & "_" & Inner2_Dr_Pivot_Desc("ID")
                                Else
                                    Inner_Arr_St_Pivot_Fields(Inner4_Ct + (Inner5_Ct * Inner5_Pivot_Length)) = ""
                                End If
                            Next
                            Inner5_Ct = Inner5_Ct + 1
                        Next

                        ' - Insert Data to Document Worksheet

                        Dim Inner_ExRange As IRange
                        Inner_ExRange = owsheet_Document.Range(GenerateChr(Inner_Pivot_Location_Range(0)) & Inner_CtCurrentRow & ":" & GenerateChr((Inner_Pivot_Location_Range(0) + ((((Inner_Pivot_Location_Range(2) - Inner_Pivot_Location_Range(0)) + 1) * (Inner_PivotHeaderCount)) - 1))) & (Inner_CtCurrentRow + (CtItems - 1)))
                        Inner_ExRange.Value = ConvertDataTo2DimArray(Inner_Dt_Pivot_Items, "", Inner_Arr_St_Pivot_Fields)

                        '[-]

                        Inner_CtCurrentRow = (Inner_CtCurrentRow + (CtItems - 1)) + 1

                        'Pivot Totals

                        Dim Arr_Dr_Dt_Pivot_Totals() As DataRow = Dt_Pivot_Totals.Select("Name = '" & Dr_Dt_Pivot("Name") & "'")
                        For Each Inner_Dr_Dt_Pivot_Totals As DataRow In Arr_Dr_Dt_Pivot_Totals

                            Dim Inner6_Pivot_Totals_Location As String = Inner_Dr_Dt_Pivot_Totals("Location")
                            Dim Inner6_Pivot_Totals_Location_Range() As Int32 = ParseExcelRange(Inner6_Pivot_Totals_Location)

                            Dim Inner6_Pivot_Totals_Count As Int32 = pDs_Pivot_Totals.Tables(CInt(Inner_Dr_Dt_Pivot_Totals("Ct") - 1)).Rows.Count

                            Dim Inner6_Ct As Int32 = 0
                            Dim Inner6_Pivot_Length As Int32 = (Inner6_Pivot_Totals_Location_Range(2) - Inner6_Pivot_Totals_Location_Range(0)) + 1

                            Dim Inner_Arr_St_Pivot_Totals_Fields() As String
                            ReDim Inner_Arr_St_Pivot_Totals_Fields((Inner6_Pivot_Totals_Count * Inner6_Pivot_Length) - 1)
                            For Each Inner_St As String In Inner_Arr_St_Pivot_Totals_Fields
                                Inner_St = ""
                            Next

                            For Each Inner6_Dr_Pivot_Totals As DataRow In pDs_Pivot_Totals.Tables(CInt(Inner_Dr_Dt_Pivot_Totals("Ct") - 1)).Rows
                                For Inner7_Ct As Int32 = 0 To Inner6_Pivot_Length - 1
                                    Dim Inner7_Arr_Dr() As DataRow
                                    Inner7_Arr_Dr = Dt_Pivot_Totals_Fields.Select("DtDataTable_Pivot_Totals_Ct = " & Inner_Dr_Dt_Pivot_Totals("Ct") & " And Position = " & Inner7_Ct)
                                    If Inner7_Arr_Dr.Length > 0 Then
                                        Inner_Arr_St_Pivot_Totals_Fields(Inner7_Ct + (Inner6_Ct * Inner6_Pivot_Length)) = Inner7_Arr_Dr(0).Item("Name") & "_" & Inner6_Dr_Pivot_Totals("ID")
                                    Else
                                        Inner_Arr_St_Pivot_Totals_Fields(Inner7_Ct + (Inner6_Ct * Inner6_Pivot_Length)) = ""
                                    End If
                                Next
                                Inner6_Ct = Inner6_Ct + 1
                            Next

                            Dim Inner_Dt_Pivot_Totals_Items As New DataTable
                            For Each Inner6_Dr_Dt_Pivot_Totals_Fields As DataRow In Dt_Pivot_Totals_Fields.Select("DtDataTable_Pivot_Totals_Ct = " & Inner_Dr_Dt_Pivot_Totals("Ct") & "")
                                For Each Inner6_Dr_Pivot_Totals As DataRow In pDs_Pivot_Totals.Tables(CInt(Inner_Dr_Dt_Pivot_Totals("Ct") - 1)).Rows
                                    Dim Inner_DataType As System.Type = GetType(String)
                                    For Each Inner_Dc As DataColumn In pDs_Pivot_Totals.Tables(CInt(Inner_Dr_Dt_Pivot_Totals("Ct") - 1)).Columns
                                        If Inner_Dc.ColumnName = Inner6_Dr_Dt_Pivot_Totals_Fields("Name") Then
                                            Inner_DataType = Inner_Dc.DataType
                                            Continue For
                                        End If
                                    Next
                                    Inner_Dt_Pivot_Totals_Items.Columns.Add(Inner6_Dr_Dt_Pivot_Totals_Fields("Name") & "_" & Inner6_Dr_Pivot_Totals("ID"), Inner_DataType)
                                Next
                            Next

                            Dim Inner_Nr As DataRow = Inner_Dt_Pivot_Totals_Items.NewRow
                            Inner_Dt_Pivot_Totals_Items.Rows.Add(Inner_Nr)

                            For Each Inner6_Dr_Dt_Pivot_Totals_Fields As DataRow In Dt_Pivot_Totals_Fields.Select("DtDataTable_Pivot_Totals_Ct = " & Dr_Dt_Pivot("Ct") & "")
                                For Each Inner6_Dr_Pivot_Totals As DataRow In pDs_Pivot_Totals.Tables(CInt(Inner_Dr_Dt_Pivot_Totals("Ct") - 1)).Rows
                                    Dim Inner6_Arr_Dr() As DataRow = pDs_Pivot_Totals.Tables(CInt(Inner_Dr_Dt_Pivot_Totals("Ct") - 1)).Select("ID = " & Inner6_Dr_Pivot_Totals("ID"))
                                    If Inner6_Arr_Dr.Length > 0 Then
                                        Try
                                            Inner_Nr(Inner6_Dr_Dt_Pivot_Totals_Fields("Name") & "_" & Inner6_Dr_Pivot_Totals("ID")) = Inner6_Arr_Dr(0)(Inner6_Dr_Dt_Pivot_Totals_Fields("Name"))
                                        Catch
                                        End Try
                                    End If
                                Next
                            Next

                            'Table Formats And Borders
                            Inner_Source_Location = GenerateChr(Inner6_Pivot_Totals_Location_Range(0)) & (Inner6_Pivot_Totals_Location_Range(1)) & ":" & GenerateChr(Inner6_Pivot_Totals_Location_Range(2)) & (Inner6_Pivot_Totals_Location_Range(1))
                            Inner_Target_Location = GenerateChr(Inner6_Pivot_Totals_Location_Range(0)) & Inner_CtCurrentRow & ":" & GenerateChr((Inner6_Pivot_Totals_Location_Range(0) + ((((Inner6_Pivot_Totals_Location_Range(2) - Inner6_Pivot_Totals_Location_Range(0)) + 1) * (Inner6_Pivot_Totals_Count)))) - 1) & (Inner_CtCurrentRow)
                            owsheet_Template.Range(Inner_Source_Location).Copy(owsheet_Document.Range(Inner_Target_Location), XlPasteType.xlPasteFormats)

                            Dim Inner6_ExRange As IRange
                            Inner6_ExRange = owsheet_Document.Range(GenerateChr(Inner6_Pivot_Totals_Location_Range(0)) & Inner_CtCurrentRow & ":" & GenerateChr((Inner6_Pivot_Totals_Location_Range(0) + ((((Inner6_Pivot_Totals_Location_Range(2) - Inner6_Pivot_Totals_Location_Range(0)) + 1) * (Inner6_Pivot_Totals_Count)))) - 1) & (Inner_CtCurrentRow))
                            Inner6_ExRange.Value = ConvertDataTo2DimArray(Inner_Nr, "", Inner_Arr_St_Pivot_Totals_Fields)
                        Next
                    Next

                    '[End Pivot Tables]

                    '[-]

                    Dr_Dt_Tables.Item("Items") = Inner_CtItems + CtItems

                Next

                Arr_Dr = Dt_Tables.Select("ISNULL(IsSubTable,0) = 0", "Items Desc")
                If Arr_Dr.Length > 0 Then
                    CtCurrentRow = CtCurrentRow + (Arr_Dr(0).Item("Items"))
                End If

                'Put Footer
                Arr_Dr = Dt_Sections.Select("Type = 'Footer'")
                If Arr_Dr.Length > 0 Then
                    Dim InnerRange() As Int32
                    Dim Location As String = ""
                    Dim Length As Int32

                    Try
                        InnerRange = ParseExcelRange(Arr_Dr(0).Item("Location"))
                    Catch ex As Exception
                        Throw New Exception(ex.Message & vbCrLf & "Exception Code 1106. Invalid Location in [#]Section: Footer.")
                    End Try

                    Length = (InnerRange(3) - InnerRange(1))

                    Location = "A" & InnerRange(1) & ":" & GenerateChr(DocumentWidth) & InnerRange(3)
                    owsheet_Template.Range(Location).Copy(owsheet_Document.Range("A" & (CtCurrentRow) & ":" & GenerateChr(DocumentWidth) & (CtCurrentRow + Length)))

                    CtCurrentRow = (CtCurrentRow + Length) + 1

                End If

                'owsheet_Document.Cells("A1:IV1").Insert(XlInsertShiftDirection.xlShiftDown)
                'owsheet_Document.Cells("A1").RowHeight = 0

                'owsheet_Document.Cells("A1:A65536").Insert(XlInsertShiftDirection.xlShiftToRight)
                'owsheet_Document.Cells("A1").ColumnWidth = 0

                owsheet_Document.Activate()
                owsheet_Document.Range("A2").Select()

                '[End Input to Document Sheet]

                '[Show the Document Excel Output]
                If IsProtected Then
                    Dim cPassword As String
                    cPassword = Generate_RandomPassword()

                    owsheet_Document.EnableSelection = XlEnableSelection.xlNoSelection
                    owsheet_Document.Protect(cPassword)
                    owbook_Document.Protect(cPassword)
                End If

                If SaveFileName = "" Then
                    SaveFileName = "Excel_File"
                End If

                owbook_Document.SaveAs(SaveFileName, FileFormat)

            Catch ex As Exception
                Throw ex
            End Try
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    '[-]

    Shared Sub NativeExcel_CreateExcelDocument_V3_SubTable(ByVal pDs As DataSet, ByVal pDt_Tables As DataTable, ByVal pDt_Fields As DataTable, ByRef pSheet_Template As NativeExcel.IWorksheet, ByRef pSheet_Document As NativeExcel.IWorksheet, ByRef pCtCurrentRow As Int32, ByVal pAdr_Tables() As DataRow, ByVal SourceKey_Dr As DataRow)

        For Each Dr As DataRow In pAdr_Tables
            Dim Inner_Range() As Int32
            Dim CtTable As Int32 = (Dr("Ct") - 1)
            Dim CtItems As Int32 = NativeExcel_CreateExcelDocument_V3_CountItem(Dr, pDt_Tables, pDs, SourceKey_Dr)

            Try
                Inner_Range = ParseExcelRange(Dr("Location"))
            Catch ex As Exception
                Throw New Exception(ex.Message & vbCrLf & "Exception Code 1105. Invalid Location in [#]DataTable: " & Dr("Name") & ".")
            End Try

            If CtItems > 0 Then

                'Copy Formatting from Template

                Dim Inner_StartRow As Int32 = pCtCurrentRow
                Dim Inner_EndRow As Int32 = Inner_StartRow + (CtItems - 1)

                Dim Source_Location As String = ""
                Dim Target_Location As String = ""

                'Table Formats And Borders
                ' - Top
                Source_Location = GenerateChr(Inner_Range(0)) & (Inner_Range(1)) & ":" & GenerateChr(Inner_Range(2)) & (Inner_Range(1))
                Target_Location = GenerateChr(Inner_Range(0)) & Inner_StartRow & ":" & GenerateChr(Inner_Range(2)) & Inner_StartRow
                pSheet_Template.Range(Source_Location).Copy(pSheet_Document.Range(Target_Location), XlPasteType.xlPasteFormats)

                ' - Middle
                If CtItems > 2 Then
                    Source_Location = GenerateChr(Inner_Range(0)) & (Inner_Range(1) + 1) & ":" & GenerateChr(Inner_Range(2)) & (Inner_Range(1) + 1)
                    Target_Location = GenerateChr(Inner_Range(0)) & (Inner_StartRow + 1) & ":" & GenerateChr(Inner_Range(2)) & (Inner_EndRow - 1)
                    pSheet_Template.Range(Source_Location).Copy(pSheet_Document.Range(Target_Location), XlPasteType.xlPasteFormats)
                End If

                ' - Bottom
                Source_Location = GenerateChr(Inner_Range(0)) & (Inner_Range(1) + 2) & ":" & GenerateChr(Inner_Range(2)) & (Inner_Range(1) + 2)
                Target_Location = GenerateChr(Inner_Range(0)) & Inner_EndRow & ":" & GenerateChr(Inner_Range(2)) & Inner_EndRow
                pSheet_Template.Range(Source_Location).Copy(pSheet_Document.Range(Target_Location), XlPasteType.xlPasteFormats)

                If CtItems = 1 Then
                    Source_Location = GenerateChr(Inner_Range(0)) & (Inner_Range(1) + 4) & ":" & GenerateChr(Inner_Range(2)) & (Inner_Range(1) + 4)
                    Target_Location = GenerateChr(Inner_Range(0)) & Inner_EndRow & ":" & GenerateChr(Inner_Range(2)) & Inner_EndRow
                    pSheet_Template.Range(Source_Location).Copy(pSheet_Document.Range(Target_Location), XlPasteType.xlPasteFormats)
                End If

                ''Clear all String Fields of all TAB and RETURN chars.
                'For Each Inner_Dr As DataRow In pDs.Tables(CtTable).Rows
                '    For Each Inner_Dc As DataColumn In pDs.Tables(CtTable).Columns
                '        If Inner_Dc.DataType.Name = GetType(System.String).Name Then
                '            Inner_Dr.Item(Inner_Dc) = Replace(IsNull(Inner_Dr.Item(Inner_Dc), ""), vbTab, "")
                '            Inner_Dr.Item(Inner_Dc) = Replace(IsNull(Inner_Dr.Item(Inner_Dc), ""), vbCrLf, "")
                '            Inner_Dr.Item(Inner_Dc) = Replace(IsNull(Inner_Dr.Item(Inner_Dc), ""), vbCr, "")
                '            Inner_Dr.Item(Inner_Dc) = Replace(IsNull(Inner_Dr.Item(Inner_Dc), ""), vbLf, "")
                '        End If
                '    Next
                'Next

                'Input Table Data

                Dim Arr_StFields() As String
                ReDim Arr_StFields(Inner_Range(2) - Inner_Range(0))

                Dim Inner_Ct As Int32 = 0
                For Inner_Ct = 0 To (Arr_StFields.Length - 1)
                    Dim Inner_Arr_Dr() As DataRow
                    Inner_Arr_Dr = pDt_Fields.Select("DtDataTable_Ct = " & Dr("Ct") & " And Position = " & Inner_Ct)
                    If Inner_Arr_Dr.Length > 0 Then
                        Arr_StFields(Inner_Ct) = Inner_Arr_Dr(0).Item("Name")
                    Else
                        Arr_StFields(Inner_Ct) = ""
                    End If
                Next

                Dim SourceKey As String = IsNull(Dr("SourceKey"), "")
                Dim TargetKey As String = IsNull(Dr("TargetKey"), "")

                Dim Condition As String = ""
                Dim SourceKey_ID As String = "0"

                If Not SourceKey = "" Then
                    Try
                        SourceKey_ID = IsNull(SourceKey_Dr(SourceKey), 0)
                        Condition = TargetKey & " = " & "'" & SourceKey_ID & "'"
                    Catch
                    End Try
                End If

                Dim Arr_Dr_Data As DataRow() = pDs.Tables(CtTable).Select(Condition)

                For Each Inner_Dr As DataRow In Arr_Dr_Data
                    Dim ExRange As IRange
                    ExRange = pSheet_Document.Range(GenerateChr(Inner_Range(0)) & pCtCurrentRow & ":" & GenerateChr(Inner_Range(2)) & pCtCurrentRow)
                    ExRange.Value = ConvertDataTo2DimArray(Inner_Dr, "", Arr_StFields)
                    pCtCurrentRow = pCtCurrentRow + 1

                    Dim Inner_TableName As String = IsNull(Dr("Name"), "")
                    Dim Inner_Arr_Dr_Group() As DataRow = pDt_Tables.Select("ISNULL(IsSubTable,0) = 1 And GroupName = '" & Inner_TableName & "'")
                    If Inner_Arr_Dr_Group.Length > 0 Then
                        NativeExcel_CreateExcelDocument_V3_SubTable(pDs, pDt_Tables, pDt_Fields, pSheet_Template, pSheet_Document, pCtCurrentRow, Inner_Arr_Dr_Group, Inner_Dr)
                    End If
                Next

            End If
        Next

    End Sub

    Shared Function NativeExcel_CreateExcelDocument_V3_CountItem(ByVal Dr_Table As DataRow, ByRef Dt_Tables As DataTable, ByRef Ds As DataSet, ByVal SourceKey_Dr As DataRow, Optional ByVal IsIncludeCurrent As Boolean = True) As Int32

        Dim SourceKey As String = IsNull(Dr_Table("SourceKey"), "")
        Dim TargetKey As String = IsNull(Dr_Table("TargetKey"), "")

        Dim Condition As String = ""
        Dim SourceKey_ID As String = 0

        If Not SourceKey = "" Then
            Try
                SourceKey_ID = IsNull(SourceKey_Dr(SourceKey), 0)
                Condition = TargetKey & " = " & "'" & SourceKey_ID & "'"
            Catch
            End Try
        End If

        Dim ReturnValue As Int32 = 0
        Dim CtTable As Int32 = (Dr_Table("Ct") - 1)
        Dim Adr_Data() As DataRow = Ds.Tables.Item(CtTable).Select(Condition)

        If IsIncludeCurrent Then
            ReturnValue = ReturnValue + Adr_Data.Length
        End If

        Dim Adr() As DataRow = Dt_Tables.Select("GroupName = '" & IsNull(Dr_Table("Name"), "") & "'")
        If Adr.Length > 0 Then
            For Each Dr As DataRow In Adr
                For Each InnerDr As DataRow In Adr_Data
                    Dim Inner_ReturnValue As Int32 = 0
                    Inner_ReturnValue = NativeExcel_CreateExcelDocument_V3_CountItem(Dr, Dt_Tables, Ds, InnerDr, True)
                    ReturnValue = ReturnValue + Inner_ReturnValue
                Next
            Next
        End If

        Return ReturnValue

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetSections(ByRef pSheet_Parameters As NativeExcel.IWorksheet) As DataTable
        Dim Dt_ReturnValue As New DataTable
        Dt_ReturnValue.Columns.Add("Ct", GetType(System.Int32))
        Dt_ReturnValue.Columns.Add("Type", GetType(System.String))
        Dt_ReturnValue.Columns.Add("Location", GetType(System.String))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        For Ct = 1 To 65536
            Select Case pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                Case "[#]Sections"
                    CtStart = Ct
                Case "[#]End_Sections"
                    CtEnd = Ct
                    Exit For
            End Select
        Next

        Dim Sections_Ct As Int32 = 0

        For Ct = CtStart To CtEnd
            Dim ExcelText As String = ""
            ExcelText = pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text

            Dim Section_Type As String = ""
            Dim Section_Location As String = ""

            Select Case True
                Case InStr(ExcelText, "[") > 0
                Case Else
                    Try
                        Section_Type = Mid(ExcelText, 1, (InStr(ExcelText, " ")) - 1)
                        Section_Location = Mid(ExcelText, (InStr(ExcelText, " ") + 1))

                        Sections_Ct = Sections_Ct + 1

                        Dim NewRow As DataRow
                        NewRow = Dt_ReturnValue.NewRow
                        NewRow.Item("Ct") = Sections_Ct
                        NewRow.Item("Type") = Section_Type
                        NewRow.Item("Location") = Section_Location

                        Dt_ReturnValue.Rows.Add(NewRow)

                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                        Throw New Exception(ex.Message & vbCrLf & "Exception Code 1101. Invalid Syntax in [#]Sections.")
                    End Try

            End Select
        Next

        Return Dt_ReturnValue

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetDataTables(ByRef pSheet_Parameters As NativeExcel.IWorksheet) As DataTable

        Dim Dt_Tables As New DataTable
        Dt_Tables.Columns.Add("Ct", GetType(System.Int32))
        Dt_Tables.Columns.Add("Name", GetType(System.String))
        Dt_Tables.Columns.Add("GroupName", GetType(System.String))
        Dt_Tables.Columns.Add("SourceKey", GetType(System.String))
        Dt_Tables.Columns.Add("TargetKey", GetType(System.String))
        Dt_Tables.Columns.Add("IsSubTable", GetType(System.Boolean))
        Dt_Tables.Columns.Add("Location", GetType(System.String))
        Dt_Tables.Columns.Add("Items", GetType(System.Int32))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        For Ct = 1 To 65536
            Select Case pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                Case "[#]DataTable"
                    CtStart = Ct
                Case "[#]End_DataTable"
                    CtEnd = Ct
                    Exit For
            End Select
        Next

        If (CtStart = 0 And CtEnd = 0) Then
            Throw New Exception("Invalid Syntax in [#]DataTable.")
        End If

        Dim DataTable_Ct As Int32 = 0
        For Ct = CtStart To CtEnd
            Dim ExcelText As String = ""
            Try
                ExcelText = pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
            Catch
            End Try

            Dim DataTable_Name As String = ""
            Dim DataTable_GroupName As String = ""
            Dim DataTable_SourceKey As String = ""
            Dim DataTable_TargetKey As String = ""
            Dim DataTable_Location As String = ""

            Select Case True
                Case InStr(ExcelText, "[") > 0
                Case Else
                    Try
                        Dim Inner_ArExcelText() As String = Split(ExcelText, " ")

                        DataTable_Name = Inner_ArExcelText(0)
                        DataTable_Location = Inner_ArExcelText(1)

                        Try
                            DataTable_GroupName = Inner_ArExcelText(2)
                            DataTable_SourceKey = Inner_ArExcelText(3)
                            DataTable_TargetKey = Inner_ArExcelText(4)
                        Catch
                        End Try

                        DataTable_Ct = DataTable_Ct + 1

                        Dim NewRow As DataRow
                        NewRow = Dt_Tables.NewRow
                        NewRow.Item("Ct") = DataTable_Ct
                        NewRow.Item("Name") = DataTable_Name
                        NewRow.Item("Location") = DataTable_Location

                        If DataTable_GroupName.Trim <> "" Then
                            NewRow.Item("GroupName") = DataTable_GroupName
                            NewRow.Item("SourceKey") = DataTable_SourceKey
                            NewRow.Item("TargetKey") = DataTable_TargetKey
                            NewRow.Item("IsSubTable") = True
                        End If

                        Dt_Tables.Rows.Add(NewRow)

                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                        Throw New Exception(ex.Message & vbCrLf & "Invalid Syntax in [#]DataTable.")
                    End Try

            End Select
        Next

        Return Dt_Tables

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetDataTable_Headers(ByRef pSheet_Parameters As NativeExcel.IWorksheet) As DataTable

        Dim Dt_Tables_Headers As New DataTable
        Dt_Tables_Headers.Columns.Add("Ct", GetType(System.Int32))
        Dt_Tables_Headers.Columns.Add("Name", GetType(System.String))
        Dt_Tables_Headers.Columns.Add("Location", GetType(System.String))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        For Ct = 1 To 65536
            Select Case pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                Case "[#]DataTable_Header"
                    CtStart = Ct
                Case "[#]End_DataTable_Header"
                    CtEnd = Ct
                    Exit For
            End Select
        Next

        If (CtStart = 0 And CtEnd = 0) Then
            'Throw New Exception("Invalid Syntax in [#]DataTable_Header.")
            Return Dt_Tables_Headers
        End If

        Dim DataTable_Ct As Int32 = 0
        For Ct = CtStart To CtEnd
            Dim ExcelText As String = ""
            Try
                ExcelText = pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
            Catch
            End Try


            Dim Name As String = ""
            Dim Location As String = ""

            Select Case True
                Case InStr(ExcelText, "[") > 0
                Case Else
                    Try
                        Name = Mid(ExcelText, 1, (InStr(ExcelText, " ")) - 1)
                        Location = Mid(ExcelText, (InStr(ExcelText, " ") + 1))

                        DataTable_Ct = DataTable_Ct + 1

                        Dim NewRow As DataRow
                        NewRow = Dt_Tables_Headers.NewRow
                        NewRow.Item("Ct") = DataTable_Ct
                        NewRow.Item("Name") = Name
                        NewRow.Item("Location") = Location

                        Dt_Tables_Headers.Rows.Add(NewRow)

                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                        Throw New Exception(ex.Message & vbCrLf & "Invalid Syntax in [#]DataTable_Header.")
                    End Try

            End Select
        Next

        Return Dt_Tables_Headers

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetDataTable_Footers(ByRef pSheet_Parameters As NativeExcel.IWorksheet) As DataTable

        Dim Dt_Tables_Footers As New DataTable
        Dt_Tables_Footers.Columns.Add("Ct", GetType(System.Int32))
        Dt_Tables_Footers.Columns.Add("Name", GetType(System.String))
        Dt_Tables_Footers.Columns.Add("Location", GetType(System.String))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        For Ct = 1 To 65536
            Select Case pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                Case "[#]DataTable_Footer"
                    CtStart = Ct
                Case "[#]End_DataTable_Footer"
                    CtEnd = Ct
                    Exit For
            End Select
        Next


        If (CtStart = 0 And CtEnd = 0) Then
            'Throw New Exception("Invalid Syntax in [#]DataTable_Footer.")
            Return Dt_Tables_Footers
        End If

        Dim DataTable_Ct As Int32 = 0
        For Ct = CtStart To CtEnd
            Dim ExcelText As String = ""
            Try
                ExcelText = pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
            Catch
            End Try


            Dim Name As String = ""
            Dim Location As String = ""

            Select Case True
                Case InStr(ExcelText, "[") > 0
                Case Else
                    Try
                        Name = Mid(ExcelText, 1, (InStr(ExcelText, " ")) - 1)
                        Location = Mid(ExcelText, (InStr(ExcelText, " ") + 1))

                        DataTable_Ct = DataTable_Ct + 1

                        Dim NewRow As DataRow
                        NewRow = Dt_Tables_Footers.NewRow
                        NewRow.Item("Ct") = DataTable_Ct
                        NewRow.Item("Name") = Name
                        NewRow.Item("Location") = Location

                        Dt_Tables_Footers.Rows.Add(NewRow)

                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                        Throw New Exception(ex.Message & vbCrLf & "Invalid Syntax in [#]DataTable_Footer.")
                    End Try

            End Select
        Next

        Return Dt_Tables_Footers

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetDataTableFields(ByVal pDt_Tables As DataTable, ByRef pSheet_Template As NativeExcel.IWorksheet) As DataTable

        Dim Dt_Tables_Fields As New DataTable
        Dt_Tables_Fields.Columns.Add("Ct", GetType(System.Int32))
        Dt_Tables_Fields.Columns.Add("DtDataTable_Ct", GetType(System.Int32))
        Dt_Tables_Fields.Columns.Add("Name", GetType(System.String))
        Dt_Tables_Fields.Columns.Add("Position", GetType(System.Int32))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        Dim StFields As String = ""
        Dim AdrDtDataTable() As DataRow
        AdrDtDataTable = pDt_Tables.Select("", "Ct")
        If AdrDtDataTable.Length > 0 Then
            For Each Dr As DataRow In AdrDtDataTable

                Ct = 0
                StFields = ""
                Dim Delimiter As String = ""

                Dim DataTable_Width As Int32 = 0
                Dim Excel_Range() As Int32 = ParseExcelRange(Dr.Item("Location"))

                DataTable_Width = Excel_Range(2) - Excel_Range(0)

                Dim DtDataTable_Fields_Ct As Int32 = 0

                For Ct = 0 To DataTable_Width
                    Dim ExcelText As String = ""
                    Try
                        ExcelText = pSheet_Template.Range(GenerateChr(Excel_Range(0) + Ct) & Excel_Range(1)).Characters.Text
                    Catch
                    End Try


                    Dim Field As String = ""
                    Select Case True
                        Case InStr(ExcelText, "[") > 0
                            Field = Mid(ExcelText, InStr(ExcelText, "[") + 1, (InStrRev(ExcelText, "]") - Len("]")) - 1)

                            DtDataTable_Fields_Ct = DtDataTable_Fields_Ct + 1

                            Dim NewRow As DataRow
                            NewRow = Dt_Tables_Fields.NewRow
                            NewRow.Item("Ct") = DtDataTable_Fields_Ct
                            NewRow.Item("DtDataTable_Ct") = Dr.Item("Ct")
                            NewRow.Item("Name") = Field
                            NewRow.Item("Position") = Ct

                            Dt_Tables_Fields.Rows.Add(NewRow)

                    End Select
                Next
            Next
        End If

        Return Dt_Tables_Fields

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetDataTable_Pivot(ByRef pSheet_Parameters As NativeExcel.IWorksheet) As DataTable

        Dim Dt_Pivot As New DataTable
        Dt_Pivot.Columns.Add("Ct", GetType(System.Int32))
        Dt_Pivot.Columns.Add("Name", GetType(System.String))
        Dt_Pivot.Columns.Add("ParentTableName", GetType(System.String))
        Dt_Pivot.Columns.Add("SourceKey", GetType(System.String))
        Dt_Pivot.Columns.Add("TargetKey", GetType(System.String))
        Dt_Pivot.Columns.Add("Location", GetType(System.String))
        Dt_Pivot.Columns.Add("Items", GetType(System.Int32))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        For Ct = 1 To 65536
            Select Case pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                Case "[#]DataTable_Pivot"
                    CtStart = Ct
                Case "[#]End_DataTable_Pivot"
                    CtEnd = Ct
                    Exit For
            End Select
        Next

        If (CtStart = 0 And CtEnd = 0) Then
            Return Dt_Pivot
        End If

        Dim DataTable_Ct As Int32 = 0
        For Ct = CtStart To CtEnd
            Dim ExcelText As String = ""
            Try
                ExcelText = pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
            Catch
            End Try

            Dim Name As String = ""
            Dim ParentTableName As String = ""
            Dim SourceKey As String = ""
            Dim TargetKey As String = ""
            Dim Location As String = ""

            Select Case True
                Case InStr(ExcelText, "[") > 0
                Case Else
                    Try
                        Dim Inner_ArExcelText() As String = Split(ExcelText, " ")

                        Name = Inner_ArExcelText(0)
                        Location = Inner_ArExcelText(1)

                        Try
                            ParentTableName = Inner_ArExcelText(2)
                            SourceKey = Inner_ArExcelText(3)
                            TargetKey = Inner_ArExcelText(4)
                        Catch
                        End Try

                        DataTable_Ct = DataTable_Ct + 1

                        Dim NewRow As DataRow
                        NewRow = Dt_Pivot.NewRow
                        NewRow.Item("Ct") = DataTable_Ct
                        NewRow.Item("Name") = Name
                        NewRow.Item("Location") = Location

                        NewRow.Item("ParentTableName") = ParentTableName
                        NewRow.Item("SourceKey") = SourceKey
                        NewRow.Item("TargetKey") = TargetKey

                        Dt_Pivot.Rows.Add(NewRow)

                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                        Throw New Exception(ex.Message & vbCrLf & "Invalid Syntax in [#]DataTable_Pivot.")
                    End Try

            End Select
        Next

        Return Dt_Pivot

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetDataTable_Pivot_Fields(ByRef pDt_Pivot As DataTable, ByRef pSheet_Template As NativeExcel.IWorksheet) As DataTable

        Dim Dt_Pivot_Fields As New DataTable
        Dt_Pivot_Fields.Columns.Add("Ct", GetType(System.Int32))
        Dt_Pivot_Fields.Columns.Add("DtDataTable_Pivot_Ct", GetType(System.Int32))
        Dt_Pivot_Fields.Columns.Add("Name", GetType(System.String))
        Dt_Pivot_Fields.Columns.Add("Position", GetType(System.Int32))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        Dim StFields As String = ""
        Dim Arr_Dr_DtDataTable_Pivot() As DataRow
        Arr_Dr_DtDataTable_Pivot = pDt_Pivot.Select("", "Ct")
        If Arr_Dr_DtDataTable_Pivot.Length > 0 Then
            For Each Dr As DataRow In Arr_Dr_DtDataTable_Pivot
                Ct = 0
                StFields = ""
                Dim Delimiter As String = ""

                Dim DataTable_Width As Int32 = 0
                Dim Excel_Range() As Int32 = ParseExcelRange(Dr.Item("Location"))

                DataTable_Width = Excel_Range(2) - Excel_Range(0)

                Dim DtDataTable_Pivot_Fields_Ct As Int32 = 0

                For Ct = 0 To DataTable_Width
                    Dim ExcelText As String = ""
                    Try
                        ExcelText = pSheet_Template.Range(GenerateChr(Excel_Range(0) + Ct) & Excel_Range(1)).Characters.Text
                    Catch
                    End Try

                    Dim Field As String = ""
                    Select Case True
                        Case InStr(ExcelText, "[") > 0
                            Field = Mid(ExcelText, InStr(ExcelText, "[") + 1, (InStrRev(ExcelText, "]") - Len("]")) - 1)
                            DtDataTable_Pivot_Fields_Ct = DtDataTable_Pivot_Fields_Ct + 1

                            Dim NewRow As DataRow
                            NewRow = Dt_Pivot_Fields.NewRow
                            NewRow.Item("Ct") = DtDataTable_Pivot_Fields_Ct
                            NewRow.Item("DtDataTable_Pivot_Ct") = Dr.Item("Ct")
                            NewRow.Item("Name") = Field
                            NewRow.Item("Position") = Ct

                            Dt_Pivot_Fields.Rows.Add(NewRow)
                    End Select
                Next
            Next
        End If

        Return Dt_Pivot_Fields

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetDataTable_Pivot_Header(ByRef pSheet_Parameters As NativeExcel.IWorksheet) As DataTable

        Dim Dt_Pivot_Header As New DataTable
        Dt_Pivot_Header.Columns.Add("Ct", GetType(System.Int32))
        Dt_Pivot_Header.Columns.Add("Name", GetType(System.String))
        Dt_Pivot_Header.Columns.Add("Location", GetType(System.String))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        For Ct = 1 To 65536
            Select Case pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                Case "[#]DataTable_Pivot_Header"
                    CtStart = Ct
                Case "[#]End_DataTable_Pivot_Header"
                    CtEnd = Ct
                    Exit For
            End Select
        Next

        If (CtStart = 0 And CtEnd = 0) Then
            Return Dt_Pivot_Header
        End If

        Dim DataTable_Ct As Int32 = 0
        For Ct = CtStart To CtEnd
            Dim ExcelText As String = ""
            Try
                ExcelText = pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
            Catch
            End Try

            Dim Name As String = ""
            Dim Location As String = ""

            Select Case True
                Case InStr(ExcelText, "[") > 0
                Case Else
                    Try
                        Name = Mid(ExcelText, 1, (InStr(ExcelText, " ")) - 1)
                        Location = Mid(ExcelText, (InStr(ExcelText, " ") + 1))

                        DataTable_Ct = DataTable_Ct + 1

                        Dim NewRow As DataRow
                        NewRow = Dt_Pivot_Header.NewRow
                        NewRow.Item("Ct") = DataTable_Ct
                        NewRow.Item("Name") = Name
                        NewRow.Item("Location") = Location

                        Dt_Pivot_Header.Rows.Add(NewRow)

                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                        Throw New Exception(ex.Message & vbCrLf & "Invalid Syntax in [#]DataTable_Pivot_Header.")
                    End Try

            End Select
        Next

        Return Dt_Pivot_Header

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetDataTable_Pivot_Header_Fields(ByRef pDt_Pivot_Header As DataTable, ByRef pSheet_Template As NativeExcel.IWorksheet) As DataTable

        Dim Dt_Pivot_Header_Fields As New DataTable
        Dt_Pivot_Header_Fields.Columns.Add("Ct", GetType(System.Int32))
        Dt_Pivot_Header_Fields.Columns.Add("DtDataTable_Pivot_Header_Ct", GetType(System.Int32))
        Dt_Pivot_Header_Fields.Columns.Add("Name", GetType(System.String))
        Dt_Pivot_Header_Fields.Columns.Add("Position", GetType(System.Int32))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        Dim StFields As String = ""
        Dim Arr_Dr_DtDataTable_Header_Pivot() As DataRow
        Arr_Dr_DtDataTable_Header_Pivot = pDt_Pivot_Header.Select("", "Ct")
        If Arr_Dr_DtDataTable_Header_Pivot.Length > 0 Then
            For Each Dr As DataRow In Arr_Dr_DtDataTable_Header_Pivot
                Ct = 0
                StFields = ""
                Dim Delimiter As String = ""

                Dim DataTable_Width As Int32 = 0
                Dim Excel_Range() As Int32 = ParseExcelRange(Dr.Item("Location"))

                DataTable_Width = Excel_Range(2) - Excel_Range(0)

                Dim DtDataTable_Pivot_Header_Fields_Ct As Int32 = 0

                For Ct = 0 To DataTable_Width
                    Dim ExcelText As String = ""
                    Try
                        ExcelText = pSheet_Template.Range(GenerateChr(Excel_Range(0) + Ct) & Excel_Range(1)).Characters.Text
                    Catch
                    End Try

                    Dim Field As String = ""
                    Select Case True
                        Case InStr(ExcelText, "[") > 0
                            Field = Mid(ExcelText, InStr(ExcelText, "[") + 1, (InStrRev(ExcelText, "]") - Len("]")) - 1)

                            DtDataTable_Pivot_Header_Fields_Ct = DtDataTable_Pivot_Header_Fields_Ct + 1
                            Dim NewRow As DataRow
                            NewRow = Dt_Pivot_Header_Fields.NewRow
                            NewRow.Item("Ct") = DtDataTable_Pivot_Header_Fields_Ct
                            NewRow.Item("DtDataTable_Pivot_Header_Ct") = Dr.Item("Ct")
                            NewRow.Item("Name") = Field
                            NewRow.Item("Position") = Ct

                            Dt_Pivot_Header_Fields.Rows.Add(NewRow)
                    End Select
                Next
            Next
        End If

        Return Dt_Pivot_Header_Fields

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetDataTable_Pivot_Totals(ByRef pSheet_Parameters As NativeExcel.IWorksheet) As DataTable

        Dim Dt_Pivot_Totals As New DataTable
        Dt_Pivot_Totals.Columns.Add("Ct", GetType(System.Int32))
        Dt_Pivot_Totals.Columns.Add("Name", GetType(System.String))
        Dt_Pivot_Totals.Columns.Add("Location", GetType(System.String))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        For Ct = 1 To 65536
            Select Case pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                Case "[#]DataTable_Pivot_Totals"
                    CtStart = Ct
                Case "[#]End_DataTable_Pivot_Totals"
                    CtEnd = Ct
                    Exit For
            End Select
        Next

        If (CtStart = 0 And CtEnd = 0) Then
            Return Dt_Pivot_Totals
        End If

        Dim DataTable_Ct As Int32 = 0
        For Ct = CtStart To CtEnd
            Dim ExcelText As String = ""
            Try
                ExcelText = pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
            Catch
            End Try

            Dim Name As String = ""
            Dim Location As String = ""

            Select Case True
                Case InStr(ExcelText, "[") > 0
                Case Else
                    Try
                        Name = Mid(ExcelText, 1, (InStr(ExcelText, " ")) - 1)
                        Location = Mid(ExcelText, (InStr(ExcelText, " ") + 1))

                        DataTable_Ct = DataTable_Ct + 1

                        Dim NewRow As DataRow
                        NewRow = Dt_Pivot_Totals.NewRow
                        NewRow.Item("Ct") = DataTable_Ct
                        NewRow.Item("Name") = Name
                        NewRow.Item("Location") = Location

                        Dt_Pivot_Totals.Rows.Add(NewRow)

                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                        Throw New Exception(ex.Message & vbCrLf & "Invalid Syntax in [#]DataTable_Pivot_Totals.")
                    End Try

            End Select
        Next

        Return Dt_Pivot_Totals

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetDataTable_Pivot_Totals_Fields(ByRef pDt_Pivot_Totals As DataTable, ByRef pSheet_Template As NativeExcel.IWorksheet) As DataTable

        Dim Dt_Pivot_Totals_Fields As New DataTable
        Dt_Pivot_Totals_Fields.Columns.Add("Ct", GetType(System.Int32))
        Dt_Pivot_Totals_Fields.Columns.Add("DtDataTable_Pivot_Totals_Ct", GetType(System.Int32))
        Dt_Pivot_Totals_Fields.Columns.Add("Name", GetType(System.String))
        Dt_Pivot_Totals_Fields.Columns.Add("Position", GetType(System.Int32))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        Dim StFields As String = ""
        Dim Arr_Dr_DtDataTable_Pivot() As DataRow
        Arr_Dr_DtDataTable_Pivot = pDt_Pivot_Totals.Select("", "Ct")
        If Arr_Dr_DtDataTable_Pivot.Length > 0 Then
            For Each Dr As DataRow In Arr_Dr_DtDataTable_Pivot
                Ct = 0
                StFields = ""
                Dim Delimiter As String = ""

                Dim DataTable_Width As Int32 = 0
                Dim Excel_Range() As Int32 = ParseExcelRange(Dr.Item("Location"))

                DataTable_Width = Excel_Range(2) - Excel_Range(0)

                Dim DtDataTable_Pivot_Fields_Ct As Int32 = 0

                For Ct = 0 To DataTable_Width
                    Dim ExcelText As String = ""
                    Try
                        ExcelText = pSheet_Template.Range(GenerateChr(Excel_Range(0) + Ct) & Excel_Range(1)).Characters.Text
                    Catch
                    End Try

                    Dim Field As String = ""
                    Select Case True
                        Case InStr(ExcelText, "[") > 0
                            Field = Mid(ExcelText, InStr(ExcelText, "[") + 1, (InStrRev(ExcelText, "]") - Len("]")) - 1)
                            DtDataTable_Pivot_Fields_Ct = DtDataTable_Pivot_Fields_Ct + 1

                            Dim NewRow As DataRow
                            NewRow = Dt_Pivot_Totals_Fields.NewRow
                            NewRow.Item("Ct") = DtDataTable_Pivot_Fields_Ct
                            NewRow.Item("DtDataTable_Pivot_Totals_Ct") = Dr.Item("Ct")
                            NewRow.Item("Name") = Field
                            NewRow.Item("Position") = Ct

                            Dt_Pivot_Totals_Fields.Rows.Add(NewRow)
                    End Select
                Next
            Next
        End If

        Return Dt_Pivot_Totals_Fields

    End Function

    Shared Function NativeExcel_CreateExcelDocument_GetParameters(ByRef pSheet_Parameters As NativeExcel.IWorksheet, ByVal pParameters() As String, ByVal pParametersValue() As Object) As DataTable

        Dim DtParameters As New DataTable
        DtParameters.Columns.Add("ParameterName", GetType(String))
        DtParameters.Columns.Add("ParameterType", GetType(String))
        DtParameters.Columns.Add("ParameterValue", GetType(String))

        Dim CtStart As Int32 = 0
        Dim CtEnd As Int32 = 0
        Dim Ct As Int32 = 0

        For Ct = 1 To 65536
            Select Case pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text
                Case "[#]Parameters"
                    CtStart = Ct
                Case "[#]End_Parameters"
                    CtEnd = Ct
                    Exit For
            End Select
        Next

        If (CtStart = 0 And CtEnd = 0) Then
            Return DtParameters
        End If

        For Ct = CtStart To CtEnd
            Dim ExcelText As String = ""
            ExcelText = pSheet_Parameters.Range("A" & Ct.ToString).Characters.Text

            Dim ParameterName As String = ""
            Dim ParameterType As String = ""

            Select Case True
                Case InStr(ExcelText, "@") > 0
                    Try
                        ParameterName = Mid(ExcelText, Len("@") + 1, (InStr(ExcelText, " ") - Len("@")) - 1)
                        ParameterType = Mid(ExcelText, (InStr(ExcelText, " ") + 1))

                        Dim NewRow As DataRow
                        NewRow = DtParameters.NewRow
                        NewRow.Item("ParameterName") = ParameterName
                        NewRow.Item("ParameterType") = ParameterType

                        DtParameters.Rows.Add(NewRow)
                    Catch
                    End Try
            End Select
        Next

        If Not pParameters Is Nothing Then
            Ct = 0
            For Ct = 0 To pParameters.Length - 1
                Dim InnerAdr() As DataRow
                InnerAdr = DtParameters.Select("ParameterName = '" & pParameters(Ct) & "'")
                If InnerAdr.Length > 0 Then
                    InnerAdr(0).Item("ParameterValue") = pParametersValue(Ct)
                End If
            Next
        End If

        Return DtParameters

    End Function

End Class
