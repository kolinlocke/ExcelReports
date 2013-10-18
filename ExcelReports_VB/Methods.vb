#Region " Imports "

Imports System.Runtime.InteropServices

#End Region

Friend Class Methods

    Public Shared Sub FixString(ByRef StringInput As String)
        Dim St As String
        Dim ArSt() As Char

        St = StringInput
        ArSt = St.ToCharArray

        St = ""

        For Each Ch As Char In ArSt
            If Not Ch = Nothing Then
                St = St & Ch.ToString
            End If

        Next

        StringInput = St

    End Sub

    Public Shared Function Decrypt(ByVal txt As String, ByVal cryptkey As String) As String
        Dim keyval As String
        Dim kv() As String
        Dim estr As String
        Dim tmp As String
        Dim i As Integer

        If (txt = "") Then
            Return ""
        End If

        If (cryptkey = "") Then
            Return ""
        End If

        keyval = KeyValue(cryptkey)
        kv = Split(keyval, "/")
        estr = ""
        tmp = ""
        For i = 1 To Len(txt)
            If (Microsoft.VisualBasic.Left(Mid(txt, i), 1) <> "") Then
                If (Asc(Microsoft.VisualBasic.Left(Mid(txt, i), 1)) > 64) And (Asc(Microsoft.VisualBasic.Left(Mid(txt, i), 1)) < 91) Then
                    If (tmp <> "") Then
                        tmp = Int(tmp / Int(kv(1)))
                        tmp = Int(tmp - Int(kv(0)))
                        estr = estr & Chr(tmp)
                        tmp = ""
                    End If
                Else
                    tmp = tmp & Microsoft.VisualBasic.Left(Mid(txt, i), 1)
                End If
            End If
        Next

        tmp = Int(tmp / Int(kv(1)))
        tmp = Int(tmp - Int(kv(0)))
        estr = estr & Chr(tmp)
        Return estr
    End Function

    Public Shared Function Encrypt(ByVal txt As String, ByVal cryptkey As String) As String
        If (txt = "") Then
            Return ""
        End If
        If (cryptkey = "") Then
            Return ""
        End If
        Dim keyval As String
        Dim kv() As String
        Dim estr As String = ""
        Dim i As Integer
        Dim e As String
        Dim rndval As Integer
        keyval = KeyValue(cryptkey)
        kv = Split(keyval, "/")
        For i = 0 To Len(txt)
            e = Mid(txt, i + 1)
            e = Microsoft.VisualBasic.Left(e, 1)
            If (e <> "") Then
                e = Asc(e)
                e = Int(Int(e) + Int(kv(0)))
                e = Int(Int(e) * Int(kv(1)))
                Randomize()
                rndval = Int((90 - 65 + 1) * Rnd() + 65)
                estr = estr & Chr(rndval) & e
            End If
        Next
        Return estr
    End Function

    Public Shared Function KeyValue(ByVal cryptkey As String) As String
        Dim keyval1 As Integer
        Dim keyval2 As Integer
        keyval1 = 0
        keyval2 = 0
        Dim i As Integer
        Dim curchr As String
        i = 1
        For i = 1 To Len(cryptkey)
            curchr = Mid(cryptkey, i + 1)
            curchr = Microsoft.VisualBasic.Left(curchr, 1)
            If curchr <> "" Then
                curchr = Asc(curchr)
                keyval1 = Int(keyval1 + curchr)
                keyval2 = Len(cryptkey)
            End If
        Next
        Return (keyval1 & "/" & keyval2)
    End Function

    Shared Function WriteTempFile(ByVal Stream As System.IO.Stream, ByVal FilePath As String) As Boolean

        Try
            Dim FileStream As System.IO.FileStream

            If Not IO.Directory.Exists(Mid(FilePath, 1, InStrRev(FilePath, "\"))) Then
                IO.Directory.CreateDirectory(Mid(FilePath, 1, InStrRev(FilePath, "\")))
            End If

            If IO.File.Exists(FilePath) Then
                IO.File.Delete(FilePath)

            End If

            Dim ByteBuffer As Byte()
            ReDim ByteBuffer(Stream.Length - 1)
            Stream.Read(ByteBuffer, 0, Stream.Length)

            IO.File.Create(FilePath).Close()
            IO.File.SetAttributes(FilePath, IO.FileAttributes.Hidden + IO.FileAttributes.Temporary + IO.FileAttributes.System)

            FileStream = New System.IO.FileStream(FilePath, System.IO.FileMode.Open, System.IO.FileAccess.Write)
            FileStream.Write(ByteBuffer, 0, Stream.Length)
            FileStream.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Shared Function DeleteTempFile(ByVal FilePath As String) As Boolean
        Try
            If IO.File.Exists(FilePath) Then
                IO.File.Delete(FilePath)
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Shared Function IsNull(ByVal Input As Object, ByVal NullOutput As Object) As Object
        Return IIf(IsDBNull(Input), NullOutput, Input)
    End Function

    Shared Function TextFiller(ByVal TextInput As String, ByVal Filler As String, ByVal TextLength As Int32) As String
        Dim ReturnValue As String = ""
        ReturnValue = Right(StrDup(TextLength, Filler) & LTrim(Left(TextInput, TextLength)), TextLength)
        Return ReturnValue
    End Function

    Shared Sub TextWriter(ByVal Filename As String, ByVal Path As String, ByVal strLine() As String)
        Try
            Dim str As String = ""
            Dim i As Integer

            Dim FileWriter As System.IO.StreamWriter
            Dim FileStream As System.IO.FileStream

            If Not IO.Directory.Exists(Path) Then
                IO.Directory.CreateDirectory(Path)
            End If

            If Not IO.File.Exists(Path & IIf(InStr(Mid(Path, Len(Path)), "\", CompareMethod.Text), "", "\") & Filename & "") Then
                FileStream = New System.IO.FileStream(Path & IIf(InStr(Mid(Path, Len(Path)), "\", CompareMethod.Text), "", "\") & Filename & "", System.IO.FileMode.CreateNew, System.IO.FileAccess.Write)
            Else
                FileStream = New System.IO.FileStream(Path & IIf(InStr(Mid(Path, Len(Path)), "\", CompareMethod.Text), "", "\") & Filename & "", System.IO.FileMode.Append, System.IO.FileAccess.Write)
            End If

            FileWriter = New System.IO.StreamWriter(FileStream)

            If Not strLine Is Nothing Then
                For i = 0 To UBound(strLine)
                    str = str & strLine(i) & "  "
                Next
            End If
            FileWriter.WriteLine(str)
            FileWriter.Flush()
            FileWriter.Close()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Shared Sub SetBit_Desc(ByRef Row As DataRow, ByVal FieldName As String, ByVal Field_DescName As String, Optional ByVal TrueDesc As String = "True", Optional ByVal FalseDesc As String = "False")
        Select Case IsNull(Row.Item(FieldName), False)
            Case True
                Row.Item(Field_DescName) = TrueDesc
            Case False
                Row.Item(Field_DescName) = FalseDesc
        End Select
    End Sub

    Shared Function GenerateChr(ByVal InputInt As Int32) As String
        Dim Ct As Int32 = 0
        Dim TmpRes As Int32 = 0
        Dim OutputChr As String = ""

        While (26 ^ Ct) < InputInt
            Ct = Ct + 1
        End While

        If Ct > 0 Then
            Ct = Ct - 1
        End If

        While InputInt > 0
            TmpRes = InputInt \ (26 ^ Ct)
            If ((InputInt Mod 26) = 0) And Ct > 0 Then
                TmpRes = TmpRes - Ct
            End If
            OutputChr = OutputChr + Chr(TmpRes + 64)
            InputInt = InputInt - ((26 ^ Ct) * TmpRes)
            Ct = Ct - 1
        End While

        Return OutputChr

    End Function

    Shared Function ParseExcelRange(ByVal Excel_Range As String) As Int32()
        Try
            'Returns Int32 Array (4 Dimensions)
            'Sample: Input = "A1:J5"
            'Return_Int(0) = 1 (A)
            'Return_Int(1) = 1 (1)
            'Return_Int(2) = 10 (J)
            'Return_Int(3) = 5 (5)

            Dim Return_Int(3) As Int32
            Return_Int(0) = 0
            Return_Int(1) = 0
            Return_Int(2) = 0
            Return_Int(3) = 0

            Dim Tmp_Excel_Range As String = ""
            Tmp_Excel_Range = Excel_Range

            Dim St_Range1 As String = ""
            Dim St_Range2 As String = ""

            St_Range1 = Mid(Tmp_Excel_Range, 1, InStr(Tmp_Excel_Range, ":") - 1)
            St_Range2 = Mid(Tmp_Excel_Range, InStr(Tmp_Excel_Range, ":") + 1)

            Dim St_Parsed1 As String = ""
            Dim St_Parsed2 As String = ""

            For Each Ch As Char In St_Range1
                Select Case True
                    Case Not IsNumeric(Ch)
                        St_Parsed1 = (St_Parsed1 & Ch.ToString).ToUpper
                    Case Else
                        St_Parsed2 = (St_Parsed2 & Ch.ToString)
                End Select

            Next


            Dim Digit As Int32 = 0
            Digit = St_Parsed1.Length - 1

            Dim Result As Int32 = 0

            For Each InnerCh As Char In St_Parsed1
                Result = Result + ((26 ^ Digit) * (Asc(InnerCh) - 64))
                Digit = Digit - 1
            Next

            Return_Int(0) = Result
            Return_Int(1) = CType(St_Parsed2, Int32)

            '[-]

            St_Parsed1 = ""
            St_Parsed2 = ""

            For Each Ch As Char In St_Range2
                Select Case True
                    Case Not IsNumeric(Ch)
                        St_Parsed1 = (St_Parsed1 & Ch.ToString).ToUpper
                    Case Else
                        St_Parsed2 = (St_Parsed2 & Ch.ToString)
                End Select
            Next

            Digit = 0
            Digit = St_Parsed1.Length - 1

            Result = 0

            For Each InnerCh As Char In St_Parsed1
                Result = Result + ((26 ^ Digit) * (Asc(InnerCh) - 64))
                Digit = Digit - 1
            Next

            Return_Int(2) = Result
            Return_Int(3) = CType(St_Parsed2, Int32)

            Return Return_Int
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Shared Function ParseExcelRange_GetHeight(ByVal Excel_Range As String) As Int32
        Dim Range() As Int32 = ParseExcelRange(Excel_Range)
        Dim Length As Int32 = 0
        Length = (Range(3) - Range(1))
        Return Length
    End Function

    Shared Function ParseExcelRange_GetWidth(ByVal Excel_Range As String) As Int32
        Dim Range() As Int32 = ParseExcelRange(Excel_Range)
        Dim Length As Int32 = 0
        Length = (Range(2) - Range(0))
        Return Length
    End Function

    Shared Function ConvertDataForExcel(ByRef theExcelData As System.Text.StringBuilder, ByVal odt As DataTable, Optional ByVal PrintOnceField As String = "", Optional ByVal Fields() As String = Nothing, Optional ByVal QuoteFields() As String = Nothing, Optional ByVal RowStart As Int32 = -1, Optional ByVal RowEnd As Int32 = -1) As Boolean
        Try
            Dim i As Integer
            Dim ctr As Integer
            Dim ctr2 As Integer
            'Dim adr As DataRow
            Dim LastData As String = ""

            Dim RowCt As Int32 = 0
            Dim RowCtEnd As Int32 = 0

            If RowStart > -1 Then
                RowCt = RowStart
            End If

            If RowEnd > -1 Then
                RowCtEnd = RowEnd
            Else
                RowCtEnd = odt.Rows.Count - 1
            End If

            If RowCt >= odt.Rows.Count Then
                RowCt = odt.Rows.Count - 1
            End If

            If RowCtEnd >= odt.Rows.Count Then
                RowCtEnd = odt.Rows.Count - 1
            End If

            While RowCt <= RowCtEnd
                If Not Fields Is Nothing Then
                    For ctr = 0 To Fields.Length - 1
                        For i = 0 To odt.Columns.Count - 1
                            If Fields(ctr).Trim = "" Then
                                theExcelData.Append(vbTab)
                                Exit For
                            End If
                            If UCase(Fields(ctr)) = UCase(odt.Columns(i).ColumnName) Then
                                '
                                ' Convert the data and fill the string. Null values become blanks.
                                '
                                Dim AddSt As String = ""
                                If Not QuoteFields Is Nothing Then
                                    For ctr2 = 0 To QuoteFields.Length - 1
                                        If UCase(QuoteFields(ctr2)) = UCase(odt.Columns(i).ColumnName) Then
                                            AddSt = "'"
                                            Exit For
                                        End If
                                    Next
                                End If

                                If odt.Rows.Item(RowCt).Item(i) Is DBNull.Value Then
                                    theExcelData.Append("")
                                Else

                                    If PrintOnceField <> "" Then
                                        If UCase(odt.Columns(i).ColumnName) = UCase(PrintOnceField) Then

                                            If LastData <> odt.Rows.Item(RowCt).Item(i).ToString Then
                                                theExcelData.Append(AddSt & odt.Rows.Item(RowCt).Item(i).ToString)
                                            Else
                                                theExcelData.Append("")
                                            End If

                                            LastData = odt.Rows.Item(RowCt).Item(i).ToString
                                        Else
                                            theExcelData.Append(AddSt & odt.Rows.Item(RowCt).Item(i).ToString)
                                        End If
                                    Else
                                        theExcelData.Append(AddSt & odt.Rows.Item(RowCt).Item(i).ToString)
                                    End If
                                End If
                                theExcelData.Append(vbTab)
                            End If
                        Next

                    Next
                Else

                    For i = 0 To odt.Columns.Count - 1

                        '
                        ' Convert the data and fill the string. Null values become blanks.
                        '
                        Dim AddSt As String = ""
                        If Not QuoteFields Is Nothing Then
                            For ctr2 = 0 To QuoteFields.Length - 1
                                If UCase(QuoteFields(ctr2)) = UCase(odt.Columns(i).ColumnName) Then
                                    AddSt = "'"
                                    Exit For
                                End If
                            Next
                        End If

                        If odt.Rows.Item(RowCt).Item(i) Is DBNull.Value Then
                            theExcelData.Append("")
                        Else

                            If PrintOnceField <> "" Then
                                If UCase(odt.Columns(i).ColumnName) = UCase(PrintOnceField) Then

                                    If LastData <> odt.Rows.Item(RowCt).Item(i).ToString Then
                                        theExcelData.Append(AddSt & odt.Rows.Item(RowCt).Item(i).ToString)
                                    Else
                                        theExcelData.Append("")
                                    End If

                                    LastData = odt.Rows.Item(RowCt).Item(i).ToString
                                Else
                                    theExcelData.Append(AddSt & odt.Rows.Item(RowCt).Item(i).ToString)
                                End If
                            Else
                                theExcelData.Append(AddSt & odt.Rows.Item(RowCt).Item(i).ToString)
                            End If

                        End If

                        theExcelData.Append(vbTab)
                    Next

                End If
                '
                ' Add a line feed to the end of each row.
                '
                theExcelData.Append(vbCrLf)

                RowCt = RowCt + 1
            End While

            Return True

        Catch ex As Exception
            ' Display an error message.
            Return False
        End Try

    End Function

    Shared Function ConvertDataTo2DimArray(ByVal odt As DataTable, Optional ByVal PrintOnceField As String = "", Optional ByVal Fields() As String = Nothing, Optional ByVal QuoteFields() As String = Nothing, Optional ByVal RowStart As Int32 = -1, Optional ByVal RowEnd As Int32 = -1) As Object(,)



        Dim i As Integer
        Dim ctr As Integer = 0
        Dim ctr2 As Integer = 0
        Dim LastData As String = ""

        Dim RowCt As Int32 = 0
        Dim RowCtEnd As Int32 = 0

        If RowStart > -1 Then
            RowCt = RowStart
        Else
            RowStart = 0
        End If

        If RowEnd > -1 Then
            RowCtEnd = RowEnd
        Else
            RowCtEnd = odt.Rows.Count - 1
        End If

        If RowCt >= odt.Rows.Count Then
            RowCt = odt.Rows.Count - 1
        End If

        If RowCtEnd >= odt.Rows.Count Then
            RowCtEnd = odt.Rows.Count - 1
        End If

        Dim RowLength As Int32 = RowCtEnd - RowCt
        Dim ColumnLength As Int32 = 0
        If Not Fields Is Nothing Then
            ColumnLength = Fields.Length
        Else
            ColumnLength = odt.Columns.Count
        End If

        Dim ReturnValue(RowLength, ColumnLength) As Object

        If odt.Rows.Count = 0 Then
            Return ReturnValue
        End If

        While RowCt <= RowCtEnd

            Dim RV_RowCt As Int32 = RowCt - RowStart

            If Not Fields Is Nothing Then
                For ctr = 0 To Fields.Length - 1
                    Dim RV_ColumnCt As Int32 = ctr

                    For i = 0 To odt.Columns.Count - 1
                        If Fields(ctr).Trim = "" Then
                            'theExcelData.Append(vbTab)
                            ReturnValue(RV_RowCt, RV_ColumnCt) = ""
                            Exit For
                        End If

                        If UCase(Fields(ctr)) = UCase(odt.Columns(i).ColumnName) Then

                            'Dim AddSt As String = ""
                            'If Not QuoteFields Is Nothing Then
                            '    For ctr2 = 0 To QuoteFields.Length - 1
                            '        If UCase(QuoteFields(ctr2)) = UCase(odt.Columns(i).ColumnName) Then
                            '            AddSt = "'"
                            '            Exit For
                            '        End If
                            '    Next
                            'End If

                            If odt.Rows.Item(RowCt).Item(i) Is DBNull.Value Then
                                'theExcelData.Append("")
                                ReturnValue(RV_RowCt, RV_ColumnCt) = ""
                            Else

                                If PrintOnceField <> "" Then
                                    If UCase(odt.Columns(i).ColumnName) = UCase(PrintOnceField) Then
                                        If LastData <> odt.Rows.Item(RowCt).Item(i).ToString Then
                                            'ReturnValue(RV_RowCt, RV_ColumnCt) = AddSt & odt.Rows.Item(RowCt).Item(i).ToString
                                            ReturnValue(RV_RowCt, RV_ColumnCt) = odt.Rows.Item(RowCt).Item(i)
                                        Else
                                            'theExcelData.Append("")
                                            ReturnValue(RV_RowCt, RV_ColumnCt) = ""
                                        End If
                                        LastData = odt.Rows.Item(RowCt).Item(i)
                                    Else
                                        'ReturnValue(RV_RowCt, RV_ColumnCt) = AddSt & odt.Rows.Item(RowCt).Item(i)
                                        ReturnValue(RV_RowCt, RV_ColumnCt) = odt.Rows.Item(RowCt).Item(i)
                                    End If
                                Else
                                    'ReturnValue(RV_RowCt, RV_ColumnCt) = AddSt & odt.Rows.Item(RowCt).Item(i).ToString
                                    ReturnValue(RV_RowCt, RV_ColumnCt) = odt.Rows.Item(RowCt).Item(i)
                                End If
                            End If
                        End If
                    Next

                Next
            Else
                For i = 0 To odt.Columns.Count - 1
                    Dim RV_ColumnCt As Int32 = i

                    'Dim AddSt As String = ""
                    'If Not QuoteFields Is Nothing Then
                    '    For ctr2 = 0 To QuoteFields.Length - 1
                    '        If UCase(QuoteFields(ctr2)) = UCase(odt.Columns(i).ColumnName) Then
                    '            AddSt = "'"
                    '            Exit For
                    '        End If
                    '    Next
                    'End If

                    If odt.Rows.Item(RowCt).Item(i) Is DBNull.Value Then
                        ReturnValue(RV_RowCt, RV_ColumnCt) = ""
                    Else

                        If PrintOnceField <> "" Then
                            If UCase(odt.Columns(i).ColumnName) = UCase(PrintOnceField) Then
                                If LastData <> odt.Rows.Item(RowCt).Item(i).ToString Then
                                    'ReturnValue(RV_RowCt, RV_ColumnCt) = AddSt & odt.Rows.Item(RowCt).Item(i).ToString
                                    ReturnValue(RV_RowCt, RV_ColumnCt) = odt.Rows.Item(RowCt).Item(i)
                                Else
                                    ReturnValue(RV_RowCt, RV_ColumnCt) = ""
                                End If
                                LastData = odt.Rows.Item(RowCt).Item(i)
                            Else
                                'ReturnValue(RV_RowCt, RV_ColumnCt) = AddSt & odt.Rows.Item(RowCt).Item(i).ToString
                                ReturnValue(RV_RowCt, RV_ColumnCt) = odt.Rows.Item(RowCt).Item(i)
                            End If
                        Else
                            'ReturnValue(RV_RowCt, RV_ColumnCt) = AddSt & odt.Rows.Item(RowCt).Item(i).ToString
                            ReturnValue(RV_RowCt, RV_ColumnCt) = odt.Rows.Item(RowCt).Item(i)
                        End If

                    End If
                Next
            End If
            RowCt = RowCt + 1
        End While

        Return ReturnValue

    End Function

    Shared Function ConvertDataTo2DimArray(ByVal pDr As DataRow, Optional ByVal PrintOnceField As String = "", Optional ByVal Fields() As String = Nothing, Optional ByVal QuoteFields() As String = Nothing) As Object(,)

        Dim i As Integer
        Dim ctr As Integer
        'Dim ctr2 As Integer
        Dim LastData As String = ""

        Dim RowLength As Int32 = 1
        Dim ColumnLength As Int32 = 0
        If Not Fields Is Nothing Then
            ColumnLength = Fields.Length
        Else
            ColumnLength = pDr.Table.Columns.Count
        End If

        Dim ReturnValue(RowLength, ColumnLength) As Object

        Dim RV_RowCt As Int32 = 0

        If Not Fields Is Nothing Then
            For ctr = 0 To Fields.Length - 1
                Dim RV_ColumnCt As Int32 = ctr

                For i = 0 To pDr.Table.Columns.Count - 1
                    If Fields(ctr).Trim = "" Then
                        ReturnValue(RV_RowCt, RV_ColumnCt) = ""
                        Exit For
                    End If
                    If UCase(Fields(ctr)) = UCase(pDr.Table.Columns(i).ColumnName) Then
                        If pDr.Item(i) Is DBNull.Value Then
                            ReturnValue(RV_RowCt, RV_ColumnCt) = ""
                        Else
                            ReturnValue(RV_RowCt, RV_ColumnCt) = pDr.Item(i)
                        End If
                    End If
                Next
            Next
        Else
            For i = 0 To pDr.Table.Columns.Count - 1
                Dim RV_ColumnCt As Int32 = i

                If pDr.Item(i) Is DBNull.Value Then
                    ReturnValue(RV_RowCt, RV_ColumnCt) = ""
                Else
                    ReturnValue(RV_RowCt, RV_ColumnCt) = pDr.Item(i)
                End If
            Next
        End If

        Return ReturnValue

    End Function

    Shared Function ConvertDate(ByVal pDateTime As DateTime) As DateTime
        Dim rDateTime As New DateTime(DatePart(DateInterval.Year, pDateTime), DatePart(DateInterval.Month, pDateTime), DatePart(DateInterval.Day, pDateTime))
        Return rDateTime
    End Function

    Shared Function CreateSessionID() As String
        Dim CurrentDate As DateTime = Now
        Return CurrentDate.Year & CurrentDate.Month & CurrentDate.Day & CurrentDate.Hour & CurrentDate.Minute & CurrentDate.Second & CurrentDate.Millisecond
    End Function

    Shared Sub AddDataRow(ByRef Dt As DataTable, ByVal Fields() As String, ByVal Values() As Object)
        Dim Nr As DataRow
        Nr = Dt.NewRow

        For Ct As Int32 = 0 To (Fields.Length - 1)
            Nr.Item(Fields(Ct)) = Values(Ct)
        Next

        Dt.Rows.Add(Nr)
    End Sub

    Shared Function Generate_RandomPassword() As String
        Dim cEnc As New Encryption
        Dim cKey As String
        Dim cPassword As String

        Randomize()
        cPassword = (CType(Rnd() * 999999, Int32) + 1).ToString
        cKey = (CType(Rnd() * 999999, Int32) + 1).ToString

        cPassword = cEnc.hexEncrypt(cPassword, cKey).Replace("-", "").Replace(" ", "")

        Return cPassword

    End Function

    Shared Function ImportExcelData(ByVal PrmPathExcelFile As String, ByVal ExcelSheetName As String, Optional ByVal strCriteria As String = "", Optional ByVal strOrderBy As String = "") As DataTable
        Dim OleConn As System.Data.OleDb.OleDbConnection
        Dim oDS As System.Data.DataSet
        Dim OleDA As System.Data.OleDb.OleDbDataAdapter

        Dim str As String
        Try
            OleConn = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; " & _
            "data source=" & PrmPathExcelFile & ";" & "Extended Properties=Excel 8.0;")

            str = "SELECT * FROM [" & ExcelSheetName & "$] " & strCriteria & " " & strOrderBy
            OleDA = New System.Data.OleDb.OleDbDataAdapter(str, OleConn)

            OleDA.TableMappings.Add("Table", "Table")
            oDS = New System.Data.DataSet
            OleDA.Fill(oDS)

            Return oDS.Tables(0)
        Catch ex As Exception
            Throw ex
        Finally
            OleDA = Nothing
            oDS = Nothing
        End Try
    End Function

    Shared Function PrepareFilterText(ByVal Field As String, ByVal DataType As String, ByVal Filter As String) As String

        Dim ReturnValue As String = ""

        Select Case DataType
            Case "String"
                ReturnValue = "[" & Field & "] Like '" & Filter & "%'"
            Case Else
                Dim TmpFilterText As String = ""
                TmpFilterText = Filter
                If ParseFilterText(TmpFilterText, DataType) Then
                    ReturnValue = "[" & Field & "]" & TmpFilterText & ""
                End If
        End Select

        Return ReturnValue

    End Function

    Shared Function ParseFilterText(ByRef FilterText As String, Optional ByVal DataType As String = "String") As Boolean
        Dim Ct As Int32 = 0
        Dim aParsedTextToken As String = ""
        Dim aParsedText() As String

        aParsedText = Split(FilterText, " ")

        For Ct = 0 To aParsedText.Length - 1
            Select Case aParsedText(Ct)
                Case ">", "<", "=", "<=", ">=", "<>"
                    aParsedTextToken = aParsedTextToken & "Boolean"
                Case Else
                    If IsNumeric(aParsedText(Ct)) Then
                        aParsedTextToken = aParsedTextToken & "Numeric"
                    ElseIf IsDate(aParsedText(Ct)) Then
                        aParsedText(Ct) = "'" & Format(CType(aParsedText(Ct), DateTime), "yyyy-MM-dd") & "'"
                        aParsedTextToken = aParsedTextToken & "DateTime"
                    Else
                        aParsedTextToken = aParsedTextToken & "String"
                    End If
            End Select
        Next

        FilterText = Join(aParsedText, " ")

        Select Case aParsedTextToken
            Case "BooleanNumeric"
                Select Case DataType
                    Case GetType(System.Int16).Name, GetType(System.Int32).Name, GetType(System.Int64).Name, GetType(System.Decimal).Name, GetType(System.Double).Name, GetType(System.Single).Name
                        Return True
                End Select
            Case "BooleanDateTime"
                If DataType = "DateTime" Then
                    Return True
                End If
            Case "Numeric"
                Select Case DataType
                    Case GetType(System.Int16).Name, GetType(System.Int32).Name, GetType(System.Int64).Name, GetType(System.Decimal).Name, GetType(System.Double).Name, GetType(System.Single).Name
                        FilterText = " = " & FilterText
                        Return True
                End Select
            Case "DateTime"
                If DataType = "DateTime" Then
                    FilterText = " = " & FilterText
                    Return True
                End If

            Case "String"
                Select Case DataType
                    Case GetType(System.Boolean).Name
                        FilterText = " = " & FilterText
                        Return True
                End Select

            Case Else
                Return False
        End Select

        Return False
    End Function

    Shared Sub ConvertCaps(ByRef Dr As DataRow)
        For Each Dc As DataColumn In Dr.Table.Columns
            Select Case Dc.DataType.Name
                Case GetType(String).Name
                    Dr(Dc.ColumnName) = UCase(IsNull(Dr(Dc.ColumnName), ""))
            End Select
        Next
    End Sub

    Shared Sub ConvertCaps(ByRef Dt As DataTable)
        For Each Dr As DataRow In Dt.Rows
            For Each Dc As DataColumn In Dt.Columns
                Select Case Dc.DataType.Name
                    Case GetType(String).Name
                        Dr(Dc.ColumnName) = UCase(IsNull(Dr(Dc.ColumnName), ""))
                End Select
            Next
        Next
    End Sub

    Shared Function Generate_OrList(ByVal Dt As DataTable, ByVal FieldName As String, Optional ByVal Field_DefaultValue As Object = "", Optional ByVal FieldName_Or As String = "") As String
        If FieldName_Or = "" Then
            FieldName_Or = FieldName
        End If

        Dim Sb_List As New System.Text.StringBuilder
        Dim Tmp_Or As String = ""
        Dim IsStart As Boolean = False

        For Each Dr As DataRow In Dt.Rows
            Sb_List.Append(" " & Tmp_Or & " " & FieldName_Or & " = " & IsNull(Dr.Item(FieldName), Field_DefaultValue))
            If Not IsStart Then
                Tmp_Or = "Or"
                IsStart = True
            End If
        Next

        Return Sb_List.ToString

    End Function

    Public Shared Function Convert_Double(ByVal Value As String) As Double
        Dim ReturnValue As Double = 0
        Double.TryParse(Value, ReturnValue)
        Return ReturnValue
    End Function

    Friend Shared Function ParseEnum(Of T As {Structure, IComparable, IFormattable, IConvertible})(ByVal Value As [String]) As T
        Return ParseEnum(Of T)(Value, Nothing)
    End Function

    Friend Shared Function ParseEnum(Of T As {Structure, IComparable, IFormattable, IConvertible})(ByVal Value As String, ByVal DefaultValue As T) As T
        If [Enum].IsDefined(GetType(T), Value) Then
            Return DirectCast([Enum].Parse(GetType(T), Value, True), T)
        End If
        Return DefaultValue
    End Function

End Class
