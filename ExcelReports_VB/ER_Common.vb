Public Class ER_Common

    Public Enum eExcelFileFormat
        xlNormal = 0
        xlExcel5 = 1
        xlExcel97 = 2
        xlOpenXMLWorkbook = 3
        xlHtml = 4
        xlCSV = 5
        xlText = 6
        xlUnicodeCSV = 7
        xlUnicodeText = 8
    End Enum

    Public Structure Str_Parameter
        Public Name As String
        Public Value As Object
    End Structure

End Class
