Public Class ReportEngine

    Public Shared Sub CreateExcelDocument(ByVal TemplateFileName As String, ByVal Parameters As List(Of ER_Common.Str_Parameter), ByRef Ds_Source As DataSet, ByRef Ds_Source_Pivot As DataSet, ByRef Ds_Source_Pivot_Desc As DataSet, ByRef Ds_Source_Pivot_Totals As DataSet, ByRef SaveFileName As String, Optional ByRef IsProtected As Boolean = False, Optional ByRef FileFormat As ER_Common.eExcelFileFormat = ER_Common.eExcelFileFormat.xlNormal)
        Dim Sp As New ClsParameters()
        Parameters.ForEach(Sub(O) Sp.Add(O.Name, O.Value))

        Dim NxlFileFormat As NativeExcel.XlFileFormat = Methods.ParseEnum(Of NativeExcel.XlFileFormat)(FileFormat.ToString())

        Methods_NativeExcel.NativeExcel_CreateExcelDocument(TemplateFileName, Sp, Ds_Source, Ds_Source_Pivot, Ds_Source_Pivot_Desc, Ds_Source_Pivot_Totals, SaveFileName, IsProtected, NxlFileFormat)
    End Sub

End Class
