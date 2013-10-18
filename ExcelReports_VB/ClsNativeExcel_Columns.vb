#Region " Imports "

Imports ExcelReports.Methods

#End Region

Friend Class ClsNativeExcel_Columns

#Region " Variables "

    Dim mListObj As New List(Of ClsNativeExcel_Columns_Obj)

#End Region

#Region " Methods "

    Public Sub Add(ByVal FieldName As String, Optional ByVal FieldDesc As String = "", Optional ByVal NumberFormat As String = "")
        If Not Me.pObj(FieldName) Is Nothing Then
            Return
        End If
        Me.mListObj.Add(New ClsNativeExcel_Columns_Obj(FieldName, FieldDesc, NumberFormat))
    End Sub

#End Region

#Region " Properties "

    Public ReadOnly Property pObj As List(Of ClsNativeExcel_Columns_Obj)
        Get
            Return Me.mListObj
        End Get
    End Property

    Public ReadOnly Property pObj(ByVal Name As String) As ClsNativeExcel_Columns_Obj
        Get
            Dim Rv As ClsNativeExcel_Columns_Obj = Nothing

            For Each Obj As ClsNativeExcel_Columns_Obj In Me.mListObj
                If Obj.mFieldName = Name Then
                    Rv = Obj
                    Exit For
                End If
            Next

            Return Rv
        End Get
    End Property

    Public ReadOnly Property pFieldName() As String()
        Get
            Dim Arr = (From O As ClsNativeExcel_Columns_Obj In Me.mListObj Select O.mFieldName).ToArray()
            Return Arr
        End Get
    End Property

    Public ReadOnly Property pFieldDesc() As String()
        Get
            Dim Arr = (From O As ClsNativeExcel_Columns_Obj In Me.mListObj Select O.mFieldDesc).ToArray()
            Return Arr
        End Get
    End Property

    Public ReadOnly Property pNumberFormat() As String()
        Get
            Dim Arr = (From O As ClsNativeExcel_Columns_Obj In Me.mListObj Select O.mNumberFormat).ToArray()
            Return Arr
        End Get
    End Property

#End Region

End Class

Friend Class ClsNativeExcel_Columns_Obj

#Region " Variables "

    Public mFieldName As String
    Public mFieldDesc As String
    Public mNumberFormat As String

#End Region

#Region " Constructor "

    Public Sub New(ByVal pFieldName As String, Optional ByVal pFieldDesc As String = "", Optional ByVal pNumberFormat As String = "")
        Me.mFieldName = pFieldName
        If pFieldDesc = "" Then
            pFieldDesc = pFieldName
        End If
        Me.mFieldDesc = pFieldDesc
        Me.mNumberFormat = pNumberFormat
    End Sub

#End Region

End Class