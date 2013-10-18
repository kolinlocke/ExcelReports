Friend Class ClsParameters

#Region " Variables "

    Dim mListObj As New List(Of Str_Parameter)

    Public Structure Str_Parameter
        Dim Name As String
        Dim Value As Object
        Public Sub New(ByVal pName As String, ByVal pValue As Object)
            Name = pName
            Value = pValue
        End Sub
    End Structure

#End Region

#Region " Methods "

    Public Sub Add(ByVal Name As String, ByVal Value As Object)
        Me.mListObj.Add(New Str_Parameter(Name, Value))
    End Sub

#End Region

#Region " Properties "

    Public ReadOnly Property pName As String()
        Get
            Dim List_Obj As New List(Of String)
            For Each Obj As Str_Parameter In Me.mListObj
                List_Obj.Add(Obj.Name)
            Next

            Return List_Obj.ToArray
        End Get
    End Property

    Public ReadOnly Property pValue As Object()
        Get
            Dim List_Obj As New List(Of Object)
            For Each Obj As Str_Parameter In Me.mListObj
                List_Obj.Add(Obj.Value)
            Next

            Return List_Obj.ToArray
        End Get
    End Property

#End Region

End Class
