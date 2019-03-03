Sub ColumnSplitter()
' Splits Category and Subcategory column into 2 separate columns

    Dim r As Integer
    Dim CatSub() As String

    For r = 2 To 4115
        CatSub = Split(Cells(r, 14).Value, "/")
        Cells(r, 17).Value = CatSub(0)
        Cells(r, 18).Value = CatSub(1)
    Next r

End Sub
