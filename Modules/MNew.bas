Attribute VB_Name = "MNew"
Option Explicit

Public Sub Main()
    MEnumVariant.InitEnumVariantVTable
    FMain.Show
End Sub

Public Function List(Of_Type As VbVarType) As List
    Set List = New List: List.New_ Of_Type
End Function
