Attribute VB_Name = "MNew"
Option Explicit

Public Sub Main()
    MEnumVariant.InitEnumVariantVTable
    FMain.Show
End Sub

Public Function List(ByVal Of_Type As VbVarType) As List
    Set List = New List: List.New_ Of_Type
End Function

Public Function ListCol(ByVal Of_Type As VbVarType, Optional ByVal UseHashing As Boolean = False) As ListCol
    Set ListCol = New ListCol: ListCol.New_ Of_Type, UseHashing
End Function

