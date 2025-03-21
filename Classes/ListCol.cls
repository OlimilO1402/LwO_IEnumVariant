VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'this is a List class based solely on a Collection object
Private m_Collection As Collection
Private m_bIsHashed  As Boolean
Private m_vType      As VbVarType

Friend Sub New_(Of_Type As VbVarType, UseHashing As Boolean)
    m_vType = Of_Type:  m_bIsHashed = UseHashing
    Clear
End Sub

'Copy to MNew
'Public Function ListCol(ByVal Of_Type As VbVarType, ByVal UseHashing As Boolean) As ListCol
'    Set ListCol = New ListCol: ListCol.New_ Of_Type, UseHashing
'End Function

Public Sub Add(Item As Variant, Optional Key As String)
    If m_bIsHashed Then m_Collection.Add Item, Key Else m_Collection.Add Item
End Sub

Public Sub Clear()
    Set m_Collection = New Collection
End Sub

Public Function Contains(ValueOrKey) As Boolean
    If m_bIsHashed Then
        Contains = MPtr.Col_Contains(m_Collection, CStr(ValueOrKey))
        Exit Function
    End If
    If m_vType = vbObject Then
        'nyi
    Else
        'nyi
    End If
End Function

Public Property Get Count()
    Count = m_Collection.Count
End Property

Public Function GetIEnumVariant() As IUnknown
Attribute GetIEnumVariant.VB_UserMemId = -4
'Attribute GetEnum.VB_UserMemId = -4
    ' Prozedur-Id = -4
    Set GetIEnumVariant = m_Collection.[_NewEnum]
End Function

Public Function GetEnumerator()
    Set GetEnumerator = m_Collection
End Function

Public Property Get IsHashed() As Boolean
    IsHashed = m_bIsHashed
End Property

Public Property Get Item(IndexOrKey) As Variant
Attribute Item.VB_UserMemId = 0
    'Prozedur-Id = (Voreinstellung) 'Default
    If m_vType = vbObject Then Set Item = m_Collection.Item(IndexOrKey) Else Item = m_Collection.Item(IndexOrKey)
End Property

Public Property Let Item(IndexOrKey, Value As Variant)
    If m_bIsHashed Then
        Dim Key As String: Key = CStr(IndexOrKey)
        If Me.Contains(Key) Then m_Collection.Remove Key
        m_Collection.Add Value, Key
    Else
        Dim i As Long: i = CLng(IndexOrKey)
        m_Collection.Add Value, , i
        m_Collection.Remove i
    End If
End Property

Public Property Set Item(IndexOrKey, Value As Variant)
    If m_bIsHashed Then
        Dim Key As String: Key = CStr(IndexOrKey)
        If Me.Contains(Key) Then m_Collection.Remove Key
        m_Collection.Add Value, Key
    Else
        Dim i As Long: i = CLng(IndexOrKey)
        m_Collection.Add Value, , i
        m_Collection.Remove i
    End If
End Property

