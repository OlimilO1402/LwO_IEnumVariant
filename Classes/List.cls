VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public Class List
'this is a prototype of a generic List, the datatype of the array will be defined during design-time
Private m_Array
Private m_Count As Long
Private m_vType As VbVarType
Private m_EnumV As TEnumVariant

Friend Sub New_(Of_Type As VbVarType)
    m_vType = Of_Type
    Clear
End Sub

Public Sub Add(Item As Variant)
    Me.Capacity = m_Count
    If m_vType = vbObject Then
        Set m_Array(m_Count) = Item
    Else
        m_Array(m_Count) = Item
    End If
    m_Count = m_Count + 1
End Sub

Public Property Get Capacity() As Long
    Capacity = UBound(m_Array) + 1
End Property

Public Property Let Capacity(Value As Long)
    Dim cap As Long: cap = UBound(m_Array) + 1
    If cap <= Value Then
        cap = cap * 2
        If cap < Value Then cap = Value
        ReDim Preserve m_Array(0 To cap - 1)
    End If
End Property
Public Property Get SAPtr() As LongPtr
    SAPtr = VArrPtr(m_Array)
End Property
Public Property Get DataPtr() As LongPtr
    DataPtr = VarPtr(m_Array(0))
End Property
Public Sub Clear()
    Dim u As Long
    u = 3
    Select Case m_vType
    Case vbInteger  '2
        ReDim m_Array(0 To u) As Integer
    Case vbLong     '3
        ReDim m_Array(0 To u) As Long
    Case vbSingle   '4
        ReDim m_Array(0 To u) As Single
    Case vbDouble   '5
        ReDim m_Array(0 To u) As Double
    Case vbCurrency '6
        ReDim m_Array(0 To u) As Currency
    Case vbDate     '7
        ReDim m_Array(0 To u) As Date
    Case vbString   '8
        ReDim m_Array(0 To u) As String
    Case vbObject   '9
        ReDim m_Array(0 To u) As Object
    Case vbBoolean  '11
        ReDim m_Array(0 To u) As Boolean
    Case vbVariant, vbDecimal '12, 14
        ReDim m_Array(0 To u) As Variant
    Case vbByte    '17
        ReDim m_Array(0 To u) As Byte
    End Select
    m_Count = 0
End Sub

Public Property Get Count()
    Count = m_Count
End Property

Public Function GetEnum() As IUnknown
Attribute GetEnum.VB_UserMemId = -4
    ' Prozedur-Id = -4
    Set GetEnum = New_Enum(m_EnumV, m_Array, m_vType, m_Count)
End Function

Public Property Get Item(Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    'Prozedur-Id = (Voreinstellung) 'Default
    If m_vType = vbObject Then
        Set Item = m_Array(Index)
    Else
        Item = m_Array(Index)
    End If
End Property

Public Property Let Item(Index As Long, Value As Variant)
    m_Array(Index) = Value
End Property

Public Property Set Item(Index As Long, Value As Variant)
    Set m_Array(Index) = Value
End Property

