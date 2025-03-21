Attribute VB_Name = "MEnumVariant"
Option Explicit
'A lightweight object for an enumerator-objekt implementing IEnumVariant
'GUID of IEnumVariant: 00020404-0000-0000-C000-000000000046
Private Type VBGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data5(0 To 7) As Byte
End Type
Private Const sIID_IUnknown     As String = "{00000000-0000-0000-C000-000000000046}"
Private Const sIID_IDispatch    As String = "{00020400-0000-0000-C000-000000000046}"
Private Const sIID_IEnumVariant As String = "{00020404-0000-0000-C000-000000000046}"
'spot the difference:  ______________________________^_____________________________
Public IID_IUnknown     As VBGuid
Public IID_IDispatch    As VBGuid
Public IID_IEnumVariant As VBGuid

Private Type TByteHiLo
    Lo As Byte
    Hi As Byte
End Type
Private Type TInteger
    Value As Integer
End Type

'a VTable contains pointers to the functions of a class
Private Type TEnumVariantVTable
    VTable(0 To 6) As LongPtr
End Type

'Differentiation between objects and simple data types using two different Next functions
Private EnumObjVTable As TEnumVariantVTable
Private m_pVTableObj  As LongPtr

'for every primitive datatype we could create a separate EnumVariant-objekt
'each with a separate Next-function
Private EnumVarVTable As TEnumVariantVTable
Private m_pVTableVar  As LongPtr

'
Public Type TEnumVariant
    pVTable As LongPtr   'the first element in an object is always a pointer to it's VTable
    'Owner   As Object
    refCnt  As Long      'the reference counter
    Array   As Variant   'this Variant contains a pointer to an Array of any type
    vt      As VbVarType
    Count   As Long      'the amount of elements in the Array to loop through
    Index   As Long      'the Index-counter, Index to the next element
End Type

Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER     As Long = &H80004003

Private Const S_OK    As Long = &H0&
Private Const S_FALSE As Long = &H1&

#If VBA7 Then
    'https://learn.microsoft.com/de-de/windows/win32/api/combaseapi/nf-combaseapi-clsidfromstring
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal pString As LongPtr, ByRef pCLSID As Any) As Long
    'https://learn.microsoft.com/de-de/windows/win32/api/combaseapi/nf-combaseapi-stringfromguid2
    Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef pGuid As Any, ByVal lpsz As LongPtr, ByVal cchMax As Long) As Long
#Else
    Private Declare Function CLSIDFromString Lib "ole32" (ByVal pString As LongPtr, ByRef pCLSID As Any) As Long
    Private Declare Function StringFromGUID2 Lib "ole32" (ByRef pGuid As Any, ByVal lpsz As LongPtr, ByVal cchMax As Long) As Long
#End If

Public Sub InitEnumVariantVTable()
    'Initialising the function pointers of the IEnumVariant lightweight class
    'call it only once per project, e.g. in Sub Main
    IID_IUnknown = New_VBGuidS(sIID_IUnknown)
    IID_IDispatch = New_VBGuidS(sIID_IDispatch)
    IID_IEnumVariant = New_VBGuidS(sIID_IEnumVariant)
    
    'In VB a Sub is simply also a function, because a HResult will always be returned
    With EnumVarVTable
        .VTable(0) = FncPtr(AddressOf FncQueryInterface) 'FncPtr is defined in MPtr
        .VTable(1) = FncPtr(AddressOf FncAddRef)
        .VTable(2) = FncPtr(AddressOf FncRelease)
        .VTable(4) = FncPtr(AddressOf SubSkip)
        .VTable(5) = FncPtr(AddressOf SubReset)
        .VTable(6) = FncPtr(AddressOf FncClone)
    End With
    EnumObjVTable = EnumVarVTable
    EnumObjVTable.VTable(3) = FncPtr(AddressOf FncNextObj) 'for object-datatypes
    EnumVarVTable.VTable(3) = FncPtr(AddressOf FncNextVar) 'for primitive datatypes
    '...
    m_pVTableVar = VarPtr(EnumVarVTable)
    m_pVTableObj = VarPtr(EnumObjVTable)
    
End Sub

' v ' ############################## ' v '    VBGuid    ' v ' ############################## ' v '
Public Function New_VBGuidS(ByVal sIID As String) As VBGuid
    VBGuid_Parse New_VBGuidS, sIID
End Function
Public Sub VBGuid_Parse(this As VBGuid, ByVal sIID As String)
    Dim hr As Long: hr = CLSIDFromString(StrPtr(sIID), this)
    If hr <> 0 Then MsgBox "Error creating guid from string: '" & sIID & "'"
End Sub
Public Function VBGuid_ToStr(this As VBGuid) As String
    VBGuid_ToStr = String(40, 0)
    Dim hr As Long: hr = StringFromGUID2(this, StrPtr(VBGuid_ToStr), 40)
    VBGuid_ToStr = MString.Trim0(VBGuid_ToStr)
End Function
Public Function VBGuid_Equals(this As VBGuid, other As VBGuid) As Boolean
    With other
        If .Data1 <> this.Data1 Then Exit Function
        If .Data2 <> this.Data2 Then Exit Function
        If .Data3 <> this.Data3 Then Exit Function
        If .Data5(0) <> this.Data5(0) Then Exit Function
        If .Data5(1) <> this.Data5(1) Then Exit Function
        If .Data5(2) <> this.Data5(2) Then Exit Function
        If .Data5(3) <> this.Data5(3) Then Exit Function
        If .Data5(4) <> this.Data5(4) Then Exit Function
        If .Data5(5) <> this.Data5(5) Then Exit Function
        If .Data5(6) <> this.Data5(6) Then Exit Function
        If .Data5(7) <> this.Data5(7) Then Exit Function
    End With
    VBGuid_Equals = True
End Function
' ^ ' ############################## ' ^ '    VBGuid    ' ^ ' ############################## ' ^ '

'Owner As Object, '_'
Public Function New_Enum(Me_ As TEnumVariant, _
                         Arr As Variant, _
                         ByVal vt As VbVarType, _
                         ByVal Count As Long) As IUnknown
    With Me_
        'we could use separate next-function for each datatype
        .pVTable = IIf(vt = vbObject, m_pVTableObj, m_pVTableVar)
        'Set .Owner = Owner
        'copy the pointer to the Array from the variant to the Variant completely
        RtlMoveMemory .Array, Arr, MPtr.SizeOf_Variant
        .Count = Count
        .Index = 0
        .refCnt = 4
    End With
    
    'now bring the object to life
    RtlMoveMemory New_Enum, VarPtr(Me_), SizeOf_LongPtr
End Function

' v ' ############################## ' v '    IUnkown    ' v ' ############################## ' v '
Private Function FncQueryInterface(Me_ As TEnumVariant, riid As VBGuid, ByVal ppvObject As LongPtr) As Long
    'Debug.Print "QueryInterface"
    'COMis "asking" whether it is an IEnumVariant-Objekt
    If ppvObject = 0 Then
        FncQueryInterface = E_POINTER
        Exit Function
    End If
    'check for IEnumVariant
    If Not VBGuid_Equals(IID_IEnumVariant, riid) Then
        'check for IDispatch
        If Not VBGuid_Equals(IID_IDispatch, riid) Then
            'check for IUnknown
            If Not VBGuid_Equals(IID_IUnknown, riid) Then
                FncQueryInterface = E_NOINTERFACE
                Exit Function
            End If
        End If
    End If
    DeRef(ppvObject) = VarPtr(Me_)
End Function

Private Function FncAddRef(Me_ As TEnumVariant) As Long
    Me_.refCnt = Me_.refCnt + 1
    FncAddRef = Me_.refCnt
End Function

Private Function FncRelease(Me_ As TEnumVariant) As Long
    Me_.refCnt = Me_.refCnt - 1
    FncRelease = Me_.refCnt
    If Me_.refCnt = 0 Then RtlZeroMemory Me_.Array, MPtr.SizeOf_Variant
End Function
' ^ ' ############################## ' ^ '    IUnkown    ' ^ ' ############################## ' ^ '

Private Function FncNextObj(Me_ As TEnumVariant, _
                            ByVal celt As Long, _
                            rgvar, _
                            pceltFetched As Long) As Long
    ' Dies ist die wichtigste Funktion von IEnumVariant.
    ' Über Count wird entschieden wann der Vorgang abgebrochen wird.
    If Me_.Index = Me_.Count Then FncNextObj = S_FALSE: Exit Function
    Set rgvar = Me_.Array(Me_.Index)
    Me_.Index = Me_.Index + 1
End Function

Private Function FncNextVar(Me_ As TEnumVariant, _
                            ByVal celt As Long, rgvar, pceltFetched As Long) As Long
    'Dies ist die wichtigste Funktion von IEnumVariant.
    ' Über Count wird entschieden wann der Vorgang abgebrochen wird.
    If Me_.Index = Me_.Count Then FncNextVar = S_FALSE: Exit Function
    rgvar = Me_.Array(Me_.Index)
    Me_.Index = Me_.Index + 1
End Function

Private Function SubSkip(Me_ As TEnumVariant, ByVal celt As Long) As Long
    'hier nur Dummy-Funktion, wird nicht verwendet
End Function

Private Function SubReset(Me_ As TEnumVariant) As Long
    'hier nur Dummy-Funktion, wird nicht verwendet
End Function

Private Function FncClone(Me_ As TEnumVariant, retOther As TEnumVariant) As Long
    'hier nur Dummy-Funktion, wird nicht verwendet
End Function
