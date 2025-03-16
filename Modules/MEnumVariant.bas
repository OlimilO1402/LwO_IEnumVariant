Attribute VB_Name = "MEnumVariant"
Option Explicit
'A lightweight object for an enumerator-objekt implementing IEnumVariant
'GUID of IEnumVariant: 00020404-0000-0000-C000-000000000046
Private Type VBGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data5(0 To 7) As Byte
End Type

#If Win64 Then
    Private Const IID_IEnumVariant As String = "{00020400-0000-0000-C000-000000000046}"
#Else
    Private Const IID_IEnumVariant As String = "{00020404-0000-0000-C000-000000000046}"
    'spot the difference: ________________________________^_____________________________
#End If

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
    Owner   As Object
    refCnt  As Long      'the reference counter
    Array   As Variant   'this Variant contains a pointer to an Array of any type
    Count   As Long      'the amount of elements in the Array to loop through
    Index   As Long      'the Index-counter, Index to the next element
End Type

Private Const S_OK    As Long = &H0&
Private Const S_FALSE As Long = &H1&

Public Sub InitEnumVariantVTable()
    'Initialising the function pointers of the IEnumVariant lightweight class
    'you should call it only once per project, e.g. in Sub Main
    'In VB a Sub is simply also a function, because a HResult will always be returned
    With EnumVarVTable
        .VTable(0) = FncPtr(AddressOf FncQueryInterface) 'FncPtr is defined in MPtr
        .VTable(1) = FncPtr(AddressOf SubAddRef)
        .VTable(2) = FncPtr(AddressOf SubRelease)
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

Public Function New_Enum(Me_ As TEnumVariant, _
                         Owner As Object, _
                         arr As Variant, _
                         ByVal vt As VbVarType, _
                         ByVal Count As Long) As IUnknown
    With Me_
        'we could use separate next-function for each datatype
        .pVTable = IIf(vt = vbObject, m_pVTableObj, m_pVTableVar)
        Set .Owner = Owner
        'copy the pointer to the Array from the variant to the Variant completely
        RtlMoveMemory .Array, arr, MPtr.SizeOf_Variant
        .Count = Count
        .Index = 0
        .refCnt = 4
    End With
    
    'now bring the object to life
    RtlMoveMemory New_Enum, VarPtr(Me_), SizeOf_LongPtr
End Function

Private Function FncQueryInterface(Me_ As TEnumVariant, riid As VBGUID, pvObj As LongPtr) As Long
    
    ' Hier frägt VB das Objekt, ob es sich "wirklich" um ein IEnumVariant-Objekt handelt.
    ' now VB is aksing the Object whether it is really an IEnumVariant-object
    
    ' Man braucht diese Abfrage eigentlich nicht, da wir ja wissen welches Objekt es ist.
    ' this question is not needed anyway, because we know which object it is
    
    ' Es soll hier nur exemplarisch gezeigt werden wie eine solche Abfrage aussehen kann.
    ' This is just an exampe of how such a Query could look like
    
    ' Es muss aber in jedem Fall in pvObj ein Zeiger auf das Objekt zurückgegeben werden.
    ' in every case you have to return
    
    With riid
        If .Data1 = &H20404 And _
           .Data2 = 0 And _
           .Data3 = 0 And _
           .Data5(0) = &HC0 And _
           .Data5(7) = &H46 Then
            pvObj = VarPtr(Me_)
        End If
    End With
    
    ' kann man auch weglassen da S_OK sowieso nur 0 ist
    FncQueryInterface = S_OK ' ja wir haben das Interface

End Function

Private Function SubAddRef(Me_ As TEnumVariant) As Long
    ' here a reference will be added
    With Me_
        .refCnt = .refCnt + 1
    End With
End Function

Private Function SubRelease(Me_ As TEnumVariant) As Long
    ' hier a reference will be removed
    ' wird diese Funktion wiederholt aufgerufen, solange bis refCounter = 0
    ' dann den Zeiger auf das Array im Variant Array wieder löschen
    With Me_
        .refCnt = .refCnt - 1
        If .refCnt = 0 Then RtlZeroMemory .Array, MPtr.SizeOf_Variant
    End With
End Function

Private Function FncNextObj(Me_ As TEnumVariant, _
                            ByVal celt As Long, _
                            rgvar, _
                            pceltFetched As Long) As Long
    ' Dies ist die wichtigste Funktion von IEnumVariant.
    ' Über Count wird entschieden wann der Vorgang abgebrochen wird.
    With Me_
        If .Index = .Count Then FncNextObj = S_FALSE: Exit Function
        Set rgvar = .Array(.Index)
        .Index = .Index + 1
    End With
End Function

Private Function FncNextVar(Me_ As TEnumVariant, _
                            ByVal celt As Long, rgvar, pceltFetched As Long) As Long
    'Dies ist die wichtigste Funktion von IEnumVariant.
    ' Über Count wird entschieden wann der Vorgang abgebrochen wird.
    With Me_
        If .Index = .Count Then FncNextVar = S_FALSE: Exit Function
        rgvar = .Array(.Index)
        .Index = .Index + 1
    End With
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
