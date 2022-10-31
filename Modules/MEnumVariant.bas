Attribute VB_Name = "MEnumVariant"
Option Explicit
'Ein Lightweight Object für ein Enumerator-Objekt das IEnumVariant implementiert
'zur Info GUID von IEnumVariant: 00020404-0000-0000-C000-000000000046
Private Type VBGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data5(0 To 7) As Byte
End Type

'Eine VTable enthält Zeiger auf die Funktionen einer Klasse
Private Type TEnumVariantVTable
    VTable(0 To 6) As Long
End Type

'Unterscheidung in Objekte und einfache Datentypen durch zwei verschiedene Next-Funktionen
Private EnumObjVTable As TEnumVariantVTable
Private m_pVTableObj As Long

'man könnte auch für jeden einfachen Datentyp ein eigenes EnumVariant-Objekt erstellen mit
'jeweils unterschiedlichen Next-Funktionen
Private EnumVarVTable As TEnumVariantVTable
Private m_pVTableVar As Long

'
Public Type TEnumVariant
    pVTable As Long      'erstes Element in einem Objekt ist immer ein Zeigr auf die VTable
    refCnt  As Long      'der Referenzzähler
    Array   As Variant   'der Variant enthält einen Zeiger auf ein Array beliebigen Typs
    Count   As Long      'die Anzahl der abzulaufenden Elemente im Array
    Index   As Long      'der Indexzähler, Index auf das nächste Element
End Type
Private Const S_OK = &H0&
Private Const S_FALSE = &H1&

Public Sub InitEnumVariantVTable()
    'Initialisierung der Funktionszeiger der IEnumVariant Lightweight Klasse
    'soll im Projekt nur einmal aufgerufen werden, z.B. von Sub Main
    'In VB ist eine Sub eigentlich auch eine Funktion, da immer ein HResult
    'zurückgegeben wird.
    With EnumVarVTable
        .VTable(0) = FncPtr(AddressOf FncQueryInterface)
        .VTable(1) = FncPtr(AddressOf SubAddRef)
        .VTable(2) = FncPtr(AddressOf SubRelease)
        .VTable(4) = FncPtr(AddressOf SubSkip)
        .VTable(5) = FncPtr(AddressOf SubReset)
        .VTable(6) = FncPtr(AddressOf FncClone)
    End With
    EnumObjVTable = EnumVarVTable
    EnumObjVTable.VTable(3) = FncPtr(AddressOf FncNextObj) 'für Objekttypen
    EnumVarVTable.VTable(3) = FncPtr(AddressOf FncNextVar) 'für einfache Datentypen
    '...
    m_pVTableVar = VarPtr(EnumVarVTable)
    m_pVTableObj = VarPtr(EnumObjVTable)
    
End Sub
'Private Function FncPtr(ByVal pFnc As Long) As Long
'    FncPtr = pFnc
'End Function

Public Function New_Enum(Me_ As TEnumVariant, _
                         Arr As Variant, _
                         ByVal vt As VbVarType, _
                         ByVal Count As Long) As IUnknown
    With Me_
        'man könnte auch für jeden Datentyp eine eigene Next-Prozedur verwenden
        .pVTable = IIf(vt = vbObject, m_pVTableObj, m_pVTableVar)
        
        'Den Zeiger auf das Array komplett aus dem Variant in den Variant kopieren
        RtlMoveMemory .Array, Arr, 16
        .Count = Count
        .Index = 0
        .refCnt = 2
    End With
    
    'das Objekt zum Leben erwecken
    RtlMoveMemory New_Enum, VarPtr(Me_), 4
End Function

Private Function FncQueryInterface(Me_ As TEnumVariant, riid As VBGuid, pvObj As Long) As Long
    
    ' Hier frägt VB das Objekt, ob es sich "wirklich" um ein IEnumVariant-Objekt handelt.
    ' Man braucht diese Abfrage eigentlich nicht, da wir ja wissen welches Objekt es ist.
    ' Es soll hier nur exemplarisch gezeigt werden wie eine solche Abfrage aussehen kann.
    ' Es muss aber in jedem Fall in pvObj ein Zeiger auf das Objekt zurückgegeben werden.
    
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
    ' hier wird eine Referenz hinzugefügt
    With Me_
        .refCnt = .refCnt + 1
    End With
End Function

Private Function SubRelease(Me_ As TEnumVariant) As Long
    ' hier wird eine Referenz abgezogen
    ' wird diese Funktion wiederholt aufgerufen, solange bis refCounter = 0
    ' dann den Zeiger auf das Array im Variant Array wieder löschen
    With Me_
        .refCnt = .refCnt - 1
        If .refCnt = 0 Then RtlZeroMemory .Array, 16
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
