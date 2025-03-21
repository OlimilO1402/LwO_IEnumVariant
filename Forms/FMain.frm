VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "FMain"
   ScaleHeight     =   3495
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton CmdWalkListColForEach2 
      Caption         =   "Walk ListCol For Each (GetEnumerator)"
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkListForEach2 
      Caption         =   "Walk List For Each (GetEnumerator)"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkCollectionForEach 
      Caption         =   "Walk Collection For Each"
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkCollectionForI 
      Caption         =   "Walk Collection For i"
      Height          =   375
      Left            =   5640
      TabIndex        =   18
      Top             =   1800
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkListColForEach 
      Caption         =   "Walk ListCol For Each (GetIEnumVariant)"
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton CmdInitListColNCollection 
      Caption         =   "Init ListCol and Collection"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton CmdInitListNArray 
      Caption         =   "Init List and Array"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkListColForI 
      Caption         =   "Walk ListCol For i"
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkRefArrayForEach 
      Caption         =   "Walk ref-Array() As Double For Each"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkRefArrayForI 
      Caption         =   "Walk ref-Array() As Double For I"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkArrayAsDoubleForEach 
      Caption         =   "Walk Array() as Double For Each"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkArrayAsDoubleForI 
      Caption         =   "Walk Array() As Double For I"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkListForEach 
      Caption         =   "Walk List For Each (GetIEnumVariant)"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton CmdWalkListForI 
      Caption         =   "Walk List For i"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label12 
      Caption         =   "..."
      Height          =   255
      Left            =   9480
      TabIndex        =   25
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "..."
      Height          =   255
      Left            =   9480
      TabIndex        =   24
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "..."
      Height          =   255
      Left            =   9480
      TabIndex        =   21
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "..."
      Height          =   255
      Left            =   9480
      TabIndex        =   20
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "..."
      Height          =   255
      Left            =   9480
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const m_u As Long = 9999999 ' 10 mio
Private myList       As List
Private myDblArr()   As Double
Private myCollection As Collection
Private myListCol    As ListCol

Private Sub CmdInitListNArray_Click()
    ' Specialising the list class during design-time as a list of Doubles
    Dim c As Long: c = m_u + 1
    MsgBox "List and array are getting filled with " & c & " elements of datatype Double (8 Byte) and filled with random values."
    Set myList = MNew.List(vbDouble)
    ReDim myDblArr(0 To m_u)
    Dim i As Long
    Dim d As Double
    For i = 0 To m_u
        d = Rnd * (i + 1)
        myList.Add d
        myDblArr(i) = d
    Next
    MsgBox "The List now contains: " & myList.Count & " elements, with a capacity of: " & myList.Capacity & vbCrLf & "List and array are now comsuming about " & FormatMByte(c + myList.Capacity, vbDouble)
End Sub

Private Sub CmdWalkListForI_Click()
    If myList Is Nothing Then
        MsgBox "Click button '" & CmdInitListNArray.Caption & "' first"
        Exit Sub
    End If
    'walk the list with "For i"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim i As Long
    For i = 0 To myList.Count - 1
        d = myList(i)
    Next
    dt = Timer - dt
    Label1 = Format(dt, "0.00000")
End Sub

Private Sub CmdWalkListForEach_Click()
    If myList Is Nothing Then
        MsgBox "Click button '" & CmdInitListNArray.Caption & "' first"
        Exit Sub
    End If
    'walk the list with "For Each"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim v
    For Each v In myList '.GetEnum
        d = v
    Next
    dt = Timer - dt
    Label2 = Format(dt, "0.00000")
End Sub

Private Sub CmdWalkListForEach2_Click()
    If myList Is Nothing Then
        MsgBox "Click button '" & CmdInitListNArray.Caption & "' first"
        Exit Sub
    End If
    'walk the list with "For Each"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim v
    For Each v In myList.GetEnumerator
        d = v
    Next
    dt = Timer - dt
    Label3 = Format(dt, "0.00000")
End Sub

Private Sub CmdWalkArrayAsDoubleForI_Click()
    If myList Is Nothing Then
        MsgBox "Click button '" & CmdInitListNArray.Caption & "' first"
        Exit Sub
    End If
    'walk an Array() As Double with "For i"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim i As Long
    For i = LBound(myDblArr) To UBound(myDblArr)
        d = myDblArr(i)
    Next
    dt = Timer - dt
    Label4 = Format(dt, "0.00000")
End Sub

Private Sub CmdWalkArrayAsDoubleForEach_Click()
    If myList Is Nothing Then
        MsgBox "Click button '" & CmdInitListNArray.Caption & "' first"
        Exit Sub
    End If
    'walk an Array() As Double with "For Each"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim v
    For Each v In myDblArr
        d = v
    Next
    dt = Timer - dt
    Label5 = Format(dt, "0.00000")
End Sub

Private Sub CmdWalkRefArrayForI_Click()
    If myList Is Nothing Then
        MsgBox "Click button '" & CmdInitListNArray.Caption & "' first"
        Exit Sub
    End If
    'walk an instance of Array() As Double with "For i"
    Dim dArr() As Double: SAPtr(ArrPtr(dArr)) = myList.SAPtr
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim i As Long
    For i = LBound(dArr) To UBound(dArr)
        d = dArr(i)
    Next
    dt = Timer - dt
    Label6 = Format(dt, "0.00000")
    ZeroSAPtr ArrPtr(dArr)
End Sub

Private Sub CmdWalkRefArrayForEach_Click()
    If myList Is Nothing Then
        MsgBox "Click button '" & CmdInitListNArray.Caption & "' first"
        Exit Sub
    End If
    'walk an instance of Array() As Double with "For Each"
    Dim dArr() As Double: SAPtr(ArrPtr(dArr)) = myList.SAPtr
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim v
    For Each v In dArr
        d = v
    Next
    dt = Timer - dt
    Label7 = Format(dt, "0.00000")
    ZeroSAPtr ArrPtr(dArr)
End Sub

' v ' ############################## ' v '    ListCol    ' v ' ############################## ' v '

Private Sub CmdInitListColNCollection_Click()
    Dim c As Long: c = m_u + 1
    MsgBox "ListCol and Collection are getting filled with " & c & " elements of datatype Double (8 Byte) and filled with random values."
    Set myCollection = New Collection
    Set myListCol = MNew.ListCol(vbDouble)
    Dim i As Long, u As Long: u = m_u
    Dim d As Double
    For i = 0 To u
        d = Rnd * (i + 1)
        myListCol.Add d
        myCollection.Add d
    Next
    MsgBox "ListCol now contains: " & myListCol.Count & " elements" & vbCrLf & "ListCol and Collection are now comsuming about " & FormatMByte(2 * c, vbVariant)
End Sub

Private Sub CmdWalkListColForI_Click()
    If myListCol Is Nothing Then
        MsgBox "Click button '" & CmdInitListColNCollection.Caption & "' first"
        Exit Sub
    End If
    'walk the collection-list with "For i"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim i As Long, c As Long: c = 100000
    MsgBox "Walking through all elements could take a while." & vbCrLf & "So we walk only " & c & " elements and extrapolating the time by the factor: " & myListCol.Count / c
    For i = 1 To c ' myListCol.Count
        d = myListCol.Item(i)
    Next
    dt = (Timer - dt) * myListCol.Count / c
    Label8 = Format(dt, "0.00000")
End Sub

Private Sub CmdWalkListColForEach_Click()
    If myListCol Is Nothing Then
        MsgBox "Click button '" & CmdInitListColNCollection.Caption & "' first"
        Exit Sub
    End If
    'walk the collection-list with "For Each"
    Dim dt As Double: dt = Timer
    Dim v, d As Double
    For Each v In myListCol
        d = v
    Next
    dt = (Timer - dt)
    Label9 = Format(dt, "0.00000")
End Sub

Private Sub CmdWalkListColForEach2_Click()
    If myListCol Is Nothing Then
        MsgBox "Click button '" & CmdInitListColNCollection.Caption & "' first"
        Exit Sub
    End If
    'walk the collection-list with "For Each"
    Dim dt As Double: dt = Timer
    Dim v, d As Double
    For Each v In myListCol.GetEnumerator
        d = v
    Next
    dt = Timer - dt
    Label10 = Format(dt, "0.00000")
End Sub

Private Sub CmdWalkCollectionForI_Click()
    If myCollection Is Nothing Then
        MsgBox "Click button '" & CmdInitListColNCollection.Caption & "' first"
        Exit Sub
    End If
    'walk the collection with "For i"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim i As Long, c As Long: c = 100000
    MsgBox "Walking through all elements could take a while." & vbCrLf & "So we walk only " & c & " elements and extrapolating the time by the factor: " & (m_u + 1) / c
    For i = 1 To c
        d = myCollection.Item(i)
    Next
    dt = (Timer - dt) * myCollection.Count / c
    Label11 = Format(dt, "0.00000")
End Sub

Private Sub CmdWalkCollectionForEach_Click()
    If myListCol Is Nothing Then
        MsgBox "Click button '" & CmdInitListColNCollection.Caption & "' first"
        Exit Sub
    End If
    'walk the collection with "For Each"
    Dim dt As Double: dt = Timer
    Dim v, d As Double
    For Each v In myCollection
        d = v
    Next
    dt = Timer - dt
    Label12 = Format(dt, "0.00000")
End Sub
