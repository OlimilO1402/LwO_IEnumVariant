VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "FMain"
   ScaleHeight     =   2895
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton CmdWalkRefArrayForEach 
      Caption         =   "Walk ref-Array() As Double For Each"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton CmdWalkRefArrayForI 
      Caption         =   "Walk ref-Array() As Double For I"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CommandButton CmdWalkArrayAsDoubleForEach 
      Caption         =   "Walk Array() as Double For Each"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton CmdWalkArrayAsDoubleForI 
      Caption         =   "Walk Array() As Double For I"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton CmdWalkListForEach 
      Caption         =   "Walk List For Each"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton CmdWalkListForI 
      Caption         =   "Walk List For i"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "..."
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "..."
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "..."
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "..."
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "..."
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "..."
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private myList As List
Private Const m_u As Long = 9999999 ' 10 mio elements
Private myDblArr(0 To m_u) As Double

Private Sub Form_Load()
    ' Specialising the list class during design-time as a list of Doubles
    Set myList = MNew.List(vbDouble)
    Dim i As Long
    Dim d As Double
    For i = 0 To m_u
        d = Rnd * (i + 1)
        myList.Add d
        myDblArr(i) = d
    Next
    MsgBox "Count: " & myList.Count & " Capacity: " & myList.Capacity
End Sub

Private Sub CmdWalkListForI_Click()
    'walk the list with "For i"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim i As Long
    For i = 0 To myList.Count - 1
        d = myList(i)
    Next
    dt = Timer - dt
    Label1 = Format(dt, "0.000")
End Sub

Private Sub CmdWalkListForEach_Click()
    'walk the list with "For Each"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim v
    For Each v In myList '.GetEnum
        d = v
    Next
    dt = Timer - dt
    Label2 = Format(dt, "0.000")
End Sub

Private Sub CmdWalkArrayAsDoubleForI_Click()
    'walk an Array() As Double with "For i"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim i As Long
    For i = LBound(myDblArr) To UBound(myDblArr)
        d = myDblArr(i)
    Next
    dt = Timer - dt
    Label3 = Format(dt, "0.000")
End Sub

Private Sub CmdWalkArrayAsDoubleForEach_Click()
    'walk an Array() As Double with "For Each"
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim v
    For Each v In myDblArr
        d = v
    Next
    dt = Timer - dt
    Label4 = Format(dt, "0.000")
End Sub

Private Sub CmdWalkRefArrayForI_Click()
    'walk an instance of Array() As Double with "For i"
    Dim dArr() As Double: SAPtr(ArrPtr(dArr)) = myList.SAPtr
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim i As Long
    For i = LBound(dArr) To UBound(dArr)
        d = dArr(i)
    Next
    dt = Timer - dt
    Label5 = Format(dt, "0.000")
    ZeroSAPtr ArrPtr(dArr)
End Sub

Private Sub CmdWalkRefArrayForEach_Click()
    'walk an instance of Array() As Double with "For Each"
    Dim dArr() As Double: SAPtr(ArrPtr(dArr)) = myList.SAPtr
    Dim dt As Double: dt = Timer
    Dim d As Double
    Dim v
    For Each v In dArr
        d = v
    Next
    dt = Timer - dt
    Label6 = Format(dt, "0.000")
    ZeroSAPtr ArrPtr(dArr)
End Sub


