VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Timer Example"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Waitable Timer"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   4215
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3480
         TabIndex        =   14
         Text            =   "100"
         Top             =   285
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Wait"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Interval"
         Height          =   195
         Left            =   2880
         TabIndex        =   15
         Top             =   330
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Event Timer"
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   4215
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3480
         TabIndex        =   10
         Text            =   "100"
         Top             =   1605
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Interval"
         Height          =   195
         Left            =   2880
         TabIndex        =   11
         Top             =   1650
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Event Interface Timer"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3480
         TabIndex        =   1
         Text            =   "100"
         Top             =   1605
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Interval"
         Height          =   195
         Left            =   2880
         TabIndex        =   5
         Top             =   1650
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements iTimer
Private mobj_timer As clsTimer

Private WithEvents mobj_timer2 As clsTimer

Private Sub Command1_Click()
  List1.Clear
  mobj_timer.Interval = CInt(Text1.Text)
  mobj_timer.StartTimer
  Command1.Enabled = False
  Command2.Enabled = True
End Sub

Private Sub Command2_Click()
  mobj_timer.StopTimer
  Command1.Enabled = True
  Command2.Enabled = False
End Sub

Private Sub Command3_Click()
  Dim obj_timer As clsTimer
  Set obj_timer = New clsTimer
  Command3.Enabled = False
  obj_timer.Wait CInt(Text3.Text)
  MsgBox "Waitable Timer Complete"
  Command3.Enabled = True
  Set obj_timer = Nothing
End Sub

Private Sub Command4_Click()
  List2.Clear
  mobj_timer2.Interval = CInt(Text2.Text)
  mobj_timer2.StartTimer
  Command4.Enabled = False
  Command5.Enabled = True
End Sub

Private Sub Command5_Click()
  mobj_timer2.StopTimer
  Command4.Enabled = True
  Command5.Enabled = False
End Sub

Private Sub Form_Load()
  Set mobj_timer = New clsTimer
  Set mobj_timer.Interface = Me
  mobj_timer.Interval = 100
  
  Set mobj_timer2 = New clsTimer
  mobj_timer2.Interval = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mobj_timer.StopTimer
  Set mobj_timer = Nothing
  mobj_timer2.StopTimer
  Set mobj_timer2 = Nothing
End Sub

Private Sub iTimer_OnTime(ByVal int_ticks As Integer, ByVal dwTime As Long)
  List1.AddItem "TICKS:" & int_ticks & ",DWTIME:" & dwTime
  List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub mobj_timer2_OnTime(ByVal int_ticks As Integer, ByVal dwTime As Long)
  List2.AddItem "TICKS:" & int_ticks & ",DWTIME:" & dwTime
  List2.ListIndex = List2.ListCount - 1
End Sub
