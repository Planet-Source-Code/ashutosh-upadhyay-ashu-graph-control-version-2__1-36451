VERSION 5.00
Object = "{828AC00A-8DB0-11D6-A6D1-0050BAA907A1}#11.0#0"; "AshuGraphControl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6804
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7704
   LinkTopic       =   "Form1"
   ScaleHeight     =   6804
   ScaleWidth      =   7704
   StartUpPosition =   3  'Windows Default
   Begin GraphControl.AshuGraphControl AshuGraphControl1 
      Height          =   4332
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   6372
      _ExtentX        =   11240
      _ExtentY        =   7641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   361
      ScaleWidth      =   531
      XUnitToPixels   =   100
      YUnitToPixels   =   100
      XGap            =   0.5
      YGap            =   0.5
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove My Curve"
      Height          =   492
      Left            =   3360
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Draw derivative of sine_curve"
      Height          =   492
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   2292
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ReadGraph"
      Height          =   492
      Left            =   5040
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save My Graph "
      Height          =   612
      Left            =   5160
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Draw My Graph"
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   2292
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sine_curve As Long
Dim index As Long

Private Sub AshuGraphControl1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "mouse down"
End Sub


Private Sub Command1_Click()
AshuGraphControl1.RemoveGraph (index)
AshuGraphControl1.Invalidate
Command6.Visible = True
End Sub

Private Sub Command2_Click()
Dim p As Single
Dim q As Single
Dim i As Single

index = AshuGraphControl1.GetFreeGraph
AshuGraphControl1.SetColor RGB(0, 200, 0), index

For i = 1 To 1000 Step 0.05
q = (5 * i)
'''' don't give negative data
p = (100 * CSng(Cos(i)) + 100)
Call AshuGraphControl1.AddData(q, p, index)
Next i
Call AshuGraphControl1.AddData(-100, -200, index)
Call AshuGraphControl1.AddData(100, 500, index)
''''DrawStyle 0 means draw points as pixels
''''drawstyle 1 means join points using lines (line clipping doesn't support)
''''drawstyle 2 means draw points as square
''''drawstyle 3 means draw points as square and join them by lines
AshuGraphControl1.SetDrawStyle index, 2
AshuGraphControl1.Invalidate
Command5.Visible = True
End Sub





Private Sub Command5_Click()
Call AshuGraphControl1.SaveGraphInTextFile(index, "data.data")
Command1.Visible = True
End Sub

Private Sub Command6_Click()
Call AshuGraphControl1.ReadGraphFromTextFile("data.data", False) ' returns index of graph && second parameter is False means don't reset the settings of graph
AshuGraphControl1.Invalidate
End Sub

Private Sub Command7_Click()
AshuGraphControl1.DrawDerivativeForGraph sine_curve
AshuGraphControl1.Invalidate
End Sub

Private Sub Form_Load()

AshuGraphControl1.InitializeMe
sine_curve = AshuGraphControl1.GetFreeGraph
AshuGraphControl1.SetColor RGB(255, 0, 0), sine_curve
AshuGraphControl1.SetName "SineCurve", sine_curve
AshuGraphControl1.SineCurveFill (sine_curve)
AshuGraphControl1.Invalidate
Call AshuGraphControl1.PrintSettings(10, 10, "ashu", "X --->", "Y ^ | | |", 1, 200, 9, -1)
''' zoom factor will only work if your pinter support zooming printing
''''paper size = 9 means  A4 , if = 10 means A4 small, see msdn help on "printer" in vb editor
''' printquality = - 1 means draft(poor)

End Sub
