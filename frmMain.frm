VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   Caption         =   "绘制函数图像"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   8775
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picTool 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   8775
      TabIndex        =   1
      Top             =   0
      Width           =   8775
      Begin VB.CheckBox chkJump 
         Caption         =   "跳过零(&J)"
         Height          =   255
         Left            =   4440
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Frame fraVal 
         Caption         =   "取值"
         Height          =   615
         Left            =   5880
         TabIndex        =   15
         Top             =   960
         Width           =   2775
         Begin VB.CommandButton cmdZero 
            Caption         =   "Z"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   19
            Top             =   240
            Width           =   255
         End
         Begin VB.HScrollBar sroVal 
            Height          =   255
            LargeChange     =   50
            Left            =   120
            Max             =   300
            Min             =   50
            TabIndex        =   16
            Top             =   240
            Value           =   150
            Width           =   1695
         End
         Begin VB.Label lblVal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "±150"
            Height          =   180
            Left            =   1920
            TabIndex        =   17
            Top             =   270
            Width           =   450
         End
      End
      Begin VB.Frame fraZoom 
         Caption         =   "缩放"
         Height          =   615
         Left            =   5880
         TabIndex        =   13
         Top             =   120
         Width           =   2775
         Begin VB.CommandButton cmdZero 
            Caption         =   "Z"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   20
            Top             =   240
            Width           =   255
         End
         Begin VB.HScrollBar sroZoom 
            Height          =   255
            LargeChange     =   50
            Left            =   120
            Max             =   100
            Min             =   5
            TabIndex        =   14
            Top             =   240
            Value           =   10
            Width           =   1695
         End
         Begin VB.Label lblZoom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "×10"
            Height          =   180
            Left            =   1920
            TabIndex        =   18
            Top             =   270
            Width           =   360
         End
      End
      Begin VB.ComboBox cboXY 
         Height          =   300
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtExp 
         Height          =   270
         Left            =   1080
         TabIndex        =   10
         Text            =   "x^2"
         Top             =   120
         Width           =   4695
      End
      Begin VB.CheckBox chkFlip 
         Caption         =   "反转(&F)"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "绘制(&D)"
         Default         =   -1  'True
         Height          =   495
         Left            =   3240
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Cancel          =   -1  'True
         Caption         =   "清除(&C)"
         Height          =   495
         Left            =   4560
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Frame fraColor 
         Caption         =   "颜色"
         Height          =   1095
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3015
         Begin VB.PictureBox picColor 
            BackColor       =   &H00000000&
            Height          =   735
            Left            =   120
            ScaleHeight     =   675
            ScaleWidth      =   435
            TabIndex        =   6
            Top             =   240
            Width           =   495
         End
         Begin VB.HScrollBar sroColor 
            Height          =   255
            Index           =   0
            LargeChange     =   50
            Left            =   720
            Max             =   255
            SmallChange     =   10
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
         Begin VB.HScrollBar sroColor 
            Height          =   255
            Index           =   1
            LargeChange     =   50
            Left            =   720
            Max             =   255
            SmallChange     =   10
            TabIndex        =   4
            Top             =   480
            Width           =   2055
         End
         Begin VB.HScrollBar sroColor 
            Height          =   255
            Index           =   2
            LargeChange     =   50
            Left            =   720
            Max             =   255
            SmallChange     =   10
            TabIndex        =   3
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         Height          =   180
         Left            =   840
         TabIndex        =   12
         Top             =   165
         Width           =   90
      End
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   0
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.PictureBox picPad 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5595
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   1680
      Width           =   5655
      Begin VB.Line ln 
         Index           =   1
         X1              =   2790
         X2              =   2790
         Y1              =   0
         Y2              =   5640
      End
      Begin VB.Line ln 
         Index           =   0
         X1              =   0
         X2              =   5640
         Y1              =   2790
         Y2              =   2790
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Zoom As Integer, iVal As Integer

Private Sub cmdClear_Click()
picPad.Cls
txtExp.Text = ""
End Sub

Private Sub cmdDraw_Click()
On Error GoTo errH:
Dim i As Long, stmp As String, tmp(2) As Long
picPad.Cls
i = -iVal
If cboXY.ListIndex = 0 Then
Do Until i = iVal
i = i + 1
If chkJump.Value = 1 And i = 0 Then i = 1
stmp = Replace(LCase(txtExp.Text), "x", i)
tmp(0) = SC.Eval(stmp)
If chkFlip.Value = 0 Then
'picPad.PSet (-(i * Zoom), tmp(0) * Zoom)
If tmp(1) <> 0 Then picPad.Line (i * Zoom, -(tmp(0) * Zoom))-(tmp(2) * Zoom, -(tmp(1) * Zoom))
tmp(1) = tmp(0)
tmp(2) = i
Else
'picPad.PSet (i * Zoom, tmp(0) * Zoom)
If tmp(1) <> 0 Then picPad.Line (-(i * Zoom), -(tmp(0) * Zoom))-(-(tmp(2) * Zoom), -(tmp(1) * Zoom))
tmp(1) = tmp(0)
tmp(2) = i
End If
Loop
Else
Do Until i = iVal
i = i + 1
If chkJump.Value = 1 And i = 0 Then i = 1
stmp = Replace(LCase(txtExp.Text), "y", i)
tmp(0) = SC.Eval(stmp)
If chkFlip.Value = 0 Then
'picPad.PSet (-(tmp(0) * Zoom), i * Zoom)
If tmp(1) <> 0 Then picPad.Line (tmp(0) * Zoom, -(i * Zoom))-(tmp(1) * Zoom, -(tmp(2) * Zoom))
tmp(1) = tmp(0)
tmp(2) = i
Else
'picPad.PSet (tmp(0) * Zoom, i * Zoom)
If tmp(1) <> 0 Then picPad.Line (tmp(0) * Zoom, i * Zoom)-(tmp(1) * Zoom, tmp(2) * Zoom)
tmp(1) = tmp(0)
tmp(2) = i
End If
Loop
End If
Exit Sub
errH:
MsgBox "绘制过程中出现错误 " & Err.Number & vbCrLf & Err.Description & vbCrLf & "请检查输入是否正确", 48, "提示"
End Sub

Private Sub cmdZero_Click(Index As Integer)
Select Case Index
Case 0
sroZoom.Value = 10
sroZoom_Change
Case 1
sroVal.Value = 150
sroVal_Change
End Select
End Sub

Private Sub Form_Load()
cboXY.ListIndex = 0
iVal = 150
Zoom = 10
End Sub

Private Sub Form_Resize()
On Error Resume Next
With picPad
.Move 0, picTool.Height, Me.ScaleWidth, Me.ScaleHeight - picTool.Height
.ScaleTop = -.ScaleHeight / 2
.ScaleLeft = -.ScaleWidth / 2
.Cls
End With
With ln(0)
.X1 = picPad.ScaleLeft
.Y1 = 0
.X2 = picPad.ScaleWidth
.Y2 = 0
End With
With ln(1)
.X1 = 0
.Y1 = picPad.ScaleTop
.X2 = 0
.Y2 = picPad.ScaleHeight
End With
End Sub

Private Sub sroColor_Change(Index As Integer)
picColor.BackColor = RGB(sroColor(0).Value, sroColor(1).Value, sroColor(2).Value)
picPad.ForeColor = picColor.BackColor
End Sub

Private Sub sroColor_Scroll(Index As Integer)
sroColor_Change Index
End Sub

Private Sub sroVal_Change()
iVal = sroVal.Value
lblVal = "±" & iVal
End Sub

Private Sub sroVal_Scroll()
sroVal_Change
End Sub

Private Sub sroZoom_Change()
Zoom = sroZoom.Value
lblZoom.Caption = "×" & Zoom
End Sub

Private Sub sroZoom_Scroll()
sroZoom_Change
End Sub
