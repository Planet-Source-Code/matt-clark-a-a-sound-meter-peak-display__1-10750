VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   150
      ScaleHeight     =   2715
      ScaleWidth      =   7215
      TabIndex        =   8
      Top             =   3900
      Width           =   7215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H0000FF00&
      Height          =   3000
      Left            =   150
      ScaleHeight     =   3000
      ScaleWidth      =   7215
      TabIndex        =   12
      Top             =   6270
      Width           =   7215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3360
      Top             =   1590
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   3180
      Top             =   2100
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   3120
      Top             =   2580
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   5670
      Top             =   840
      Width           =   1635
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Get Input"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5700
      TabIndex        =   16
      Top             =   870
      Width           =   1425
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   5940
      Top             =   3210
      Width           =   525
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Reset"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5970
      TabIndex        =   15
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5100
      TabIndex        =   14
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Graph View"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5700
      TabIndex        =   13
      Top             =   1230
      Width           =   1425
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   5670
      Top             =   1200
      Width           =   1635
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      X1              =   5010
      X2              =   4710
      Y1              =   1470
      Y2              =   2340
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FFFFFF&
      X1              =   4620
      X2              =   5100
      Y1              =   1470
      Y2              =   2370
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   4620
      X2              =   5010
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   4620
      X2              =   5010
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5700
      TabIndex        =   11
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   5670
      Top             =   2370
      Width           =   1635
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sound Meter V1.5"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5670
      TabIndex        =   10
      Top             =   450
      Width           =   1365
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form2.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   1875
      Left            =   2130
      TabIndex        =   2
      Top             =   2550
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   6750
      X2              =   6750
      Y1              =   3870
      Y2              =   3570
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Max"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6810
      TabIndex        =   9
      Top             =   3600
      Width           =   1185
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   5670
      Top             =   2010
      Width           =   1635
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   5670
      Top             =   1590
      Width           =   1635
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5190
      Picture         =   "Form2.frx":00BC
      Top             =   1890
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Freeze"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5700
      TabIndex        =   7
      Top             =   2040
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5220
      Picture         =   "Form2.frx":04FE
      Top             =   1500
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Display Sensitivity X2"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5700
      TabIndex        =   4
      Top             =   1620
      Width           =   1605
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   1020
      X2              =   5910
      Y1              =   7170
      Y2              =   7170
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   120
      X2              =   3600
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   195
      Left            =   210
      Top             =   3630
      Width           =   75
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   3810
      X2              =   4170
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Min"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   30
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Max"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   3330
      Width           =   1185
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   3810
      X2              =   4170
      Y1              =   3450
      Y2              =   3450
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   105
      Left            =   3900
      Top             =   810
      Width           =   225
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   990
      TabIndex        =   3
      Top             =   6900
      Visible         =   0   'False
      Width           =   4905
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Too Sensitive"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3090
      TabIndex        =   1
      Top             =   6930
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   30
      X2              =   30
      Y1              =   30
      Y2              =   2070
   End
   Begin VB.Line Line1 
      X1              =   1470
      X2              =   30
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Shape Shape1 
      Height          =   1845
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   1995
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   30
      X2              =   1170
      Y1              =   30
      Y2              =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5400
      TabIndex        =   0
      Top             =   60
      Width           =   1965
   End
   Begin VB.Line Line4 
      X1              =   30
      X2              =   1920
      Y1              =   30
      Y2              =   1800
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Command1_Click()
Form3.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub



Private Sub Label1_Change()
If vol1 > Max Then
Max = vol1
GoTo 1
End If
If vol1 < Min Then
Min = vol1
GoTo 1
End If
Exit Sub
1
Label13.Caption = "Max = " & Max / 1000 & "   Min = " & Min / 1000
End Sub

Private Sub Label11_Click()
End
End Sub

Private Sub Label12_Click()
Form3.Show
End Sub

Private Sub Label14_Click()
  Max = vol1
  Min = vol1
  Label13.Caption = "Max = " & Max / 1000 & "   Min = " & Min / 1000
End Sub

Private Sub Label15_Click()
StartInput
End Sub

Private Sub Label2_Click()
Label3.Visible = True
Timer3.Enabled = True
End Sub

Private Sub Label4_Click()
Label3.Visible = True
Timer3.Enabled = True
End Sub

Private Sub Label5_Click()
If XSens = Empty Or XSens = False Then
XSens = True
Image1.Visible = True
Else
XSens = False
Image1.Visible = False
End If
End Sub

Private Sub Label8_Click()
If Form1.Timer1.Enabled = False Then
Form1.Timer1.Enabled = True
Image2.Visible = False
Else
Form1.Timer1.Enabled = False
Image2.Visible = True
End If
End Sub

Private Sub Timer1_Timer()
If Label2.Visible = False Then
Label2.Visible = True
Line5.Visible = True
Else
Label2.Visible = False
Line5.Visible = False
End If
End Sub
Sub TooSens()
If Timer2.Interval >= 8000 Then GoTo 1
Timer2.Interval = Timer2.Interval + 4000
1
Timer1.Enabled = True
Timer2.Enabled = True
Label4.Visible = True
End Sub

Private Sub Timer2_Timer()
Timer1.Enabled = False
Timer2.Enabled = False
Line5.Visible = False
Label2.Visible = False
Label4.Visible = False
End Sub

Private Sub Timer3_Timer()
Label3.Visible = False
Timer3.Enabled = False
End Sub
