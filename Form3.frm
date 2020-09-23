VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FF00&
      Height          =   2715
      Left            =   30
      ScaleHeight     =   2715
      ScaleWidth      =   11835
      TabIndex        =   1
      Top             =   2880
      Width           =   11835
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FF00&
      Height          =   2715
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   0
      Width           =   11835
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FF00&
      Height          =   2715
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   11835
      TabIndex        =   2
      Top             =   5760
      Width           =   11835
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   10200
      TabIndex        =   3
      Top             =   8550
      Width           =   1605
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   10170
      Top             =   8550
      Width           =   1635
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()

End Sub

Private Sub Label8_Click()
Form3.Hide
End Sub

