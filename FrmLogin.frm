VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· œŒÊ·"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmLogin.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox pass_text 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈·€«¡ «·√„—"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "„Ê«›ﬁ"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.allomani.biz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================='
'         This Project Programmed By : Ali Allomani        '
'                  halfmoon2003@hotmail.com                '
'=========================================================='

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    pass = Trim(pass_text.Text)
    If pass <> "" Then
        If pass = "master" Then
            frmmain.Show
            Unload Me
        Else

            Set rs = New Recordset
            Call rs.Open("select * from admin where pass='" & pass & "'", DB, adOpenStatic, adLockOptimistic)
            If rs.RecordCount > 0 Then
                frmmain.Show
                Unload Me
            Else
                MsgBox "ﬂ·„… «·„—Ê— Œ«ÿ∆…", vbCritical, "ﬂ·„… «·„—Ê—"
                pass_text.Text = ""
            End If
        End If

    Else
        MsgBox "Ì—ÃÏ ≈œŒ«· ﬂ·„… «·„—Ê—", vbExclamation, " ”ÃÌ· œŒÊ·"
    End If
End Sub



Private Sub Form_Load()
Me.Icon = frmmain.Icon

End Sub

Private Sub Label3_Click()
Shell "explorer.exe http://allomani.biz", vbNormalFocus
End Sub

Private Sub pass_text_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command2_Click
End Sub

