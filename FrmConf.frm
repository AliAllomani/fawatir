VERSION 5.00
Begin VB.Form FrmConf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ŒÌ«—« "
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   Icon            =   "FrmConf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   4410
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "ﬂ·„… «·„—Ê— "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   4215
      Begin VB.TextBox edit_pass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox edit_pass_conf 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ·„… «·„—Ê— : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " √ﬂÌœ ﬂ·„… «·„—Ê— :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.CommandButton ok_cmd 
      Caption         =   "„Ê«›ﬁ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton can_cmd 
      Caption         =   "≈·€«¡ «·√„—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "FrmConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub can_cmd_Click()
    Unload Me
End Sub

Private Sub edit_pass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ok_cmd_Click
End Sub
Private Sub edit_pass_conf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ok_cmd_Click
End Sub


Private Sub ok_cmd_Click()
    new_pass = Trim(edit_pass.text)
    If new_pass <> "" Then
        If new_pass = Trim(edit_pass_conf.text) Then
            Dim rs_pass As New Recordset
            rs_pass.Open "update admin set pass='" & new_pass & "'", DB
            Unload Me
        Else
            MsgBox "ﬂ·„… «·„—Ê— Ê  √ﬂÌœÂ« €Ì— „ ÿ«»ﬁ Ì‰ ", vbExclamation, " €Ì— ﬂ·„… «·„—Ê—"
        End If
    Else
        MsgBox "Ì—ÃÏ «œŒ«· ﬂ·„… «·„—Ê— ", vbExclamation, " €Ì— ﬂ·„… «·„—Ê—"
    End If
End Sub

