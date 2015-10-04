VERSION 5.00
Begin VB.Form FrmMenu 
   Caption         =   "›Ê« Ì—"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   Icon            =   "FrmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "„Ê«›ﬁ"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox Combo 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "FrmMenu.frx":000C
      Left            =   1320
      List            =   "FrmMenu.frx":0016
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "Programmed by : Ali Allomani [www.allomani.biz]"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   3120
      Width           =   4215
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================='
'         This Project Programmed By : Ali Allomani        '
'                  halfmoon2003@hotmail.com                '
'=========================================================='

Private Sub Combo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub Command1_Click()
If Combo.ListIndex = 0 Then
frmmain.Caption = frmmain.Caption & " / " & Combo.text
Call frmmain.view_cols



'----------- Main Labels --------------------------


'If FrmEdit.text(1).Visible = True Then frmmain.text(1).Visible = True Else frmmain.text(1).Visible = False
'frmmain.text(4).Visible = FrmAdd.text(5).Visible
'frmmain.text(2).Visible = FrmAdd.text(2).Visible
'frmmain.text(5).Visible = FrmAdd.text(3).Visible
'frmmain.text(0).Visible = FrmAdd.text(0).Visible
'frmmain.text(3).Visible = FrmAdd.text(4).Visible
'frmmain.text(6).Visible = FrmAdd.text(7).Visible
'frmmain.text(7).Visible = FrmAdd.text(6).Visible

Dim i As Integer

For i = 1 To 4
frmmain.flx.Row = 0
frmmain.flx.Col = i

If frmmain.flx.ColWidth(i) > 0 Then frmmain.select_by.AddItem frmmain.flx.text
Next
frmmain.select_by.ListIndex = 1

'===============================================

Me.Hide
frmmain.Show

ElseIf Combo.ListIndex = 1 Then

Frmproducts.Show

Else
MsgBox "Â–« ﬁ”„ —∆Ì”Ì ·«Ì„ﬂ‰ «Œ Ì«—Â ", vbCritical
End If

End Sub

Private Sub Form_Load()

Combo.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
