VERSION 5.00
Begin VB.Form FrmRecEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ⁄œÌ· »Ì«‰"
   ClientHeight    =   2670
   ClientLeft      =   6945
   ClientTop       =   4470
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox type_txt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox idtxt 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton edit_cmd 
         Caption         =   " ⁄œÌ·"
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
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox price_txt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox count_txt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox nametxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "«·ÊÕœ…"
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
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬂ„Ì…"
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
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "”⁄— «·«›—«œÌ"
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
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "«”„ «·»Ì«‰ : "
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
         Index           =   0
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmRecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub edit_cmd_Click()


If nametxt.Text <> "" And count_txt.Text <> "" And price_txt.Text <> "" Then
GoTo doit
Else
MsgBox ("Ì—ÃÏ «ﬂ„«· Ã„Ì⁄ «·ÕﬁÊ·")
End If


doit:
Set rs = New Recordset

nametxt.Text = Trim(nametxt.Text)


rs.Open "update records set name='" & nametxt.Text & "' , count_txt='" & count_txt.Text & _
"' , price='" & price_txt.Text & "',type_txt='" & type_txt.Text & "'  where id=" & idtxt.Text, DB


Call FrmRecords.show_rec_records
Unload Me


End Sub

