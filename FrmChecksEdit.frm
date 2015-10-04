VERSION 5.00
Begin VB.Form FrmChecksEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Check"
   ClientHeight    =   3600
   ClientLeft      =   6945
   ClientTop       =   5220
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4935
   Begin VB.TextBox idtxt 
      Height          =   615
      Left            =   3720
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton add_cmd 
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
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox customer_id 
         Height          =   315
         ItemData        =   "FrmChecksEdit.frx":0000
         Left            =   240
         List            =   "FrmChecksEdit.frx":0002
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox daytxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox monthtxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox yeartxt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox check_number 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox name_txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox amount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox status 
         Height          =   315
         ItemData        =   "FrmChecksEdit.frx":0004
         Left            =   1200
         List            =   "FrmChecksEdit.frx":0014
         TabIndex        =   1
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "«·⁄„Ì· : "
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
         Index           =   1
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "—ﬁ„ «·‘Ìﬂ: "
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
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   " «—ÌŒ «·‘Ìﬂ:"
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
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "«”„ «·„’—Ê› ·Â : "
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
         Index           =   2
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "ﬁÌ„… «·‘Ìﬂ : "
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
         Index           =   3
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Õ«·… «·‘Ìﬂ :"
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
         Index           =   4
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmChecksEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_cmd_Click()
If check_number.Text <> "" And daytxt.Text <> "" And monthtxt.Text <> "" And yeartxt.Text <> "" Then
GoTo doit
Else
MsgBox ("Ì—ÃÏ «ﬂ„«· Ã„Ì⁄ «·ÕﬁÊ·")
End If


doit:
Set rs = New Recordset

check_number.Text = Trim(check_number.Text)
date_text = monthtxt.Text & "/" & daytxt.Text & "/" & yeartxt.Text


rs.Open "update checks set customer_id='" & customer_id.ItemData(customer_id.ListIndex) & "' ,serial='" & check_number.Text _
& "' , date_txt='" & date_text & "' , name ='" & name_txt.Text & "' , status='" & status.List(status.ListIndex) & "' , amount='" & Val(amount.Text) & "' where id=" & Val(idtxt.Text), DB


Call FrmChecks.show_checks_records
Unload Me

End Sub

