VERSION 5.00
Begin VB.Form FrmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ⁄œÌ· ”Ã·"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox idtxt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
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
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3120
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
      Height          =   2655
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox customer_id 
         Height          =   315
         ItemData        =   "FrmEdit.frx":0000
         Left            =   240
         List            =   "FrmEdit.frx":0002
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox daytxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox monthtxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox yeartxt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox commentstxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   240
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox bill_number 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2655
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
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "„·«ÕŸ«  :"
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
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "—ﬁ„ «·›« Ê—… : "
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
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   " «—ÌŒ «·›« Ê—… :"
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
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_cmd_Click()

bill_number.Text = Trim(bill_number.Text)


If bill_number.Text <> "" And daytxt.Text <> "" And monthtxt.Text <> "" And yeartxt.Text <> "" Then
GoTo doit
Else
MsgBox ("Ì—ÃÏ «ﬂ„«· Ã„Ì⁄ «·ÕﬁÊ·")
End If


doit:
Set rs = New Recordset

bill_number.Text = Trim(bill_number.Text)
date_text = monthtxt.Text & "/" & daytxt.Text & "/" & yeartxt.Text


rs.Open "update bills set customer='" & customer_id.ItemData(customer_id.ListIndex) & "' , serial='" & bill_number.Text & _
"' , date_txt='" & date_text & "' , comments='" & commentstxt.Text & "' where id=" & idtxt.Text, DB

Call frmmain.refresh_now
Unload Me



End Sub

