VERSION 5.00
Begin VB.Form FrmCusAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«÷«›… ⁄„Ì·"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "«÷«›…"
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
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox cus_name 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox cus_tel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox cus_add 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   885
         IMEMode         =   3  'DISABLE
         Left            =   240
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox cus_comments 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   885
         IMEMode         =   3  'DISABLE
         Left            =   240
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox cus_user 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "≈”„ «·⁄„Ì· :"
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
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·Â« › :"
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄‰Ê«‰ :"
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ÊﬂÌ· :"
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmCusAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Set rs = New Recordset

sql_mainx = "insert into customers (name,tel,address,user_txt,comments) " & _
"values ('" & cus_name.Text & "','" & _
cus_tel.Text & "','" & cus_add.Text & "','" & _
cus_user.Text & "','" & cus_comments.Text & "')"

rs.Open sql_mainx, DB

Call FrmCustomers.refresh_customers

Unload Me

End Sub

