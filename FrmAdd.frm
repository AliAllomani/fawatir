VERSION 5.00
Begin VB.Form FrmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����� ������"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "����� ������� ��� ������� "
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
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin VB.CommandButton add_cmd 
      Caption         =   "�����"
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
      TabIndex        =   4
      Top             =   3480
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
         ItemData        =   "FrmAdd.frx":0000
         Left            =   240
         List            =   "FrmAdd.frx":0002
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox daytxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox monthtxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox yeartxt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   8
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
         Caption         =   "������ : "
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
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "������� :"
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
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "��� �������� : "
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
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "����� �������� :"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_cmd_Click()



bill_number.Text = Trim(bill_number.Text)


If bill_number.Text <> "" And daytxt.Text <> "" And monthtxt.Text <> "" And yeartxt.Text <> "" Then
GoTo doit
Else
MsgBox ("���� ����� ���� ������")
End If


doit:
Set rs = New Recordset

bill_number.Text = Trim(bill_number.Text)
date_text = monthtxt.Text & "/" & daytxt.Text & "/" & yeartxt.Text





rs.Open "insert into bills (customer,serial,date_txt,comments) values " & _
"('" & customer_id.ItemData(customer_id.ListIndex) & "','" & bill_number.Text & "','" & date_text & "','" & commentstxt.Text & "')", DB

If Check1.Value = Checked Then
Unload Me
Else
bill_number.Text = ""
daytxt.Text = ""
monthtxt.Text = ""
yeartxt.Text = ""

End If



End Sub

Private Sub Form_Load()
Me.Icon = frmmain.Icon


Set rs = New Recordset

sql_main = "select * from customers order by name asc"

rs.Open sql_main, DB, adOpenStatic, adLockOptimistic

If rs.RecordCount = 0 Then

MsgBox "�� ���� ����� , ���� ����� ����� ����"
Unload Me


Else
For i = 1 To rs.RecordCount

Call customer_id.AddItem(rs!Name)
customer_id.ItemData(customer_id.ListCount - 1) = rs!id

rs.MoveNext

Next i

customer_id.ListIndex = 0

End If
End Sub

