VERSION 5.00
Begin VB.Form FrmCustomers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈œ«—… «·⁄„·«¡"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "⁄„Ì· ÃœÌœ"
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox customer_id 
      Height          =   315
      ItemData        =   "FrmCustomers.frx":0000
      Left            =   1200
      List            =   "FrmCustomers.frx":0002
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   150
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   4935
      Begin VB.CommandButton Command3 
         Caption         =   "Õ–›"
         DragIcon        =   "FrmCustomers.frx":0004
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox cus_user 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   13
         Top             =   1320
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
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
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox cus_comments 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   885
         IMEMode         =   3  'DISABLE
         Left            =   360
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox cus_add 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   885
         IMEMode         =   3  'DISABLE
         Left            =   360
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox cus_tel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox cus_name 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3375
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1320
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
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2880
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
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   855
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
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   855
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄„Ì· :"
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
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "FrmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub refresh_customers()
customer_id.Clear

Set rs = New Recordset

sql_mainx = "select * from customers order by name asc"

rs.Open sql_mainx, DB, adOpenStatic, adLockOptimistic

If rs.RecordCount <> 0 Then

For i = 1 To rs.RecordCount

Call customer_id.AddItem(rs!Name)
customer_id.ItemData(customer_id.ListCount - 1) = rs!id

rs.MoveNext

Next i

customer_id.ListIndex = 0

End If
End Sub

Private Sub Command1_Click()
Call FrmCusAdd.Show(1)
End Sub

Private Sub Command2_Click()

Set rs = New Recordset

sql_mainx = "update customers set name='" & cus_name.Text & "' , tel='" & _
cus_tel.Text & "' , address='" & cus_add.Text & "' , user_txt ='" & _
cus_user.Text & "' , comments='" & cus_comments.Text & "' where id=" & customer_id.ItemData(customer_id.ListIndex)


rs.Open sql_mainx, DB

End Sub

Private Sub Command3_Click()
If Trim(customer_id.ItemData(customer_id.ListIndex)) <> "" Then
id_txt = customer_id.ItemData(customer_id.ListIndex)

If MsgBox("Â· «‰  „ √ﬂœ „‰ √‰ﬂ  —Ìœ Õ–› «·⁄„Ì·  " & customer_id.List(customer_id.ListIndex) & " ø ", vbQuestion + vbYesNo) = vbYes Then
Set rs = New Recordset
rs.Open "delete from customers where id=" & id_txt, DB
Call refresh_customers
End If
End If
End Sub

Private Sub customer_id_Click()
On Error Resume Next

Set rs = New Recordset

sql_mainx = "select * from customers where id=" & customer_id.ItemData(customer_id.ListIndex)

rs.Open sql_mainx, DB, adOpenStatic, adLockOptimistic

cus_name.Text = ""
cus_tel.Text = ""
cus_add.Text = ""
cus_user.Text = ""
cus_comments.Text = ""


If rs.RecordCount <> 0 Then


cus_name.Text = rs!Name
cus_tel.Text = rs!tel
cus_add.Text = rs!address
cus_user.Text = rs!User_txt
cus_comments.Text = rs!Comments

End If
End Sub

Private Sub Form_Load()
Call refresh_customers
End Sub
