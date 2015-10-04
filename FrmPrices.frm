VERSION 5.00
Begin VB.Form FrmPrices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáÃÓÚÇÑ"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1935
      Begin VB.CommandButton Command1 
         Caption         =   "ÊÚÏíá"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox price_txt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.ComboBox prices_list 
      Height          =   4275
      Left            =   2520
      RightToLeft     =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "prices_list"
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "FrmPrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub update_prices()
prices_list.Clear

Set rsx = New Recordset

sql_cus = "select name from records group by name order by name asc"

rsx.Open sql_cus, DB, adOpenStatic, adLockOptimistic


For i = 1 To rsx.RecordCount

Call prices_list.AddItem(rsx!Name)


rsx.MoveNext

Next i

End Sub

Private Sub Command1_Click()

Set rs = New Recordset


sqlm = "select * from prices where name_txt like '" & prices_list.List(prices_list.ListIndex) & "'"


rs.Open sqlm, DB, adOpenStatic, adLockOptimistic

If rs.RecordCount > 0 Then

Set rs = New Recordset


rs.Open "update prices set price_txt='" & Val(price_txt.Text) & "' where name_txt='" & prices_list.List(prices_list.ListIndex) & "'", DB
Else


Set rs = New Recordset


rs.Open "insert into prices (name_txt,price_txt) values('" & prices_list.List(prices_list.ListIndex) & "','" & Val(price_txt.Text) & "')", DB
End If

End Sub

Private Sub Form_Load()

Call update_prices
End Sub

Private Sub prices_list_Click()
Set rs = New Recordset

sqlm = "select * from prices where name_txt like '" & prices_list.List(prices_list.ListIndex) & "'"


rs.Open sqlm, DB, adOpenStatic, adLockOptimistic


If rs.RecordCount > 0 Then
price_txt.Text = rs!price_txt

Else
price_txt.Text = "0"
End If
End Sub
