VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmRecords 
   Caption         =   "«·”Ã·« "
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   22
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox total_txt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   " ⁄œÌ· »Ì«‰«  «·›« Ê—…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13560
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ÕœÌÀ"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox bill_id 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "«÷«›… »Ì«‰ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   8775
      Begin VB.ComboBox type_txt 
         Height          =   315
         Left            =   3000
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox name_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         TabIndex        =   9
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox count_txt 
         Height          =   285
         Left            =   4680
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox price_txt 
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton add_cmd 
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
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label8 
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
         Left            =   3480
         TabIndex        =   24
         Top             =   240
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
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   975
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
         Left            =   4800
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "«·»Ì‹‹‹‹‹‹«‰"
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
         Left            =   6600
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CommandButton del_cmd 
      Caption         =   "Õ–›"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton edit_cmd 
      Caption         =   " ⁄œÌ·"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid flx 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   7
      ForeColor       =   -2147483642
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      Enabled         =   -1  'True
      RightToLeft     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "«·„Ã„Ê⁄ :"
      Height          =   375
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label cust_name 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "«”„ «·⁄„Ì· :"
      Height          =   255
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label bill_date 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   " «—ÌŒ «·›« Ê—… : "
      Height          =   255
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label bill_number 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "—ﬁ„ «·›« Ê—… : "
      Height          =   255
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "FrmRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub show_rec_records()

Dim x_total As Integer

Call update_products
Call view_rec_cols

Set rs = New Recordset
rs.Open "select * from records where bill=" & Trim(bill_id.Text), DB, adOpenStatic, adLockOptimistic

x_total = 0

If rs.RecordCount <> 0 Then

flx.Rows = rs.RecordCount + 1

For i = 1 To rs.RecordCount
flx.Row = i
flx.Col = 0
flx.Text = rs!id
flx.Col = 1
flx.Text = i

flx.Col = 2
flx.Text = rs!price * rs!count_txt

x_total = x_total + (rs!price * rs!count_txt)


flx.Col = 3
flx.Text = rs!Name


flx.Col = 4
flx.Text = rs!count_txt

flx.Col = 5
flx.Text = rs!price

flx.Col = 6
flx.Text = rs!type_txt


rs.MoveNext

Next i

total_txt.Text = x_total

End If
End Sub
Public Sub view_rec_cols()

flx.Clear

flx.Rows = 1
flx.Row = 0
flx.Col = 1
flx.Text = "#"

flx.ColWidth(0) = 0
flx.ColWidth(1) = 500
flx.ColWidth(2) = 600
flx.ColWidth(3) = Me.Width / 1.7
flx.ColWidth(4) = 500
flx.ColWidth(5) = 1000
flx.ColWidth(6) = 500


flx.Col = 2
flx.Text = "«·”⁄— «·«Ã„«·Ì"
flx.Col = 3
flx.Text = "«·»Ì«‰ "
flx.Col = 4
flx.Text = "«·ﬂ„Ì…"



flx.Col = 5
flx.Text = "”⁄— «·«›—«œÌ"


flx.Col = 6
flx.Text = "«·ÊÕœ…"





End Sub

Private Sub add_cmd_Click()
If name_txt.Text <> "" And count_txt.Text <> "" And price_txt.Text <> "" Then
Set rsi = New Recordset
rsi.Open "insert into records (bill,name,price,count_txt,type_txt) values ('" & Val(bill_id.Text) & "','" & name_txt.Text & _
"','" & Val(price_txt.Text) & "','" & Val(count_txt.Text) & "','" & type_txt.Text & "')", DB

Call show_rec_records

price_txt.Text = ""
count_txt.Text = ""
Else
MsgBox " «·—Ã«¡ «ﬂ„«· Ã„Ì⁄ «·ÕﬁÊ·"
End If
End Sub

Private Sub Command1_Click()
Call FrmRecords.show_rec_records

End Sub

Private Sub Command2_Click()
FrmEdit.idtxt.Text = bill_id.Text
Call frmmain.show_edit

End Sub

Private Sub Command3_Click()
Call Grid2HTML(flx, "File=c:\Cars_Gen.html", "<html dir=rtl><title>Generated Bill</title> <br><b> «”„ «·⁄„Ì· : </b> " & cust_name.Caption & "<br> <b> —ﬁ„ «·›« Ê—… : </b> " & bill_number.Caption & "<br> <b>  «—ÌŒ «·›« Ê—… : </b>" & bill_date.Caption & "<br><br>", total_txt.Text)
Shell "explorer.exe c:\Cars_Gen.html", vbNormalFocus
End Sub

Private Sub del_cmd_Click()
If Trim(flx.TextMatrix(flx.Row, 0)) <> "" Then
id_txt = Trim(flx.TextMatrix(flx.Row, 0))

If MsgBox("Â· «‰  „ √ﬂœ „‰ √‰ﬂ  —Ìœ Õ–› «·”Ã· —ﬁ„ : " & Trim(flx.TextMatrix(flx.Row, 1)) & " ø ", vbQuestion + vbYesNo) = vbYes Then
Set rs = New Recordset
rs.Open "delete from records where id=" & id_txt, DB
Call show_rec_records
End If
End If
End Sub

Private Sub edit_cmd_Click()
Set rs = New Recordset

If Trim(Trim(flx.TextMatrix(flx.Row, 0))) <> "" Then

sql_x = "select * from records where id=" & Trim(flx.TextMatrix(flx.Row, 0))

rs.Open sql_x, DB, adOpenStatic, adLockOptimistic

If rs.RecordCount <> 0 Then


FrmRecEdit.nametxt.Text = rs!Name
FrmRecEdit.count_txt.Text = rs!count_txt
FrmRecEdit.price_txt.Text = rs!price

FrmRecEdit.idtxt.Text = rs!id


Call FrmRecEdit.Show(1)

End If
End If
End Sub

Private Sub flx_DblClick()
Call edit_cmd_Click
End Sub

Private Sub Form_Load()
Call update_products
End Sub

Private Sub Form_Resize()
On Error Resume Next
flx.Width = Me.Width - 300
flx.Height = Me.Height - 3000
flx.ColWidth(3) = Me.Width / 1.7


End Sub

Private Sub update_products()
On Error Resume Next

name_txt.Clear

Set rsx = New Recordset

sql_cus = "select name from records group by name order by name asc"

rsx.Open sql_cus, DB, adOpenStatic, adLockOptimistic


For i = 1 To rsx.RecordCount

Call name_txt.AddItem(rsx!Name)


rsx.MoveNext

Next i

type_txt.Clear

Set rsx = New Recordset

sql_cus = "select type_txt from records group by type_txt order by type_txt asc"

rsx.Open sql_cus, DB, adOpenStatic, adLockOptimistic


For z = 1 To rsx.RecordCount

Call type_txt.AddItem(rsx!type_txt)


rsx.MoveNext

Next z
End Sub

Private Sub name_txt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then add_cmd_Click
End Sub

Private Sub count_txt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then add_cmd_Click
End Sub

Private Sub price_txt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then add_cmd_Click
End Sub

