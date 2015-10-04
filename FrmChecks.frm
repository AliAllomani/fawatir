VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmChecks 
   Caption         =   "‘Ìﬂ« "
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   3600
      TabIndex        =   18
      Top             =   1920
      Width           =   4575
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
         Left            =   240
         TabIndex        =   20
         Top             =   240
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "«·„Ã„Ê⁄ :"
         Height          =   375
         Index           =   0
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "«· — Ì» Õ”»"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   720
      Width           =   2655
      Begin VB.ComboBox select_by 
         Height          =   315
         ItemData        =   "FrmChecks.frx":0000
         Left            =   120
         List            =   "FrmChecks.frx":0002
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   " ’«⁄œÌ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   " ‰«“·Ì"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "«·»ÕÀ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   8655
      Begin VB.CommandButton srch_cmd 
         Caption         =   "»ÕÀ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox serial_txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox date_txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox customer_id 
         Height          =   315
         ItemData        =   "FrmChecks.frx":0004
         Left            =   4440
         List            =   "FrmChecks.frx":0006
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "—ﬁ„ «·›« Ê—… :"
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "«· «—ÌŒ : "
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
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1335
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
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "«÷«›…"
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton edit_cmd 
      Caption         =   " ⁄œÌ·"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton del_cmd 
      Caption         =   "Õ–›"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox bill_id 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ÕœÌÀ"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid flx 
      Height          =   5055
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   8
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
End
Attribute VB_Name = "FrmChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub show_checks_records()

Dim x_total As Integer
Dim sql_checks As String

Call view_checks_cols

Select Case select_by.ListIndex
Case 0: str_order = "id"
Case 1: str_order = "date_txt"
Case 2: str_order = "serial"
Case 3: str_order = "customer_id"
Case 4: str_order = "name"
Case 5: str_order = "amount"
Case 6: str_order = "status"
End Select


If Option2.Value = True Then
str_desc = "DESC"
Else
str_desc = "ASC"
End If

Set rs = New Recordset


sql_checks = "select * from checks where"

If customer_id.ItemData(customer_id.ListIndex) <> 0 Then
sql_checks = sql_checks & "  customer_id=" & customer_id.ItemData(customer_id.ListIndex) & " and "
End If

sql_checks = sql_checks & " serial like '%" & serial_txt.Text & "%' and date_txt like '%" & date_txt.Text & "%'"



sql_checks = sql_checks & " order by " & str_order & " " & str_desc



rs.Open sql_checks, DB, adOpenStatic, adLockOptimistic

total_txt.Text = 0

If rs.RecordCount <> 0 Then

flx.Rows = rs.RecordCount + 1

For i = 1 To rs.RecordCount
flx.Row = i
flx.Col = 0
flx.Text = rs!id
flx.Col = 1
flx.Text = i

flx.Col = 2
flx.Text = Format(rs!date_txt, "dd/mm/yyyy")



flx.Col = 3
flx.Text = rs!serial


Set rs2 = New Recordset
sql_x = "select * from customers where id=" & rs!customer_id & ""

rs2.Open sql_x, DB, adOpenStatic, adLockOptimistic

flx.Col = 4
flx.Text = rs2!Name

flx.Col = 5
flx.Text = rs!Name

flx.Col = 6
flx.Text = rs!amount

flx.Col = 7
flx.Text = rs!status

total_txt.Text = Val(total_txt.Text) + rs!amount

rs.MoveNext


Next i



End If
End Sub
Public Sub view_checks_cols()

flx.Clear

flx.Rows = 1
flx.Row = 0
flx.Col = 1
flx.Text = "#"

flx.ColWidth(0) = 0
flx.ColWidth(1) = 500
flx.ColWidth(2) = 1000
flx.ColWidth(3) = 2500
flx.ColWidth(4) = 3500
flx.ColWidth(5) = 5000
flx.ColWidth(6) = 2000
flx.ColWidth(7) = 1500


flx.Col = 2
flx.Text = " «—ÌŒ «·‘Ìﬂ"
flx.Col = 3
flx.Text = "—ﬁ„ «·‘Ìﬂ"

flx.Col = 4
flx.Text = "«”„ «·⁄„Ì·"
flx.Col = 5
flx.Text = "«”„ «·„’—Ê› ·Â "
flx.Col = 6
flx.Text = "ﬁÌ„… «·‘Ìﬂ"

flx.Col = 7
flx.Text = "Õ«·… «·‘Ìﬂ"






End Sub



Private Sub CmdAdd_Click()
FrmChecksAdd.Show 1
End Sub

Private Sub Command1_Click()
Call show_checks_records

End Sub



Private Sub Command3_Click()
Call Grid2HTML(flx, "File=c:\Cars_Gen.html", "<html dir=rtl><title>Generated Bill</title>")
Shell "explorer.exe c:\Cars_Gen.html", vbNormalFocus
End Sub

Private Sub del_cmd_Click()
If Trim(flx.TextMatrix(flx.Row, 0)) <> "" Then
id_txt = Trim(flx.TextMatrix(flx.Row, 0))

If MsgBox("Â· «‰  „ √ﬂœ „‰ √‰ﬂ  —Ìœ Õ–› «·”Ã· —ﬁ„ : " & Trim(flx.TextMatrix(flx.Row, 1)) & " ø ", vbQuestion + vbYesNo) = vbYes Then
Set rs = New Recordset
rs.Open "delete from checks where id=" & id_txt, DB
Call show_checks_records
End If
End If
End Sub

Private Sub edit_cmd_Click()
Set rs = New Recordset

sql_x = "select * from checks where id=" & Trim(flx.TextMatrix(flx.Row, 0))

rs.Open sql_x, DB, adOpenStatic, adLockOptimistic

'MsgBox rs.RecordCount

date_arr = Split(rs!date_txt, "/")


FrmChecksEdit.check_number.Text = rs!serial

FrmChecksEdit.daytxt.Text = date_arr(0)
FrmChecksEdit.monthtxt.Text = date_arr(1)
FrmChecksEdit.yeartxt.Text = date_arr(2)

FrmChecksEdit.name_txt.Text = rs!Name
FrmChecksEdit.amount.Text = rs!amount


FrmChecksEdit.idtxt.Text = rs!id

For z = 0 To 3
If rs!status = FrmChecksEdit.status.List(z) Then
FrmChecksEdit.status.ListIndex = z
End If
Next z



'---------------------------------------
FrmChecksEdit.customer_id.Clear
x_customer = rs!customer_id
Set rs = New Recordset

sql_mainx = "select * from customers order by name asc"

rs.Open sql_mainx, DB, adOpenStatic, adLockOptimistic


For i = 1 To rs.RecordCount

Call FrmChecksEdit.customer_id.AddItem(rs!Name)
FrmChecksEdit.customer_id.ItemData(FrmChecksEdit.customer_id.ListCount - 1) = rs!id

If rs!id = x_customer Then
FrmChecksEdit.customer_id.ListIndex = (i - 1)
End If

rs.MoveNext

Next i

Call FrmChecksEdit.Show(1)
End Sub

Private Sub flx_DblClick()
Call edit_cmd_Click
End Sub

Private Sub flx_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Or KeyCode = 110 Then
        Call del_cmd_Click
        End If
End Sub

Private Sub Form_Load()



Me.Icon = frmmain.Icon


Set rs = New Recordset

sql_main = "select * from customers order by name asc"

rs.Open sql_main, DB, adOpenStatic, adLockOptimistic

If rs.RecordCount = 0 Then

MsgBox "·« ÌÊÃœ ⁄„·«¡ , Ì—ÃÏ «÷«›… ⁄„·«¡ √Ê·«"
Unload Me


Else

Call customer_id.AddItem("«·ﬂ‹‹‹·")
customer_id.ItemData(0) = 0


For i = 1 To rs.RecordCount

Call customer_id.AddItem(rs!Name)
customer_id.ItemData(customer_id.ListCount - 1) = rs!id

rs.MoveNext

Next i

customer_id.ListIndex = 0

Call view_checks_cols

select_by.Clear

For i = 1 To 7
flx.Row = 0
flx.Col = i

If flx.ColWidth(i) > 0 Then select_by.AddItem flx.Text
Next
select_by.ListIndex = 1


Call show_checks_records

End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
flx.Width = Me.Width - 300
flx.Height = Me.Height - 3000



End Sub



Private Sub srch_cmd_Click()
Call show_checks_records

End Sub
