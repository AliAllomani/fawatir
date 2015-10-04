VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmmain 
   Caption         =   "«·›Ê« Ì—"
   ClientHeight    =   8925
   ClientLeft      =   855
   ClientTop       =   1530
   ClientWidth     =   14280
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   14280
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "«·—’Ìœ "
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
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   840
      Width           =   2655
      Begin VB.TextBox row_total 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   480
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.TextBox flx_size 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11640
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Text            =   "9"
      Top             =   240
      Width           =   495
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
      Height          =   1215
      Left            =   10920
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   2415
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
         TabIndex        =   10
         Top             =   720
         Width           =   975
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.ComboBox select_by 
         Height          =   315
         ItemData        =   "Form1.frx":6852
         Left            =   120
         List            =   "Form1.frx":6854
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6856
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6F6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":767E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":84A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":92CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":99E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A0F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AD4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B45E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BB72
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C286
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsImageList1 
      Left            =   7320
      Top             =   8160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C99A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D5EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E242
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EE96
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FAEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1073E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11392
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1388E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":144E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15136
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15D8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":169DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17632
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18286
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A782
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B3D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C02A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CC7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D8D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DFE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E6FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F34E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1FA62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":20176
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":211DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   690
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   1217
      ButtonWidth     =   1931
      ButtonHeight    =   1217
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "›« Ê—… ÃœÌœ…"
            Object.Tag             =   "add"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Õ–›"
            Object.Tag             =   "del"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " ⁄œÌ·"
            Object.Tag             =   "edit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " ’œÌ— ··ÿ»«⁄…"
            Object.Tag             =   "print"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " ÕœÌÀ"
            Object.Tag             =   "refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "≈œ«—… «·⁄„·«¡"
            Object.Tag             =   "customers"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«·‘Ìﬂ« "
            Object.Tag             =   "checks"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«·√”⁄«—"
            Object.Tag             =   "prices"
            ImageIndex      =   6
         EndProperty
      EndProperty
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
      Height          =   1455
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   7095
      Begin VB.ComboBox customer_id 
         Height          =   315
         ItemData        =   "Form1.frx":27A40
         Left            =   1800
         List            =   "Form1.frx":27A42
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Text            =   "customer_id"
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox date_txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox serial_txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton srch_cmd 
         Caption         =   "»ÕÀ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   1215
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
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   960
         Width           =   1335
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
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx 
      Height          =   6015
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   10610
      _Version        =   393216
      Cols            =   6
      ForeColor       =   -2147483642
      BackColorBkg    =   -2147483633
      WordWrap        =   -1  'True
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
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "»Ìﬂ”·"
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
      Left            =   10680
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "ÕÃ„ «·Œÿ :"
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
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Menu file_cmd 
      Caption         =   "&„·›"
      Begin VB.Menu mainmenu_cmd 
         Caption         =   "&«·ﬁ«∆„… «·—∆Ì”Ì…"
      End
      Begin VB.Menu ext_cmd 
         Caption         =   "Œ—ÊÃ"
      End
   End
   Begin VB.Menu tools_cmd 
      Caption         =   "&√œÊ« "
      Begin VB.Menu repair_cmd 
         Caption         =   "÷€ÿ Ê «’·«Õ ﬁ«⁄œ… «·»Ì«‰« "
      End
      Begin VB.Menu cut 
         Caption         =   "-"
      End
      Begin VB.Menu conf_cmd 
         Caption         =   "ŒÌ«—«  "
      End
   End
   Begin VB.Menu help_cmd 
      Caption         =   "&„”«⁄œ…"
      Begin VB.Menu about_cmd 
         Caption         =   "⁄‰ «·»—‰«„Ã"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================='
'         This Project Programmed By : Ali Allomani        '
'                  halfmoon2003@hotmail.com                '
'=========================================================='


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim X As Integer
Dim topRow As Integer
Dim ctl As Control
Dim lngResult As Long
Dim rs2 As Recordset
Dim i2 As Integer
Dim x_total As Integer

Private Sub about_cmd_Click()
FrmAbout.Show 1
End Sub



Private Sub conf_cmd_Click()
FrmConf.Show 1
End Sub

Private Sub ext_cmd_Click()
End
End Sub

Private Sub flx_DblClick()

If Trim(flx.TextMatrix(flx.Row, 0)) <> "" Then
FrmRecords.bill_id.Text = Trim(flx.TextMatrix(flx.Row, 0))
FrmRecords.bill_number.Caption = Trim(flx.TextMatrix(flx.Row, 2))
FrmRecords.bill_date.Caption = Trim(flx.TextMatrix(flx.Row, 4))
FrmRecords.cust_name.Caption = Trim(flx.TextMatrix(flx.Row, 3))


Call FrmRecords.view_rec_cols
Call FrmRecords.show_rec_records


FrmRecords.Show 1

End If

End Sub

Private Sub flx_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Or KeyCode = 110 Then
        Call Del_Rec
        End If
        
End Sub


Private Sub flx_size_Change()
On Error Resume Next
flx.Font.Size = Int(flx_size.Text)
End Sub

Private Sub Form_Resize()
On Error Resume Next
flx.Width = Me.Width - 300
flx.Height = Me.Height - 3000

flx.ColWidth(1) = 400
flx.ColWidth(2) = Me.Width / 2

flx.ColWidth(4) = Me.Width / 7

End Sub

Private Sub Form_Unload(Cancel As Integer)
DB.Close
'Shell App.Path & "/" & App.EXEName & ".exe nonpass", vbNormalFocus
End
End Sub

Private Sub mainmenu_cmd_Click()
DB.Close
'Shell App.Path & "/" & App.EXEName & ".exe nonpass", vbNormalFocus
End
End Sub

Private Sub repair_cmd_Click()

ms1 = MsgBox("”Ì „ ≈’·«Õ Ê÷€ÿ ﬁ«⁄œ… «·»Ì«‰«  ”Ì” €—ﬁ –·ﬂ »⁄÷ «·Êﬁ . " & vbNewLine & vbNewLine & "«·»œ¡ ›Ì «·⁄„·Ì… ø ", 524288 + 52, "«’·«Õ Ê ÷€ÿ ﬁ«⁄œ… «·»Ì«‰« ")

If ms1 = 6 Then
Me.Enabled = False
DB.Close
'1 -  €ÌÌ— „ƒ‘— «·›√—…
MousePointer = vbHourglass
  
'  «· «ﬂœ „‰ ⁄œ„ ÊÃÊœ «·„·› «·„ƒﬁ  Ê Õœ›Â ›Ì Õ«·… ÊÃÊœÂ
  If Dir(MyPath + "Temp_db.db", vbHidden) <> "" Then Kill MyPath + "Temp_db.db"
  '  ÷€ÿ Ê√’·«Õ „·› ﬁ«⁄œ… «·»Ì«‰«  Ê Ê÷⁄Â ›Ì «·„·› «·„ƒﬁ 
  DBEngine.RepairDatabase MyDBF
  DBEngine.CompactDatabase MyDBF, MyPath + "Temp_db.db", dbLangArabic, , ";pwd=master"
  '  Õœ› „·› ﬁ«⁄œ… «·»Ì«‰« 
  If Dir(MyDBF, vbHidden) <> "" Then Kill MyDBF
  '  «” »œ«· «”„ «·„·› «·„ƒﬁ  «·Ï «”„ „·› ﬁ«⁄œ… «·»Ì«‰« 
  Name MyPath + "Temp_db.db" As MyDBF
  '3 - «⁄«œ… «·„ƒ‘— ··Õ«·… «·ÿ»Ì⁄Ì…
MousePointer = vbNormal
Me.Enabled = True
u = MsgBox(" „  ⁄„·Ì… «’·«Õ Ê÷€ÿ „·› ﬁ«⁄œ… «·»Ì«‰«  »‰Ã«Õ ", 524288 + 64, "‰Ã«Õ «·⁄„·ÌÂ")
Call OpenDB
End If
End Sub
Private Sub Form_Load()


Call frmmain.view_cols

Dim i As Integer

For i = 1 To 4
frmmain.flx.Row = 0
frmmain.flx.Col = i

If frmmain.flx.ColWidth(i) > 0 Then frmmain.select_by.AddItem frmmain.flx.Text
Next
frmmain.select_by.ListIndex = 1



'Me.Width = Screen.Width
'Me.Height = Screen.Height

 lpFormObj = ObjPtr(Me)




'SetProp frmmain.hwnd, "PrevWndProc", SetWindowLong(frmmain.hwnd, GWL_WNDPROC, AddressOf WndProc)


topRow = 1

Call update_customers

End Sub
'
Public Sub ScrollUp()
    ' scroll up..
    If topRow > 1 Then
        topRow = topRow - 1
        flx.topRow = topRow
    End If
End Sub

'//--[ScrollDown]---------------------------//
'
'  called from the MainModule WndProc sub
'  when a down-scrolling mouse message is
'  received
'
Public Sub ScrollDown()
    ' scroll down..
    If topRow < flx.Rows - 1 Then
        topRow = topRow + 1
        flx.topRow = topRow
    End If
End Sub

Public Sub view_cols()
flx.Rows = 1
flx.Row = 0
flx.Col = 1
flx.Text = "#"

flx.ColWidth(0) = 0
flx.ColWidth(1) = 400
flx.ColWidth(2) = 5000
flx.ColWidth(3) = 2000
flx.ColWidth(4) = 2000
flx.ColWidth(5) = 500

flx.Col = 2
flx.Text = "—ﬁ„ «·›« Ê—…"
flx.Col = 3
flx.Text = "«·⁄„Ì· "
flx.Col = 4
flx.Text = "«· «—ÌŒ"

flx.Col = 5
flx.Text = "«·„Ã„Ê⁄"






End Sub
Private Sub srch_cmd_Click()
On Error Resume Next

Select Case select_by.ListIndex
Case 0: str_order = "id"
Case 1: str_order = "serial"
Case 2: str_order = "customer"
Case 3: str_order = "date_txt"
End Select

If Option2.Value = True Then
str_desc = "DESC"
Else
str_desc = "ASC"
End If


sql_main = "select * from bills where"

If customer_id.ItemData(customer_id.ListIndex) <> 0 Then
sql_main = sql_main & "  customer=" & customer_id.ItemData(customer_id.ListIndex) & " and "
End If

sql_main = sql_main & " serial like '%" & serial_txt.Text & "%' and date_txt like '%" & date_txt.Text & "%'"



sql_main = sql_main & " order by " & str_order & " " & str_desc



Call show_records


End Sub
Private Sub update_customers()

customer_id.Clear

Set rsx = New Recordset

sql_cus = "select * from customers order by name asc"

rsx.Open sql_cus, DB, adOpenStatic, adLockOptimistic


For i = 1 To rsx.RecordCount

Call customer_id.AddItem(rsx!Name)
customer_id.ItemData(customer_id.ListCount - 1) = rsx!id

rsx.MoveNext

Next i
End Sub
Private Sub show_records(Optional Msg As Integer)
On Error Resume Next


Call update_customers

row_total.Text = "0"

Set rs = New Recordset
rs.Open sql_main, DB, adOpenStatic, adLockOptimistic

If rs.RecordCount = 0 Then
If Msg <> 1 Then
MsgBox "·«  ÊÃœ ‰ «∆Ã"
Else
flx.Clear
Call view_cols
End If

Else
flx.Rows = rs.RecordCount + 1

For i = 1 To rs.RecordCount
flx.Row = i
flx.Col = 0
flx.Text = rs!id
flx.Col = 1
flx.Text = i
flx.Col = 2
flx.Text = rs!serial

Set rs2 = New Recordset
sql_x = "select * from customers where id=" & rs!customer & ""

rs2.Open sql_x, DB, adOpenStatic, adLockOptimistic
flx.Col = 3
flx.Text = rs2!Name


flx.Col = 4
flx.Text = Format(rs!date_txt, "dd/mm/yyyy")



x_total = 0


Set rs3 = New Recordset

sql_x = "select * from records where bill=" & rs!id

rs3.Open sql_x, DB, adOpenStatic, adLockOptimistic

For i2 = 0 To rs3.RecordCount

x_total = x_total + (rs3!price * rs3!count_txt)


rs3.MoveNext

Next i2


flx.Col = 5
flx.Text = x_total

row_total.Text = Val(row_total.Text) + x_total

rs.MoveNext

Next i
End If
End Sub
Private Sub text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then srch_cmd_Click
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next

Select Case Button.Tag

Case "add"
FrmAdd.Show 1


Case "prices"

FrmPrices.Show 1

Case "edit"
Call show_edit


Case "del"
Call Del_Rec

Case "print"
Call Grid2HTML(flx, "File=c:\Cars_Gen.html", "<html dir=rtl><title>Generated Bill</title>", row_total.Text)
Shell "explorer.exe c:\Cars_Gen.html", vbNormalFocus


Case "refresh"
Call refresh_now

Case "customers"
Call FrmCustomers.Show(1)

Case "checks"
Call FrmChecks.Show(1)
End Select

End Sub
Private Sub Del_Rec()
If Trim(flx.TextMatrix(flx.Row, 0)) <> "" Then
id_txt = Trim(flx.TextMatrix(flx.Row, 0))

If MsgBox("Â· «‰  „ √ﬂœ „‰ √‰ﬂ  —Ìœ Õ–› «·”Ã· —ﬁ„ : " & Trim(flx.TextMatrix(flx.Row, 1)) & " ø ", vbQuestion + vbYesNo) = vbYes Then
Set rs = New Recordset
rs.Open "delete from bills where id=" & id_txt, DB
rs.Open "delete from records where bill=" & id_txt, DB
Call refresh_now
End If
End If
End Sub
Public Sub refresh_now()
If Trim(sql_main) <> "" Then
Call show_records(1)
End If
End Sub

Public Sub show_edit()


Set rs = New Recordset

sql_x = "select * from bills where id=" & Trim(flx.TextMatrix(flx.Row, 0))

rs.Open sql_x, DB, adOpenStatic, adLockOptimistic

'MsgBox rs.RecordCount

date_arr = Split(rs!date_txt, "/")


FrmEdit.bill_number.Text = rs!serial
FrmEdit.commentstxt.Text = rs!Comments

FrmEdit.daytxt.Text = date_arr(0)
FrmEdit.monthtxt.Text = date_arr(1)
FrmEdit.yeartxt.Text = date_arr(2)

FrmEdit.idtxt.Text = rs!id

x_customer = rs!customer

'---------------------------------------
FrmEdit.customer_id.Clear
x_customer = rs!customer
Set rs = New Recordset

sql_mainx = "select * from customers order by name asc"

rs.Open sql_mainx, DB, adOpenStatic, adLockOptimistic


For i = 1 To rs.RecordCount

Call FrmEdit.customer_id.AddItem(rs!Name)
FrmEdit.customer_id.ItemData(FrmEdit.customer_id.ListCount - 1) = rs!id

If rs!id = x_customer Then
FrmEdit.customer_id.ListIndex = (i - 1)
End If

rs.MoveNext

Next i

Call FrmEdit.Show(1)

End Sub
