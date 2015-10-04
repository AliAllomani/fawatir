VERSION 5.00
Begin VB.Form FrmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ»«⁄…"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   3915
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   3615
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄‰Ê«‰ : "
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
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ «»… „ﬂ«‰ «· Ê«Ãœ ﬂ⁄‰’— ›Ì «·ÃœÊ·"
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
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ «»… „ﬂ«‰ «· Ê«Ãœ ﬂ⁄‰Ê«‰ ··’›Õ… "
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
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„Ê«›ﬁ"
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
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Line Line3 
      X1              =   3840
      X2              =   3360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   3840
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   3840
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "FrmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
Set rs = New Recordset
rs.Open sql_main, DB
drp.Sections("section2").Controls("header").Caption = frmmain.flx.TextMatrix(1, 6)
Set drp.DataSource = rs
Unload Me
drp.Show

Else
Set rs = New Recordset
rs.Open sql_main, DB
If Trim(Text1.text) <> "" Then
drp2.Sections("section2").Controls("Shape1").BackColor = "&H00E0E0E0"
drp2.Sections("section2").Controls("Shape1").BorderColor = vbBlack
drp2.Sections("section2").Controls("header").Caption = Text1.text
Else
drp2.Sections("section2").Controls("Shape1").BackColor = vbWhite
drp2.Sections("section2").Controls("Shape1").BorderColor = vbWhite
End If

Set drp2.DataSource = rs
Unload Me
drp2.Show
End If
End Sub

Private Sub Form_Load()
Me.Icon = frmmain.Icon
End Sub
