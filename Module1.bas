Attribute VB_Name = "Module1"
'=========================================================='
'         This Project Programmed By : Ali Allomani        '
'                  info@allomani.com		                '
'=========================================================='


Public MyDBF As String
Public sql_main As String
Public MyPath As String

Sub Main()

MyPath = App.Path & "\"
 MyDBF = App.Path & "\" & "database.mdb"
 
 If Dir(MyDBF) <> "" Then

Call OpenDB

If Command$ = "nonpass" Then
frmmain.Show
Else
FrmLogin.Show
'frmmain.Show

End If

  Else
    MsgBox "·„ Ì „ «·⁄ÀÊ— ⁄·Ï „·› ﬁ«⁄œ… «·»Ì«‰«  ! ", vbExclamation, "Œÿ√  Õ„Ì·"
    End If
    
End Sub

