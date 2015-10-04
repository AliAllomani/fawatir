VERSION 5.00
Begin VB.UserControl RgheedFixedControl 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   20
      X2              =   5
      Y1              =   10
      Y2              =   20
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   10
      X2              =   20
      Y1              =   10
      Y2              =   20
   End
End
Attribute VB_Name = "RgheedFixedControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const mdef_FixRight = False
Const mdef_FixLeft = True
Const mdef_FixTop = True
Const mdef_FixBottom = False
'Property Variables:
Dim mFixed As Control
Dim mFixRight As Boolean
Dim mFixLeft As Boolean
Dim mFixTop As Boolean
Dim mFixBottom As Boolean

Dim mTop    As Integer
Dim mLeft   As Integer
Dim mRight  As Integer
Dim mBottom As Integer

Private Sub UserControl_Resize()
    Dim H As Integer
    Dim W As Integer
    
    H = UserControl.ScaleHeight
    W = UserControl.ScaleWidth
    
    Line1.X1 = 0: Line1.X2 = W
    Line1.Y1 = 0: Line1.Y2 = H
    
    Line2.X1 = W: Line2.X2 = 0
    Line2.Y1 = 0: Line2.Y2 = H
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=15,0,2,0
Public Property Get FixedControl() As Object
Attribute FixedControl.VB_Description = "«·⁄‰’— «·–Ì  —Ìœ  À»ÌÀÂ"
Attribute FixedControl.VB_MemberFlags = "400"
    Set FixedControl = mFixed
End Property

Public Property Set FixedControl(ByVal New_FixedControl As Control)
    If Ambient.UserMode = False Then Err.Raise 383
    Set mFixed = New_FixedControl
    PropertyChanged "FixedControl"

    GetFixedborders
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get FixRight() As Boolean
Attribute FixRight.VB_ProcData.VB_Invoke_Property = ";0-—€Ìœ"
    FixRight = mFixRight
End Property

Public Property Let FixRight(ByVal New_FixRight As Boolean)
    mFixRight = New_FixRight
    PropertyChanged "FixRight"
    
    GetFixedborders
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get FixLeft() As Boolean
Attribute FixLeft.VB_ProcData.VB_Invoke_Property = ";0-—€Ìœ"
    FixLeft = mFixLeft
End Property

Public Property Let FixLeft(ByVal New_FixLeft As Boolean)
    mFixLeft = New_FixLeft
    PropertyChanged "FixLeft"
    
    GetFixedborders
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,True
Public Property Get FixTop() As Boolean
Attribute FixTop.VB_ProcData.VB_Invoke_Property = ";0-—€Ìœ"
    FixTop = mFixTop
End Property

Public Property Let FixTop(ByVal New_FixTop As Boolean)
    mFixTop = New_FixTop
    PropertyChanged "FixTop"
    
    GetFixedborders
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get FixBottom() As Boolean
Attribute FixBottom.VB_ProcData.VB_Invoke_Property = ";0-—€Ìœ"
    FixBottom = mFixBottom
End Property

Public Property Let FixBottom(ByVal New_FixBottom As Boolean)
    mFixBottom = New_FixBottom
    PropertyChanged "FixBottom"
    
    GetFixedborders
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    mFixRight = mdef_FixRight
    mFixLeft = mdef_FixLeft
    mFixTop = mdef_FixTop
    mFixBottom = mdef_FixBottom
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set mFixed = PropBag.ReadProperty("FixedControl", Nothing)
    mFixRight = PropBag.ReadProperty("FixRight", mdef_FixRight)
    mFixLeft = PropBag.ReadProperty("FixLeft", mdef_FixLeft)
    mFixTop = PropBag.ReadProperty("FixTop", mdef_FixTop)
    mFixBottom = PropBag.ReadProperty("FixBottom", mdef_FixBottom)
End Sub

Private Sub UserControl_Show()
 '   If UserControl.Ambient.UserMode Then MsgBox "Yes"
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("FixedControl", mFixed, Nothing)
    Call PropBag.WriteProperty("FixRight", mFixRight, mdef_FixRight)
    Call PropBag.WriteProperty("FixLeft", mFixLeft, mdef_FixLeft)
    Call PropBag.WriteProperty("FixTop", mFixTop, mdef_FixTop)
    Call PropBag.WriteProperty("FixBottom", mFixBottom, mdef_FixBottom)
End Sub

Public Sub Reset()
    Dim nT As Integer   ' New Top
    Dim nL As Integer   ' New Left
    Dim nW As Integer   ' New Width
    Dim nH As Integer   ' New Height
    Dim pW As Integer   ' Parent Width
    Dim pH As Integer   ' Parent Height
    Dim fW As Integer   ' Fixed Width
    Dim fH As Integer   ' Fixed Height
    
    On Error Resume Next
    
    pW = UserControl.Parent.Width
    pH = UserControl.Parent.Height
    fW = mFixed.Width
    fH = mFixed.Height
    
    nL = IIf(mFixRight And Not mFixLeft, pW - fW - mRight, mLeft)
    nT = IIf(mFixBottom And Not mFixTop, pH - fH - mBottom, mTop)
    nW = IIf(mFixLeft And mFixRight, pW - mLeft - mRight, fW)
    nH = IIf(mFixTop And mFixBottom, pH - mTop - mBottom, fH)
        
    mFixed.Move nL, nT, nW, nH
End Sub

Private Sub GetFixedborders()
    If Not (mFixed Is Nothing) Then
       mTop = mFixed.Top
       mLeft = mFixed.Left
       mRight = UserControl.Parent.Width - mFixed.Width - mLeft
       mBottom = UserControl.Parent.Height - mFixed.Height - mTop
    End If
End Sub
