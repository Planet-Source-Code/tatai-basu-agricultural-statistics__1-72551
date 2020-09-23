VERSION 5.00
Begin VB.UserControl UPB 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   ScaleHeight     =   17
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   Begin VB.Timer Animate 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   225
      Top             =   615
   End
End
Attribute VB_Name = "UPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Unknown Progress Bar v1.2.1 (User Control)
' Ed Wilk, 04/20/2007
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, _
    ByVal edge As Long, ByVal grfFlags As Long) As Long

Private rc As RECT

Public Enum Border
    Flat = 6
    Raised = 9
End Enum

Private mBorders        As Border
Private mEnabled        As Boolean
Private mForeSpeed      As Integer
Private mDelay          As Integer
Private mForeColor      As Long
Private mFollowColor    As Long
Private mVerticalColors As Boolean
Private mBlocks         As Boolean
Private mBlockSize      As Integer
Private mBlockSpace     As Integer
Private Head            As Long
Private Tail            As Long

Private Type GRADIENT_RECT
    UpperLeft  As Long
    LowerRight As Long
End Type

Private Type TRIVERTEX
    X     As Long
    Y     As Long
    Red   As Integer
    Green As Integer
    Blue  As Integer
    Alpha As Integer
End Type

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" _
    (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, _
     pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
     
Public Property Get BlockSpace() As Integer
     BlockSpace = mBlockSpace
End Property
Property Let BlockSpace(ByVal NewValue As Integer)
    
    If NewValue < 1 Then
        NewValue = 1
    ElseIf NewValue + 1 > mBlockSize Then
        NewValue = mBlockSize - 2
    End If
       
    mBlockSpace = NewValue
    
End Property
Public Property Get BlockSize() As Integer
     BlockSize = mBlockSize
End Property

Property Let BlockSize(ByVal NewValue As Integer)
    
    If NewValue < 4 Then
        NewValue = 4
    ElseIf NewValue > UserControl.ScaleWidth / 4 Then
        NewValue = UserControl.ScaleWidth / 4
    End If
    
    mBlockSize = NewValue
    If mBlockSpace + 1 > mBlockSize Then mBlockSpace = mBlockSize - 2
    
End Property
Public Property Get ForeSpeed() As Integer
     ForeSpeed = mForeSpeed
End Property

Public Property Get Delay() As Integer
    Delay = mDelay
End Property
Public Property Let Borders(ByVal NewValue As Border)
    mBorders = NewValue
    PropertyChanged "Borders"
End Property

Public Property Get Borders() As Border
    Borders = mBorders
End Property

Public Property Get Blocks() As Boolean
     Blocks = mBlocks
End Property

Public Property Let Blocks(ByVal NewValue As Boolean)
     mBlocks = NewValue
    PropertyChanged "Blocks"
End Property

Private Sub DrawBar(H As Double, Head As Long, Tail As Long, mHDC As Long, _
ByVal iColor1 As Long, ByVal iColor2 As Long, Solid As Boolean, Blocks As Boolean)

Dim gRect As GRADIENT_RECT
Dim Horiz(1) As TRIVERTEX
Dim Y As Integer
Dim S As Long
    
 
 If Solid = True Then S = 1
    
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
  
    With Horiz(0)
        .X = Head
        .Y = 0
    End With
        
    With Horiz(1)
        .X = Tail
        .Y = H
    End With
     
        
' Draw gradient (from head to tail)
    GradientFillColor Horiz(0), iColor2
    GradientFillColor Horiz(1), iColor1
    GradientFillRect mHDC, Horiz(0), 2, gRect, 1, S

    If Blocks = True Then
' Draw white blocks
        For Y = 0 To Head Step mBlockSize
            UserControl.Line (Y, H)-(Y + mBlockSpace, 0), vbWhite, BF
        Next Y
    End If

' draw selected border on all sides
    DrawEdge mHDC, rc, mBorders, 15

End Sub

Private Sub GradientFillColor(ByRef tTV As TRIVERTEX, ByVal iColor As Long)

    Dim iRed   As Long
    Dim iGreen As Long
    Dim iBlue  As Long

    '/* Separate color into RGB
    iRed = (iColor And &HFF&) * &H100&
    iGreen = (iColor And &HFF00&)
    iBlue = (iColor And &HFF0000) \ &H100&
    
    '/* Make Red color a UShort
    If (iRed And &H8000&) = &H8000& Then
       tTV.Red = (iRed And &H7F00&)
       tTV.Red = tTV.Red Or &H8000
    Else
       tTV.Red = iRed
    End If
    '/* Make Green color a UShort
    If (iGreen And &H8000&) = &H8000& Then
       tTV.Green = (iGreen And &H7F00&)
       tTV.Green = tTV.Green Or &H8000
    Else
       tTV.Green = iGreen
    End If
    '/* Make Blue color a UShort
    If (iBlue And &H8000&) = &H8000& Then
       tTV.Blue = (iBlue And &H7F00&)
       tTV.Blue = tTV.Blue Or &H8000
    Else
       tTV.Blue = iBlue
    End If
    
    tTV.Alpha = 0

End Sub

Public Property Get Enabled() As Boolean
     Enabled = mEnabled
End Property

Public Property Get VerticalColors() As Boolean
     VerticalColors = mVerticalColors
End Property

Public Property Let VerticalColors(ByVal NewValue As Boolean)

     mVerticalColors = NewValue
    PropertyChanged "VerticalColors"
    
End Property

Property Let ForeSpeed(ByVal S As Integer)
    
    If S < 1 Then
        S = 1
    ElseIf S > 60 Then
        S = 60
    End If
       
    mForeSpeed = S
    
End Property

Property Let Delay(ByVal NewValue As Integer)
    
    If NewValue < 1 Then
        NewValue = 1
    ElseIf NewValue > 60 Then
        NewValue = 60
    End If
    
    mDelay = NewValue
    Animate.Interval = NewValue
    
End Property

Property Let Reset(ByVal Yes As Boolean)
    Head = 0
    Tail = 0
    UserControl.Cls
    DrawEdge UserControl.hdc, rc, mBorders, 15
End Property

Public Property Get ForeColor() As OLE_COLOR
     ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    mForeColor = NewValue
    PropertyChanged "ForeColor"
End Property

Public Property Get FollowColor() As OLE_COLOR
     FollowColor = mFollowColor
End Property

Public Property Let FollowColor(ByVal NewValue As OLE_COLOR)
    mFollowColor = NewValue
    PropertyChanged "FollowColor"
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    Animate.Enabled = NewValue
    mEnabled = NewValue
    UserControl.Cls
    DrawEdge UserControl.hdc, rc, mBorders, 15
End Property

Private Sub Animate_Timer()

Static Foward As Boolean
        
' draw progress bar back and forth
' First draft, could probably use some tweaking!

If Foward = False Then
    If Head < UserControl.ScaleWidth Then
        Head = Head + mForeSpeed
        UserControl.Cls
        DrawBar ScaleHeight, Head, Tail, hdc, mFollowColor, mForeColor, mVerticalColors, mBlocks
        Exit Sub
    End If
       
    If Head >= UserControl.ScaleWidth And Tail < UserControl.ScaleWidth Then
        Tail = Tail + mForeSpeed
        UserControl.Cls
        DrawBar ScaleHeight, Head, Tail, hdc, mFollowColor, mForeColor, mVerticalColors, mBlocks
    Else
        Foward = True
        Exit Sub
    End If
     
Else

    If Head > 0 And Tail > 0 Then
        Tail = Tail - mForeSpeed
        UserControl.Cls
        DrawBar ScaleHeight, Head, Tail, hdc, mForeColor, mFollowColor, mVerticalColors, mBlocks
        Exit Sub
    End If
        
    If Tail <= 0 And Head > 0 Then
        Head = Head - mForeSpeed
    Else
        Foward = False
        Exit Sub
    End If
        UserControl.Cls
        DrawBar ScaleHeight, Head, Tail, hdc, mForeColor, mFollowColor, mVerticalColors, mBlocks
   End If
 
End Sub


Private Sub UserControl_InitProperties()
    
    Animate.Enabled = False
    mDelay = 16
    mForeSpeed = 8
    Head = 0
    Tail = 0
    
    mForeColor = &H80FF&     ' orange
    mFollowColor = vbWhite
    mVerticalColors = False
    
    mBlocks = True
    mBlockSize = 8
    mBlockSpace = 1
    
    mBorders = Flat
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ForeColor = PropBag.ReadProperty("ForeColor", 0)
    FollowColor = PropBag.ReadProperty("FollowColor", 0)
    VerticalColors = PropBag.ReadProperty("VerticalColors", False)
    Blocks = PropBag.ReadProperty("Blocks", True)
    Borders = PropBag.ReadProperty("Borders", 0)
    Delay = PropBag.ReadProperty("Delay", 16)
    ForeSpeed = PropBag.ReadProperty("ForeSpeed", 6)
    BlockSize = PropBag.ReadProperty("BlockSize", 8)
    BlockSpace = PropBag.ReadProperty("BlockSpace", 1)
    
End Sub

Private Sub UserControl_Resize()
    
    rc.Right = UserControl.ScaleWidth
    rc.Bottom = UserControl.ScaleHeight
    UserControl.Cls
    DrawEdge UserControl.hdc, rc, mBorders, 15

End Sub

Private Sub UserControl_Terminate()
    Animate.Enabled = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    PropBag.WriteProperty "ForeColor", mForeColor, 0
    PropBag.WriteProperty "FollowColor", mFollowColor, 0
    PropBag.WriteProperty "VerticalColors", mVerticalColors, False
    PropBag.WriteProperty "Blocks", mBlocks, True
    PropBag.WriteProperty "Borders", mBorders, 0
    PropBag.WriteProperty "Delay", mDelay, 16
    PropBag.WriteProperty "ForeSpeed", mForeSpeed, 6
    PropBag.WriteProperty "BlockSize", mBlockSize, 8
    PropBag.WriteProperty "BlockSpace", mBlockSpace, 1
    
    
End Sub
