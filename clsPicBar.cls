VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPicBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'  The Graphical Progress Bar by Brian Middleton
'  Thanks to Todd Spangler for his example of animating a graphic.
'  His example can be found on Planet Source Code:
'     http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=3636
'
'  To use this class add it to your project.  Create a PictureBox on your form
'  and choose the picture for it.  Then create a varaible for it:
'        Private (or public, etc...) someVarName as clsPicBar
'
'  Then initalize it in the form load or wherever you want:
'     Set someVarName = New clsPicBar
'     With someVarName
'        Set .PicBox = varNameForPictureBoxControl
'        .Incrementor = 5
'     End With
'
'  See frmExample to see actual code.
'********************************************************************************

Option Explicit

Public Enum DirType
   GOLEFT = -1
   GORIGHT = 1
End Enum

Private m_AutoDraw As Boolean
Private m_Increment As Integer
Private m_pic As PictureBox
Private m_Direction As DirType
Private A As Integer
Private B As Integer

'initializes any variables that you want when an instance is created
Private Sub Class_Initialize()
   m_Direction = GORIGHT
End Sub

'Retrieves the which PictureBox object is used.
Public Property Get PicBox() As PictureBox
   Set PicBox = m_pic
End Property

'sets the PictureBox object to be manipulated.  AutoRedraw should be True and
'the ScaleMode must be 3-Pixel
Public Property Set PicBox(somePic As PictureBox)
   Set m_pic = somePic
   m_pic.AutoRedraw = True
   m_pic.ScaleMode = 3 'pixel mode, must be set to pixel mode
End Property

'returns the value of AutoRedraw
Public Property Get ReDraw() As Boolean
   ReDraw = m_AutoDraw
End Property

'allows user to alter auto redraw option
Public Property Let ReDraw(bTmp As Boolean)
   m_AutoDraw = bTmp
End Property

'retrieves the increment value
Public Property Get Incrementor() As Integer
   Incrementor = m_Increment
End Property

'increase or decrease how many pixels to move the picture
'the higher the number the faster it moves
Public Property Let Incrementor(iTmp As Integer)
   m_Increment = iTmp
End Property

'retrieves which direction the graphic should be moving
Public Property Get PicDir() As DirType
   PicDir = m_Direction
End Property

'set whether or not you want the graghic to move left or right
Public Property Let PicDir(dTmp As DirType)
   m_Direction = dTmp
End Property

'Scrolls a picture
'Call this Sub each time you would normally increment a progress bar
Public Sub BGScroll()
   m_pic.Picture = m_pic.Image
   If m_Direction = GORIGHT Then
      m_pic.PaintPicture m_pic.Picture, A + Incrementor, B
      m_pic.PaintPicture m_pic.Picture, 0, 0, , , m_pic.ScaleWidth - Incrementor
   Else
      m_pic.PaintPicture m_pic.Picture, A - Incrementor, B
      m_pic.PaintPicture m_pic.Picture, m_pic.ScaleWidth - Incrementor, 0
   End If
End Sub

