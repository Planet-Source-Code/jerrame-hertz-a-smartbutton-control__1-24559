VERSION 5.00
Begin VB.UserControl SmartButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   ScaleHeight     =   1665
   ScaleWidth      =   5205
   ToolboxBitmap   =   "SmartButton.ctx":0000
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   0
      Width           =   615
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "SmartButton.ctx":0312
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SmartButton"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
   End
End
Attribute VB_Name = "SmartButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' I kept this simple so beginners could understanding how it works
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Ret As Long
Private Clicked As Boolean
Private FlagInside As Boolean

Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y) ' send event to usercontrol
End Sub

Private Sub Label1_Click()
    RaiseEvent Click ' Event Click()
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y) ' send event to usercontrol
End Sub

Private Sub Picture1_Click()
    RaiseEvent Click ' Event Click()
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click ' Event Click()
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y) ' Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clicked = True
    DrawButtonDown
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y) ' MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Clicked Then
        FlagInside = False
        UserControl_MouseOut Button, Shift, X, Y
    End If
End Sub

Function UserControl_MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseOut(Button, Shift, X, Y) ' MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FlagInside = False
    If X < 0 Or X > UserControl.Width Or Y < 0 Or Y > UserControl.Height Then
        FlagInside = False ' set mouse is outside
        Ret = ReleaseCapture()
        Cls
    Else
        If FlagInside = False Then
            FlagInside = True ' set mouse is inside
            Ret = SetCapture(UserControl.hwnd)
            DrawButtonUp
        End If
    End If
End Function

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y) ' MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clicked = False ' set mouse as up
    UserControl_MouseOut Button, Shift, X, Y ' UserControl_MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub

Friend Sub DrawButtonUp()
    Cls ' clear UserControl
    Line (0, 0)-(ScaleWidth, 0), vbWhite ' Line fron top-left to top-right
    Line (0, 0)-(0, ScaleHeight), vbWhite ' Line from top-left to bottom-left
    Line (Width - 20, 0)-(Width - 20, ScaleHeight), vbBlack ' Line from top-right to bottom-right
    Line (0, Height - 20)-(Width - 20, Height - 20), vbBlack ' Line from bottom-left to bottom-right
End Sub

Friend Sub DrawButtonDown()
    Cls ' clear UserControl
    Line (0, 0)-(ScaleWidth, 0), vbBlack ' Line fron top-left to top-right
    Line (0, 0)-(0, ScaleHeight), vbBlack ' Line from top-left to bottom-left
    Line (Width - 20, 0)-(Width - 20, ScaleHeight), vbWhite ' Line from top-right to bottom-right
    Line (0, Height - 20)-(Width - 20, Height - 20), vbWhite ' Line from bottom-left to bottom-right
End Sub

Private Sub UserControl_Resize()
    ' make sure usercontrol is larger then label
    If ScaleHeight < Label1.Height + 100 Then
        Height = Label1.Height + 100
    End If
    ' position the picturebox
    Picture1.Top = 20
    Picture1.Left = 20
    Picture1.Width = ScaleWidth - 40
    Picture1.Height = (ScaleHeight - 40) - Label1.Height
    ' center the image in the picturebox
    Image1.Top = (Picture1.ScaleHeight / 2) - (Image1.Height / 2)
    Image1.Left = (Picture1.ScaleWidth / 2) - (Image1.Width / 2)
    ' position the label
    Label1.Width = (ScaleWidth / 3) * 2
    Label1.Top = Picture1.Top + Picture1.Height
    Label1.Left = (ScaleWidth / 2) - (Label1.Width / 2)
    ' only caption if not tall enough for image to show
    If Picture1.Height < Image1.Height Then Image1.Visible = False Else Image1.Visible = True
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Label1.Caption = PropBag.ReadProperty("Caption", "Label1")
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Label1")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Image1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Image1.Picture = New_Picture
    PropertyChanged "Picture"
End Property

