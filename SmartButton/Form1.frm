VERSION 5.00
Object = "*\ASmartButtonControl.vbp"
Begin VB.Form Form1 
   Caption         =   "SmartButton - DEMO"
   ClientHeight    =   1245
   ClientLeft      =   7125
   ClientTop       =   4215
   ClientWidth     =   5880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   5880
   Begin SmartButtonControl.SmartButton SmartButton3 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      Caption         =   "SmartButton"
      Picture         =   "Form1.frx":030A
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1800
      ScaleHeight     =   945
      ScaleWidth      =   2145
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin SmartButtonControl.SmartButton SmartButton1 
         Height          =   360
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         Caption         =   "1"
         Picture         =   "Form1.frx":0624
      End
      Begin SmartButtonControl.SmartButton SmartButton1 
         Height          =   360
         Index           =   1
         Left            =   735
         TabIndex        =   3
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         Caption         =   "2"
         Picture         =   "Form1.frx":093E
      End
      Begin SmartButtonControl.SmartButton SmartButton1 
         Height          =   360
         Index           =   2
         Left            =   1110
         TabIndex        =   4
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         Caption         =   "3"
         Picture         =   "Form1.frx":0C58
      End
      Begin SmartButtonControl.SmartButton SmartButton1 
         Height          =   360
         Index           =   3
         Left            =   1485
         TabIndex        =   5
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         Caption         =   "4"
         Picture         =   "Form1.frx":0F72
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   280
         Top             =   300
         Width           =   1650
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   0
         Left            =   480
         Top             =   120
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   1
         Left            =   855
         Top             =   120
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   2
         Left            =   1230
         Top             =   120
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   3
         Left            =   1605
         Top             =   120
         Width           =   135
      End
   End
   Begin SmartButtonControl.SmartButton SmartButton2 
      Height          =   975
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      Caption         =   "EXIT"
      Picture         =   "Form1.frx":128C
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' see usercontrol for comments
Option Explicit
Dim Clicked As Boolean

Private Sub SmartButton1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clicked = True
    Shape2(Index).BackColor = &HFF0000
End Sub

Private Sub SmartButton1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Clicked Then
        Shape2(Index).BackColor = &HFF00&
    Else
        Shape2(Index).BackColor = &HFF0000
    End If
End Sub

Private Sub SmartButton1_MouseOut(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or X > SmartButton1(Index).Width Or Y < 0 Or Y > SmartButton1(Index).Height Then
        Shape2(Index).BackColor = &HFF&
    End If
End Sub

Private Sub SmartButton1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clicked = False
End Sub

Private Sub SmartButton2_Click()
    End
End Sub
