VERSION 5.00
Begin VB.Form InterfaceWindow 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4650
   ClientLeft      =   -15
   ClientTop       =   570
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "Interface.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   StartUpPosition =   2  'CenterScreen
   Begin AnalogClockProgram.AnalogClockControl AnalogClock 
      Height          =   3960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3960
      _extentx        =   6985
      _extenty        =   6985
   End
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface window.
Option Explicit

'This procedure adjusts the window's size when it's activated.
Private Sub Form_Activate()
   Me.Width = AnalogClock.Width * Screen.TwipsPerPixelX
   Me.Height = AnalogClock.Height * Screen.TwipsPerPixelY
   Me.Width = Me.Width + (Me.Width - (Me.ScaleWidth * Screen.TwipsPerPixelX))
   Me.Height = Me.Height + (Me.Height - (Me.ScaleHeight * Screen.TwipsPerPixelY))
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
   With App
      Me.Caption = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
   End With
End Sub

'This procedure displays information about this program.
Private Sub InformationMenu_Click()
   MsgBox App.Comments, vbInformation, Me.Caption
End Sub

'This procedure closes this window.
Private Sub QuitMenu_Click()
   Unload Me
End Sub


