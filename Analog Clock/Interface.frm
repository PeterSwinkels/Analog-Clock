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
      _ExtentX        =   6985
      _ExtentY        =   6985
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

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap

   With Me
      .Width = AnalogClock.Width * Screen.TwipsPerPixelX
      .Height = AnalogClock.Height * Screen.TwipsPerPixelY
      .Width = .Width + (.Width - (.ScaleWidth * Screen.TwipsPerPixelX))
      .Height = .Height + (.Height - (.ScaleHeight * Screen.TwipsPerPixelY))
      .Caption = ProgramInformation()
   End With
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap

   MsgBox App.Comments, vbInformation, ProgramInformation()
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure closes this window.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   
   Unload Me
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub


