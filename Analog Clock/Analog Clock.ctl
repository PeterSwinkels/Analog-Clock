VERSION 5.00
Begin VB.UserControl AnalogClockControl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   KeyPreview      =   -1  'True
   ScaleHeight     =   192
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   Begin VB.PictureBox ClockBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.Timer AnalogClockTimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   120
      End
   End
End
Attribute VB_Name = "AnalogClockControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This control contains the procedures for drawing an analog clock.
Option Explicit
Private Const CLOCK_LINE_WIDTH As Long = 2            'Defines the width of the lines used to draw the clock.
Private Const CLOCK_SIZE As Long = 120                'Defines the clock face's diameter in pixels.
Private Const HAND_NUT_SIZE As Long = 3               'Defines the hands' nut size.
Private Const HOURS_TO_DEGREES As Long = 30           'Defines the value used to convert hours to degrees.
Private Const LARGE_MARK_INTERVAL As Long = 3         'Defines the interval between the large marks in hours.
Private Const MINUTES_TO_DEGREES As Long = 6          'Defines the value used to convert minutes to degrees.
Private Const MINUTES_TO_FRACTION As Double = 1 / 60  'Defines the value used to convert minutes to the fractional part of an hour.
Private Const NO_HOUR As Long = -1                    'Defines a value indicating "No hour.".
Private Const NO_MINUTE As Long = -1                  'Defines a value indicating "No minute.".
Private Const NO_SECOND As Long = -1                  'Defines a value indicating "No second.".
Private Const PI As Double = 3.14159265358979         'Defines the mathematical constant PI.
Private Const SECONDS_TO_DEGREES As Long = 6          'Defines the value used to convert seconds to degrees.
Private Const TWELVE_HOUR_ANGLE As Long = -90         'Defines the angle for noon/midnight in degrees.

Private Const CLOCK_X As Long = CLOCK_SIZE * 1.1                     'Defines the clock face's horizontal center in pixels.
Private Const CLOCK_Y As Long = CLOCK_SIZE * 1.1                     'Defines the clock face's vertical center in pixels.
Private Const DEGREES_PER_RADIAN As Double = 180 / PI                'Defines the number of degrees per radian.
Private Const HOUR_HAND_LENGTH As Long = CLOCK_SIZE / 1.6            'Defines the hour hand's length.
Private Const LARGE_MARK_LENGTH As Long = HOUR_HAND_LENGTH / 2.5     'Defines the size of the marking's used to mark every third hour.
Private Const MINUTE_HAND_LENGTH As Long = HOUR_HAND_LENGTH * 1.5    'Defines the minutes hand's length.
Private Const SECOND_HAND_LENGTH As Long = HOUR_HAND_LENGTH * 1.5    'Defines the seconds hand's length.
Private Const SMALL_MARK_LENGTH As Long = LARGE_MARK_LENGTH / 2      'Defines the size of the marking's used to mark the hours.

'This structure defines the time displayed by the clock.
Private Type TimeStr
   Hour As Long    'Contains the hour.
   Minute As Long  'Contains the minute.
   Second As Long  'Contains the second.
End Type

'This procedure manages the time displayed by the clock.
Private Function CurrentTime(Optional Advance As Boolean = False, Optional NewHour As Long = NO_HOUR, Optional NewMinute As Long = NO_MINUTE, Optional NewSecond As Long = NO_SECOND) As TimeStr
Static TimeV As TimeStr

   With TimeV
      If Advance Then
         If .Second = 59 Then
            .Second = 0
            If .Minute = 59 Then
               .Minute = 0
               If .Hour = 11 Then
                  .Hour = 0
               Else
                  .Hour = .Hour + 1
               End If
            Else
               .Minute = .Minute + 1
            End If
         Else
            .Second = .Second + 1
         End If
      Else
         If Not NewHour = NO_HOUR Then .Hour = NewHour
         If Not NewMinute = NO_MINUTE Then .Minute = NewMinute
         If Not NewSecond = NO_SECOND Then .Second = NewSecond
      End If
   End With
   
   CurrentTime = TimeV
End Function


'This procedure draws an analog clock displaying the specified time.
Private Sub DrawClock(DisplayedTime As TimeStr)
Dim HourAsRadians As Double
Dim HourOnFace As Long
Dim MarkLength As Long
Dim MinuteAsRadians As Double
Dim SecondAsRadians As Double

Static HourHandX As Long
Static HourHandY As Long
Static MinuteHandX As Long
Static MinuteHandY As Long
Static SecondHandX As Long
Static SecondHandY As Long

   ClockBox.Line (CLOCK_X, CLOCK_Y)-(HourHandX, HourHandY), ClockBox.BackColor
   ClockBox.Line (CLOCK_X, CLOCK_Y)-(MinuteHandX, MinuteHandY), ClockBox.BackColor
   ClockBox.Line (CLOCK_X, CLOCK_Y)-(SecondHandX, SecondHandY), ClockBox.BackColor
   
   For HourOnFace = 0 To 11
      HourAsRadians = ((HourOnFace * HOURS_TO_DEGREES) + TWELVE_HOUR_ANGLE) / DEGREES_PER_RADIAN
      If HourOnFace Mod LARGE_MARK_INTERVAL = 0 Then MarkLength = LARGE_MARK_LENGTH Else MarkLength = SMALL_MARK_LENGTH
      ClockBox.Line ((Cos(HourAsRadians) * CLOCK_SIZE) + CLOCK_X, (Sin(HourAsRadians) * CLOCK_SIZE) + CLOCK_Y)-((Cos(HourAsRadians) * (CLOCK_SIZE - MarkLength)) + CLOCK_X, (Sin(HourAsRadians) * (CLOCK_SIZE - MarkLength)) + CLOCK_Y), vbYellow
   Next HourOnFace
   ClockBox.Circle (CLOCK_X, CLOCK_Y), CLOCK_SIZE, vbBlue
   
   With DisplayedTime
      HourAsRadians = (((.Hour + .Minute * MINUTES_TO_FRACTION) * HOURS_TO_DEGREES) + TWELVE_HOUR_ANGLE) / DEGREES_PER_RADIAN
      SecondAsRadians = ((.Second * SECONDS_TO_DEGREES) + TWELVE_HOUR_ANGLE) / DEGREES_PER_RADIAN
      MinuteAsRadians = ((.Minute * MINUTES_TO_DEGREES) + TWELVE_HOUR_ANGLE) / DEGREES_PER_RADIAN
   End With
   
   HourHandX = (Cos(HourAsRadians) * HOUR_HAND_LENGTH) + CLOCK_X
   HourHandY = (Sin(HourAsRadians) * HOUR_HAND_LENGTH) + CLOCK_Y
   MinuteHandX = (Cos(MinuteAsRadians) * MINUTE_HAND_LENGTH) + CLOCK_X
   MinuteHandY = (Sin(MinuteAsRadians) * MINUTE_HAND_LENGTH) + CLOCK_Y
   SecondHandX = (Cos(SecondAsRadians) * SECOND_HAND_LENGTH) + CLOCK_X
   SecondHandY = (Sin(SecondAsRadians) * SECOND_HAND_LENGTH) + CLOCK_Y
   
   ClockBox.Line (CLOCK_X, CLOCK_Y)-(HourHandX, HourHandY), vbGreen
   ClockBox.Line (CLOCK_X, CLOCK_Y)-(MinuteHandX, MinuteHandY), vbGreen
   ClockBox.Line (CLOCK_X, CLOCK_Y)-(SecondHandX, SecondHandY), vbRed
   ClockBox.Circle (CLOCK_X, CLOCK_Y), HAND_NUT_SIZE, vbWhite
End Sub

'This procedure returns the hour displayed on the clock face at the specified position.
Private Function GetHour(X As Long, Y As Long) As Long
Dim HourAsRadians As Double
Dim HourAtPosition As Long
Dim HourOnFace As Long
Dim HourSize As Long
Dim HourX As Long
Dim HourY As Long

   HourAtPosition = NO_HOUR
   HourSize = (CLOCK_SIZE * PI) / 12
   For HourOnFace = 0 To 11
      HourAsRadians = ((HourOnFace * HOURS_TO_DEGREES) + TWELVE_HOUR_ANGLE) / DEGREES_PER_RADIAN
      HourX = (Cos(HourAsRadians) * CLOCK_SIZE) + CLOCK_X
      HourY = (Sin(HourAsRadians) * CLOCK_SIZE) + CLOCK_Y
   
      If X >= HourX - HourSize And X <= HourX + HourSize And Y >= HourY - HourSize And Y <= HourY + HourSize Then
         HourAtPosition = HourOnFace
         Exit For
      End If
   Next HourOnFace
   
   GetHour = HourAtPosition
End Function

'This procedure returns the minute displayed on the clock face at the specified position.
Private Function GetMinute(X As Long, Y As Long) As Long
Dim MinuteAsRadians As Double
Dim MinuteAtPosition As Long
Dim MinuteOnFace As Long
Dim MinuteSize As Long
Dim MinuteX As Long
Dim MinuteY As Long

   MinuteAtPosition = NO_MINUTE
   MinuteSize = (CLOCK_SIZE * PI) / 60
   For MinuteOnFace = 0 To 59
      MinuteAsRadians = ((MinuteOnFace * MINUTES_TO_DEGREES) + TWELVE_HOUR_ANGLE) / DEGREES_PER_RADIAN
      MinuteX = (Cos(MinuteAsRadians) * CLOCK_SIZE) + CLOCK_X
      MinuteY = (Sin(MinuteAsRadians) * CLOCK_SIZE) + CLOCK_Y
   
      If X >= MinuteX - MinuteSize And X <= MinuteX + MinuteSize And Y >= MinuteY - MinuteSize And Y <= MinuteY + MinuteSize Then
         MinuteAtPosition = MinuteOnFace
         Exit For
      End If
   Next MinuteOnFace
   
   GetMinute = MinuteAtPosition
End Function

'This procedure gives the command to display an analog clock.
Private Sub AnalogClockTimer_Timer()
   DrawClock CurrentTime(Advance:=True)
End Sub

'This procedure gives the command to change the time being displayed.
Private Sub ClockBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Button
      Case vbMiddleButton
         CurrentTime , NewHour:=GetHour(CLng(X), CLng(Y)), NewMinute:=0
      Case vbLeftButton
         CurrentTime , NewHour:=GetHour(CLng(X), CLng(Y))
      Case vbRightButton
         CurrentTime , , NewMinute:=GetMinute(CLng(X), CLng(Y))
   End Select
   
   DrawClock CurrentTime(Advance:=False)
End Sub

'This procedure initializes this control.
Private Sub UserControl_Initialize()
   CurrentTime , NewHour:=Hour(Time$()), NewMinute:=Minute(Time$()), NewSecond:=Second(Time$())
   DrawClock CurrentTime(Advance:=False)
   AnalogClockTimer.Enabled = True
   
   With ClockBox
      .DrawWidth = CLOCK_LINE_WIDTH
      .FillStyle = vbFSTransparent
      .ToolTipText = "Click near the face's edge or use the plus key to set the time."
   End With
End Sub

'This procedure handles the user's keystrokes.
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TimeV As TimeStr

   If KeyCode = vbKeyAdd Then
      TimeV = CurrentTime()
      With TimeV
         If Shift And vbShiftMask = vbShiftMask Then
            If .Hour = 11 Then .Hour = 0 Else .Hour = .Hour + 1
         Else
            If .Minute = 59 Then
               .Minute = 0
               If .Hour = 11 Then .Hour = 0 Else .Hour = .Hour + 1
            Else
               .Minute = .Minute + 1
            End If
         End If
         DrawClock CurrentTime(, NewHour:=.Hour, NewMinute:=.Minute)
      End With
   End If
End Sub

'This procedure adjusts this window to its new size.
Private Sub UserControl_Resize()
   ClockBox.Width = ScaleWidth
   ClockBox.Height = ScaleHeight
End Sub

'This procedure sets this control's size.
Private Sub UserControl_Show()
   Width = (CLOCK_SIZE * 2.2) * Screen.TwipsPerPixelX
   Height = (CLOCK_SIZE * 2.2) * Screen.TwipsPerPixelY
End Sub


