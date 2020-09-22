VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   FillColor       =   &H00404000&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   72
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   572
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Colin Woor
'Email: colin@woor.co.uk
'
'Hello, if you have any comments etc. then feel free to email me.
'If you use any of this code, then please mention me in your code.
'
'Have fun :)



Option Explicit
Private mButton As Boolean
Private mbtype As String
Private mX As Integer, mY As Integer
Private mOldSize As Long
Private mLoading As Boolean

Public Sub GoSnow()
Dim i  As Long
Dim NewX As Long
Dim NewY As Long

    Do
        For i = 0 To mFlakeNum - 1
            'Change the draw width to be the size of the flake
            frmMain.DrawWidth = Flakes(i).FlakeSize
            
            'Used PSet instead of SetPixel
            'As with PSet you can change the size of the pixels using DrawWidth
                        
            'Draw the snow flake again in its old position, but with black to remove it
            frmMain.PSet (Flakes(i).oldX, Flakes(i).oldY), vbBlack
            
            'Now draw it again in its new position using the snow color
            frmMain.PSet (Flakes(i).X, Flakes(i).Y), GetSnowColor
        Next i
        
        For i = 0 To mFlakeNum - 1
            'Store the flakes old x & y
            Flakes(i).oldX = Flakes(i).X
            Flakes(i).oldY = Flakes(i).Y
            
            'Get a new xspeed
            NewX = GetXSpeed(i)
                               
            'Keep the X coordinate withing the form
            If NewX < 0 Then NewX = 0
            If NewX >= frmMain.ScaleWidth Then NewX = frmMain.ScaleWidth - 1
            
            'Set the flakes new X coordinate
            Flakes(i).XSpeed = NewX + frmMain.DrawWidth
            'Set the flakes Y cood, adding on the YSpeed and the DrawWidth so it moves
            'in proportion to its size
            NewY = Flakes(i).Y + (Flakes(i).YSpeed + frmMain.DrawWidth)
                                    
            'If the new target coordinates are black then we can
            'set the actual flakes x & y to the new x & y.
            If GetPixel(frmMain.hdc, NewX, NewY) = vbBlack Then
                Flakes(i).Y = NewY
                Flakes(i).X = NewX
            Else
                'This section attempts to let the snow move in a realistic kinda way
                'Basically it looks to see if there is anywhere for it to go left or right
                'This simulates snow sliding
                
                'We look left to see if there is any room for us to move
                If GetPixel(mfrmMainHDC, (Flakes(i).X + 1) + frmMain.DrawWidth, (Flakes(i).Y + 1) + frmMain.DrawWidth) = vbBlack Then
                    Flakes(i).X = Flakes(i).X + 1
                    Flakes(i).Y = Flakes(i).Y + 1
                    
                'We look right to see if there is any room for us to move
                ElseIf GetPixel(mfrmMainHDC, (Flakes(i).X - 1) - frmMain.DrawWidth, (Flakes(i).Y + 1) + frmMain.DrawWidth) = vbBlack Then
                    Flakes(i).X = Flakes(i).X - 1
                    Flakes(i).Y = Flakes(i).Y + 1
                Else
                    'Theres no where to go, so the flake in the array can be initiated
                    'Because the actual flake has been drawn onto the screen, it looks
                    'as if the flake has settled.
                    InitFlake i
                End If
            End If
            
            'Keep the flakes inside of the form
            If Flakes(i).Y >= frmMain.ScaleHeight Then
                InitFlake i
            End If
            
        Next i
        
        'Draw the text. This is done each cycle of the do-loop
        'in case the form has been resized.
        DrawText
        
        DoEvents
        
        'Loop until we set the myStopNow variable, via the Property Let
        'This is done from the settings form, to reset the flakes variables etc.
    Loop Until mtStopSnow = True
End Sub

Private Function GetXSpeed(ByVal Index As Long) As Long
Dim NewX As Long

    'Add some wind to the XSpeed
    'This method allows us to get positive
    'and negative numbers
    NewX = Flakes(Index).X + (mRightWind * Rnd)
    NewX = NewX - (mLeftWind * Rnd)
    GetXSpeed = NewX
End Function

Private Sub Form_DblClick()
    'Show the settings
    frmSettings.Show
End Sub

Private Sub Form_Load()
    mLoading = True
    'Set up a couple of variables
    mSnowText = "Ho Ho Ho!" 'The text in the middle of the form
    mFlakeNum = 500 'How many flakes we gonna have?
    mtFlakeSize = 2 'This is the draw width
    
    'Position the form
    frmMain.Left = (Screen.Width / 2) - (frmMain.Width / 2)
    frmMain.Top = (Screen.Height / 2) - (frmMain.Height / 2)
    frmMain.Show
    mLoading = False
    frmSettings.Show
    'Get everything setup
    Setup
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This allows us to draw red and black lines on the screen.
    'Red with the left and black with the right
    mbtype = Button
    frmMain.PSet (X, Y), ButtonColor
    mX = X
    mY = Y
    mButton = True
End Sub

Private Function ButtonColor() As Long
    If mbtype = 1 Then
        ButtonColor = RGB(255, 0, 0)
    Else
        ButtonColor = RGB(0, 0, 0)
    End If
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mButton = True Then
        mOldSize = frmMain.DrawWidth
        frmMain.DrawWidth = 6
        Me.ForeColor = ButtonColor
        frmMain.Line (mX, mY)-(X, Y)
        frmMain.PSet (X, Y), ButtonColor
        frmMain.DrawWidth = mOldSize
        mX = X
        mY = Y
        Me.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mButton = False
End Sub

Private Sub Form_Resize()
    If mLoading = False Then
        mtStopSnow = True
        frmMain.Cls
        Setup
        mtStopSnow = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Function GetSnowColor() As Long
Dim tmpRnd As Long
Dim tNumOptions As Long
    
    'This function allows you to have flakes with lots of different colours
    'or different shades of grey/white

    'This alters out random number to be within the same
    'number as the number of case statements
    tNumOptions = 2
    
    tmpRnd = Int((tNumOptions * Rnd) + 1)
    Select Case tmpRnd
        Case 1
            GetSnowColor = 16777215
        Case 2
            GetSnowColor = 14737632
        'You can add as many colours as you like
        'but you need to set the tNumOptions variable to be the same
        'as the amount of case statements you have
    End Select
    
    'This is hard coded to always return white.
    'Just remove this line to allow the random
    'colours above to be used instead.
    GetSnowColor = vbWhite
End Function

Public Property Let StopSnow(ByVal InStop As Boolean)
    'Either stop the loop (if true)
    'or let it start again (if false)
    'This is set from the settings form.
    mtStopSnow = InStop
End Property

