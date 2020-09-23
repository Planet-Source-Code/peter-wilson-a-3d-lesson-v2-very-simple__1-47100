VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "MIDAR's Simple 3D Lesson"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   60
      Top             =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API used for reading the keyboard.
' ==================================
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


' This is a 3 dimensional vector, because it holds 3 values (X, Y & Z).
' Now you understand multi-dimensional vectors. See?, Vectors are not hard!
' =========================================================================
Private Type Vector3
    X As Single
    Y As Single
    Z As Single
End Type


' m_Dots & m_Temp will hold an Array of Vectors (as defined above).
' ======================================================
Private m_Dots() As Vector3 ' << We define our Dots only once, and store them here.
Private m_Temp() As Vector3 ' << This is a temporary working area, used to apply the camera.


' This is the position of our Camera.
' Don't get too excited, this camera is *very* primitive indeed (but we still need one)
' =====================================================================================
Private m_Camera As Vector3
Private m_CameraZoom As Single  ' Zoom is very easy to implement, and I have included it for people wanting to "zoom" in for the "Head-Shot!" action.

Private Sub CreateGrid()

    ' =====================================================
    ' Create a nice big test grid (and some ground clutter)
    ' =====================================================
    
    Screen.MousePointer = vbHourglass
    
    Dim lngIndex As Long
    
    Dim intX As Integer
    Dim intY As Integer
    Dim intZ As Integer
    
    lngIndex = -1 ' Reset to -1, because we'll soon be increasing this value to 0 (the start of our array)
    
    ' ===============================================================
    ' Create some random ground clutter (ie. grass blades, whatever?)
    ' (If you are feeling adventurous, you might like to introduce
    ' colour into this application to make the grass green.
    ' ===============================================================
    For intX = 0 To 250                             '   << Try increase the number of ground clutter dots
        lngIndex = lngIndex + 1
        ReDim Preserve m_Dots(lngIndex)
        m_Dots(lngIndex).X = (Rnd * 100) - 50       '   << ie. Random number between -50 and +50
        m_Dots(lngIndex).Y = 0                      '   << Because this is the ground, the elevation is zero.
        m_Dots(lngIndex).Z = (Rnd * 100) - 50
    Next intX
    
    
    ' ====================================================================
    ' Create 3 large lines out of dots, representing the 3 axes (x, y & z)
    ' ====================================================================
    For intX = -100 To 100 Step 10                  '   << Positive X points to the Right
        For intZ = -100 To 100 Step 10              '   << Positive Z points *into* the monitor - away from you.
            For intY = -100 To 100 Step 10          '   << Positive Y goes Up
                
                If (intX = 0 And intY = 0) Or (intX = 0 And intZ = 0) Or (intY = 0 And intZ = 0) Then
                
                    lngIndex = lngIndex + 1
                    ReDim Preserve m_Dots(lngIndex)
                    m_Dots(lngIndex).X = intX
                    m_Dots(lngIndex).Y = intY
                    m_Dots(lngIndex).Z = intZ
                    
                End If
            Next intY
        Next intZ
    Next intX
    
    
    ' ===================
    ' Create a Test Grid.
    ' ===================
    For intX = -100 To 100 Step 10                  '   << Positive X points to the Right
        For intZ = -100 To 100 Step 10              '   << Positive Z points *into* the monitor - away from you.
            For intY = -100 To 100 Step 10          '   << Positive Y goes Up

                If (Abs(intX) = 100 Or Abs(intZ) = 100 Or intX = 0 Or intZ = 0) And (intY = 0 Or Abs(intY) = 100) Then

                    ' Create the basement (below ground level), floor and roof (of the test grid)
                    lngIndex = lngIndex + 1
                    ReDim Preserve m_Dots(lngIndex)
                    m_Dots(lngIndex).X = intX / 100
                    m_Dots(lngIndex).Y = intY / 100
                    m_Dots(lngIndex).Z = intZ / 100

                ElseIf Abs(intX) = 100 And Abs(intZ) = 100 Then

                    ' Put some corners on it (ie. 4 support beams)
                    lngIndex = lngIndex + 1
                    ReDim Preserve m_Dots(lngIndex)
                    m_Dots(lngIndex).X = intX / 100
                    m_Dots(lngIndex).Y = intY / 100
                    m_Dots(lngIndex).Z = intZ / 100

                End If
            Next intY
        Next intZ
    Next intX
           
           
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub CreateGridWaves()
    
    Screen.MousePointer = vbHourglass
    
    Dim lngIndex As Long
    
    Dim intX As Integer
    Dim intY As Integer
    Dim intZ As Integer
    
    lngIndex = -1 ' Reset to -1, because we'll soon be increasing this value to 0 (the start of our array)
    For intX = -100 To 100 Step 5       '   << Try adjusting the Step value to 1, 2, 5, 10 or 25.
        For intZ = -100 To 100 Step 5   '   << Try adjusting the Step value to 1, 2, 5, 10 or 25.
            
            lngIndex = lngIndex + 1
            ReDim Preserve m_Dots(lngIndex)
            m_Dots(lngIndex).X = intX / 100                                 '   << Positive X points to the Right
            m_Dots(lngIndex).Y = Cos(Sqr(intX * intX + intZ * intZ) / 30)   '   << Positive Y goes Up
            m_Dots(lngIndex).Z = intZ / 100                                 '   << Positive Z points *into* the monitor - away from you.
            
        Next intZ
    Next intX
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub DoDisplay3DPoints()

    ' ===============================
    ' Draws the Dots onto the screen.
    ' ===============================
    
    On Error GoTo errTrap
    
    Dim lngIndex As Long
    Dim PixelX As Single
    Dim PixelY As Single
    
    
    ' Clear the form and set the drawing style and width, etc.
    ' ========================================================
    Me.Cls                              '   << Clear the screen.
    Me.DrawWidth = 1                    '   << Set the Width of the Pen. Any value higher than 1 will slow down animation.
    Me.ForeColor = RGB(255, 255, 255)   '   << Bright White
    
    
    ' Set the size of the temporary array, to the same size as the Dots array.
    ReDim m_Temp(UBound(m_Dots))
    
    
    ' Loop through from the "Lower Boundry" of the Array, to the "Upper Boundry" of the Array.
    For lngIndex = LBound(m_Dots) To UBound(m_Dots)
    
        ' ===========================================================================================
        ' Apply the Camera to the Dots (ie. move the Dots away from us, to the left, the right, etc.)
        ' Remember, this is a *very* primitive camera.
        ' ===========================================================================================
        m_Temp(lngIndex).X = m_Dots(lngIndex).X - m_Camera.X    ' << Note: This is slightly different. Version 1 was used positive instead of negative, Sorry.
        m_Temp(lngIndex).Y = m_Dots(lngIndex).Y - m_Camera.Y    ' << etc.
        m_Temp(lngIndex).Z = m_Dots(lngIndex).Z - m_Camera.Z    ' << etc.
        
        
        ' Only draw dots in front of the camera (and not behind us).
        ' Good place to insert "fading" brightness depending on how far away the dots are.
        If m_Temp(lngIndex).Z > 0 Then
            
            
            ' ********************************************************************************
            ' Transform our 3D vector, down to a 2D vector. This little part here, is the most
            ' important thing about displaying 3D computer graphics... and it's so simple!!
            ' ********************************************************************************
            PixelX = m_Temp(lngIndex).X / m_Temp(lngIndex).Z
            PixelY = -m_Temp(lngIndex).Y / m_Temp(lngIndex).Z   ' << This is negative Y pixel, because in this application I want positive Y to go up, however Microsoft's Operating System has positive Y going down.
                        
                        
            ' Plot the point
            Me.PSet (PixelX, PixelY)
            
        End If
        
     Next lngIndex
    
    
    Exit Sub
errTrap:
    ' Occasionally we get overflow errors plotting a pixel with small values (a MS bug?),
    ' just ignore them and exit the sub-routine. They are not important enough to worry
    ' about in an application this simple.
    
End Sub

Private Sub ShowParameters()

    ' ==========================================================================================
    ' This routine slows down the program, because printing text is very slow.
    ' Speed has been sacrificed for instructional clarity for beginners to 3D Computer Graphics.
    ' Remember that by-and-large I am programming things the slow way, in an effort to be clear.
    ' You can always speed up my code by making your own clever adjustments.
    ' ==========================================================================================
    
    Dim sngX As Single
    
    ' Set our start printing position
    ' Remember, The origin of our screen has been moved into the center of the window, but we want text top-left.
    Me.ForeColor = RGB(255, 255, 192)
    sngX = Me.ScaleLeft
    Me.CurrentY = Me.ScaleTop
    
    
    ' Show product name.
    Me.CurrentX = sngX
    Me.Print App.ProductName & " - " & App.LegalCopyright
    
    
    ' Show helpful reminders.
    Me.CurrentX = sngX
    Me.Print "Keys:  ESC, Left/Right/Up/Down, Shift-Up/Down, Page-Up/Down, Spacebar."
    
    
    ' Show helpful reminders.
    Me.CurrentX = sngX
    Me.Print "Mouse:  Move mouse over dots to display original defined coordinates." & vbNewLine
    
    
    ' Show current Camera position.
    Me.CurrentX = sngX
    Me.Print "Camera:  x: " & Format(m_Camera.X, "Fixed") & "   y: " & Format(m_Camera.Y, "Fixed") & "   z: " & Format(m_Camera.Z, "Fixed")
    
    
    ' Show current Camera Zoom value.
    Me.CurrentX = sngX
    Me.Print "Camera Zoom: " & Format((1 / m_CameraZoom), "Fixed")
    
End Sub

Private Sub UpdateCameraParameters()

    ' ===================================================================
    ' This routine looks at the keyboard, and adjusts the camera position
    ' and Zoom values depending on which keys are held down.
    ' ===================================================================
    
    Dim lngKeyState As Long
    Dim sngCameraStep As Single
    
    sngCameraStep = 1    ' << Adjust this to move the camera faster or slower (any value not zero)
    
    ' =======================
    ' Move Camera Left/Right.
    ' =======================
    lngKeyState = GetKeyState(vbKeyLeft)
    If (lngKeyState And &H8000) Then m_Camera.X = m_Camera.X - sngCameraStep
    lngKeyState = GetKeyState(vbKeyRight)
    If (lngKeyState And &H8000) Then m_Camera.X = m_Camera.X + sngCameraStep

    lngKeyState = GetKeyState(vbKeyShift)
    If (lngKeyState And &H8000) Then
        
        ' ======================================================================
        ' Shift Key is down, the user must want to move closer, or further away.
        ' ======================================================================
        lngKeyState = GetKeyState(vbKeyUp)
        If (lngKeyState And &H8000) Then m_Camera.Z = m_Camera.Z + sngCameraStep
        lngKeyState = GetKeyState(vbKeyDown)
        If (lngKeyState And &H8000) Then m_Camera.Z = m_Camera.Z - sngCameraStep
    
    Else
    
        ' =============================================
        ' Shift Key is *not* down. Move camera up/down.
        ' =============================================
        lngKeyState = GetKeyState(vbKeyUp)
        If (lngKeyState And &H8000) Then m_Camera.Y = m_Camera.Y + sngCameraStep
        lngKeyState = GetKeyState(vbKeyDown)
        If (lngKeyState And &H8000) Then m_Camera.Y = m_Camera.Y - sngCameraStep
    
    End If
    
    
    ' ==============================================================================================
    ' Modify the following:
    '   * Field Of View (FOV)
    '   * Camera's Zoom value
    '
    '   Note: These two values are pretty much the same thing, it depends on how you think about it.
    '         You could also think of this as the "Perspective Distortion" as well.
    '
    ' All of this is achieved simply by adjusting the height/width of the window.
    ' It might sound simple, but in reality this is pretty much what the complex 3D engines do.
    ' ==============================================================================================
    lngKeyState = GetKeyState(vbKeyPageUp)
    If (lngKeyState And &H8000) Then
        If m_CameraZoom > 0.2 Then
            m_CameraZoom = m_CameraZoom - 0.1   '   Adjust Zoom Value
            Call Form_Resize                    '   Redefine the Height/Width of our drawing window.
        End If
    End If
    lngKeyState = GetKeyState(vbKeyPageDown)
    If (lngKeyState And &H8000) Then
        m_CameraZoom = m_CameraZoom + 0.1       '   Adjust Zoom Value
        Call Form_Resize                        '   Redefine the Height/Width of our drawing window.
    End If
    
    
    ' ====================================
    ' Reset Camera to a starting position.
    ' ====================================
    lngKeyState = GetKeyState(vbKeySpace)
    If (lngKeyState And &H8000) Then
        ' Reset Camera
        m_Camera.X = 0
        m_Camera.Y = 3
        m_Camera.Z = -15

        
        ' Reset Zoom/FOV
        m_CameraZoom = 1
        Call Form_Resize
    End If
    
    
    ' ========================================
    ' Check for ESC / Quit / Exit Application.
    ' ========================================
    lngKeyState = GetKeyState(vbKeyEscape)
    If (lngKeyState And &H8000) Then
        ' Quit Application
        Me.Timer1.Enabled = False
        Unload Me
    End If
    
    
End Sub

Private Sub Form_Load()

    ' Set some basic form properties.
    ' ===============================
    Me.AutoRedraw = True
    Me.BackColor = RGB(0, 0, 0)
    
    ' Create our test data (two different types)
    ' ==========================================
    Call CreateGrid
'    Call CreateGridWaves
    
    
    ' Position the Camera away from the center of the Dots.
    ' =====================================================
    '   * Positive X points to the Right
    '   * Positive Z points *into* the monitor - away from you.
    '   * Positive Y goes Up
    m_Camera.X = 0
    m_Camera.Y = 3
    m_Camera.Z = -15
    
    
    ' Reset the Camera's Zoom setting.
    ' ================================
    m_CameraZoom = 1
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' *** Important Notice *** Important Notice *** Important Notice ***
    '
    ' This routine slows down the program when the mouse is moved. This is because it loops through
    ' all of the dots (again). This routine is almost identical to the
    '
    ' *** Important Notice *** Important Notice *** Important Notice ***
    
    On Error GoTo errTrap
    
    Dim lngIndex As Long
    Dim PixelX As Single
    Dim PixelY As Single
    Dim sngAutoTolerance As Single
    
    Me.Font = "Arial"
    Me.FontSize = 7
    Me.DrawWidth = 4
    Me.ForeColor = RGB(255, 255, 0)
    
    sngAutoTolerance = m_CameraZoom / 25
    
    For lngIndex = LBound(m_Temp) To UBound(m_Temp)
        
        ' Ignore dots behind the camera.
        If m_Temp(lngIndex).Z > 0 Then
        
            PixelX = m_Temp(lngIndex).X / m_Temp(lngIndex).Z
            PixelY = -m_Temp(lngIndex).Y / m_Temp(lngIndex).Z
            
            ' Is the mouse close to the X coordinate?
            If Abs(PixelX - X) < sngAutoTolerance Then

                ' Is the mouse close to the Y coordinate?
                If Abs(PixelY - Y) < sngAutoTolerance Then
                    
                    ' Plot the pixel
                    Me.PSet (PixelX, PixelY)
                    Me.Print "x:" & Format(m_Dots(lngIndex).X, "Fixed") & " y:" & Format(m_Dots(lngIndex).Y, "Fixed") & " z:" & Format(m_Dots(lngIndex).Z, "Fixed")
                    
                End If
            End If
        End If
    Next lngIndex
    
    Exit Sub
errTrap:
    
End Sub

Private Sub Form_Resize()

    ' Reset the width and height of our form, and also move the origin (0,0) into
    ' the centre of the form. This makes our life much easier. The form's Top-Left
    ' corner will be (-1,-1) whilst the lower right corner will be (1,1)
    
    Me.ScaleWidth = m_CameraZoom
    Me.ScaleLeft = -m_CameraZoom / 2
    
    Me.ScaleHeight = m_CameraZoom
    Me.ScaleTop = -m_CameraZoom / 2
    
End Sub

Private Sub Timer1_Timer()
        
    Call DoDisplay3DPoints
    
    Call ShowParameters
        
    Call UpdateCameraParameters
    
End Sub

