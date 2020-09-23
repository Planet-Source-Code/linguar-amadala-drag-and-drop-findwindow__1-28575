VERSION 5.00
Begin VB.Form frmMouseWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Window"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "frmMouseWindow.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4845
      TabIndex        =   7
      Top             =   585
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   4845
      TabIndex        =   6
      Top             =   135
      Width           =   960
   End
   Begin VB.Frame frFind 
      Caption         =   "Find Window"
      Height          =   2625
      Left            =   75
      TabIndex        =   8
      Top             =   60
      Width           =   4710
      Begin VB.TextBox txtRect 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2175
         Width           =   3780
      End
      Begin VB.TextBox txtClass 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1845
         Width           =   3780
      End
      Begin VB.TextBox txtCaption 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1485
         Width           =   3780
      End
      Begin VB.PictureBox pSel2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   1515
         Picture         =   "frmMouseWindow.frx":000C
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         Top             =   810
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   1035
         Picture         =   "frmMouseWindow.frx":015E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   11
         Top             =   825
         Width           =   480
         Begin VB.PictureBox picFinder 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   135
            Picture         =   "frmMouseWindow.frx":0A28
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   15
            TabIndex        =   12
            Top             =   165
            Width           =   225
         End
      End
      Begin VB.Label lblRect 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rect:"
         Height          =   195
         Left            =   345
         TabIndex        =   4
         Top             =   2175
         Width           =   390
      End
      Begin VB.Label lblClass 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   1845
         Width           =   420
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption:"
         Height          =   195
         Left            =   135
         TabIndex        =   0
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find Cursor:"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   945
         Width           =   840
      End
      Begin VB.Label lblFinderInfo 
         AutoSize        =   -1  'True
         Caption         =   "Use the find cursor in the image below to select the window you want to get more information about."
         Height          =   390
         Left            =   135
         TabIndex        =   9
         Top             =   255
         Width           =   4440
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMouseWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//Author  :  Allen Copeland
'//Purpose :  To select a window using the mouse cursor and return
'//Warning :  The comments on the api are a mere joke, if you think it best, remove the _
              Overly tabulated and organized structure of the declarations.

'/----------------------------------------------------\
'|               Module Level Declarations            |
'\----------------------------------------------------/
Private m_lngWindowCur As Long
    '//The Currently Selected window
Private m_lngWorkingDC As Long
    '//The Device Context Derived from the window to do simple drawing tasks
Public SelectedWindow As Long
    '//Result window

'/----------------------------------------------------\
'|                     Structures                     |
'\----------------------------------------------------/
Private Type FindWindow_Point
    '//Point Structure
    X As Long
        '//X axis value
    Y As Long
        '//Y axis value
End Type

Private Type FindWindow_Rectangle
    '//Rectangle Structure
    Left As Long
        '//Left side of the rect
    Top As Long
        '//Top side of the rect
    Right As Long
        '//Right side of the rect
    Bottom As Long
        '//Bottom side of the rect
End Type

'/----------------------------------------------------\
'|                   API Declarations                 |
'\----------------------------------------------------/
    '/----------------------------------------------------\
    '|                 User32 Declares                    |
    '\----------------------------------------------------/
        '/----------------------------------------------------\
        '|                   Window Classes                   |
        '\----------------------------------------------------/
            Private Declare Function GetClassName _
                Lib "user32" _
                Alias "GetClassNameA" _
                    (ByVal Window As Long, _
                     ByVal Buffer As String, _
                     ByVal BufferLength As Long) _
                As Long
        '/----------------------------------------------------\
        '|                  Rectangles                        |
        '\----------------------------------------------------/
            '/----------------------------------------------------\
            '|                        Size                        |
            '\----------------------------------------------------/
                Private Declare Function InflateRect _
                    Lib "user32" _
                        (Rectangle As FindWindow_Rectangle, _
                         ByVal X As Long, _
                         ByVal Y As Long) _
                    As Long
                Private Declare Function GetWindowRect _
                    Lib "user32" _
                        (ByVal Window As Long, _
                         Rectangle As FindWindow_Rectangle) _
                    As Long
            '/----------------------------------------------------\
            '|                        GDI                         |
            '\----------------------------------------------------/
                Private Declare Function DrawFocusRect _
                    Lib "user32" _
                        (ByVal DeviceContext As Long, _
                         Rectangle As FindWindow_Rectangle) _
                    As Long
        '/----------------------------------------------------\
        '|                      Points                        |
        '\----------------------------------------------------/
            Private Declare Function GetCursorPos _
                Lib "user32" _
                    (Point As FindWindow_Point) _
                As Long
            Private Declare Function WindowFromPoint _
                Lib "user32" _
                    (ByVal X As Long, _
                     ByVal Y As Long) _
                As Long
        '/----------------------------------------------------\
        '|                    Device Contexts                 |
        '\----------------------------------------------------/
            Private Declare Function GetWindowDC _
                Lib "user32" _
                    (ByVal Window As Long) _
                As Long
            Private Declare Function _
                ReleaseDC Lib "user32" _
                    (ByVal Window As Long, _
                     ByVal DeviceContext As Long) _
                As Long
        '/----------------------------------------------------\
        '|                       Windows                      |
        '\----------------------------------------------------/
            Private Declare Function GetParent _
                Lib "user32" _
                    (ByVal Window As Long) _
                As Long
            '/----------------------------------------------------\
            '|                      Validation                    |
            '\----------------------------------------------------/
                Private Declare Function IsWindow _
                    Lib "user32" _
                        (ByVal Window As Long) _
                    As Long
            '/----------------------------------------------------\
            '|                      Strings                       |
            '\----------------------------------------------------/
                Private Declare Function GetWindowText _
                    Lib "user32" _
                    Alias "GetWindowTextA" _
                        (ByVal Window As Long, _
                         ByVal Text As String, _
                         ByVal Length As Long) _
                    As Long
                Private Declare Function GetWindowTextLength _
                    Lib "user32" _
                    Alias "GetWindowTextLengthA" _
                        (ByVal Window As Long) _
                    As Long
            '/----------------------------------------------------\
            '|                       Flags                        |
            '\----------------------------------------------------/
                Private Declare Function GetWindowLong _
                    Lib "user32" _
                    Alias "GetWindowLongA" _
                        (ByVal Window As Long, _
                         ByVal Flag As Long) _
                    As Long
'/----------------------------------------------------\
'|                     API Constants                  |
'\----------------------------------------------------/
    Private Const GWL_HINSTANCE = (-6)
        '//Flag for the Get/SetWindowLong API Call(s).
        '//Allows for you to obtain the application instance of a specific window.

Private Sub cmdCancel_Click()
    SelectedWindow = 0
        '//Indicate failure
    Unload Me
        '//Unload the dialog
End Sub

Private Sub cmdOK_Click()
    Unload Me
        '//Unload the dialog, success or failure is dependant upon the SelectedWindow value
End Sub

Private Sub picFinder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//pFinder::MouseDown
    '//Purpose: To initialize the window selection process
    If ((Button And vbLeftButton) = MouseButtonConstants.vbLeftButton) Then
        picFinder.Visible = False
        Set Screen.MouseIcon = pSel2.Picture
        Screen.MousePointer = MousePointerConstants.vbCustom
        SelectedWindow = 0
        UpdateInfo SelectedWindow
    End If
End Sub

Private Sub picFinder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim m_fwpCursor As FindWindow_Point
        '//The Cursor position
    Dim m_lngCurWindow As Long
        '//The Window under the cursor
    If ((Button And vbLeftButton) = MouseButtonConstants.vbLeftButton) Then
        '//If the left button is down, then...
        If Screen.MousePointer = MousePointerConstants.vbDefault And Screen.MouseIcon Is Nothing Then
            '//Pointer check...
            Set Screen.MouseIcon = pSel2.Picture
                '//Reset the icon
            Screen.MousePointer = MousePointerConstants.vbCustom
                '//Reset the flag
        End If
        GetCursorPos m_fwpCursor
            '//Obtain the cursor's position
        With m_fwpCursor
            '//Select the Point variable's namespace
            m_lngCurWindow = WindowFromPoint(.X, .Y)
                '//Obtain the window under the mouse pointer
        End With
        If GetMainWindow(m_lngCurWindow) = hWnd Then
            '//If the window belongs to the app instance of this window, then...
            If (m_lngWindowCur <> 0) Then
                '//If there is a selected window, then...
                DrawTriRect m_lngWindowCur, m_lngWorkingDC
                    '//Remove the previously drawn rectangles (it's inversion
                    '//So it's removed on a second inversion)
                If Not m_lngWorkingDC = 0 Then
                    '//If the working dc hasn't been released already, then...
                    ReleaseDC m_lngWindowCur, m_lngWorkingDC
                        '//Release it.
                    m_lngWorkingDC = 0
                End If
                m_lngWindowCur = 0
                    '//Release the selected window handle (so this isn't done a second time
                    '//which would cause a flash of the rect from constant reinversion)
                UpdateInfo m_lngWindowCur
                    '//Update the control data
            End If '//(m_lngWindowCur <> 0)
            Exit Sub
                '//Exit, we're done.
        End If '//[Window comparison]
        If (m_lngCurWindow <> m_lngWindowCur) Then
            '//If the current window is something other then the selected window, then...
            If IsWindow(m_lngWindowCur) Then
                '//If the old window is still there, then...
                DrawTriRect m_lngWindowCur, m_lngWorkingDC
                    '//Draw the triple inversion (focus) rectangle
                ReleaseDC m_lngWindowCur, m_lngWorkingDC
                    '//Release the working device context
                m_lngWorkingDC = 0
                    '//Release the Working Device Context handle
            End If
            m_lngWorkingDC = GetWindowDC(m_lngCurWindow)
                '//Obtain a working device context for drawing
            m_lngWindowCur = m_lngCurWindow
                '//Store the window under the cursor, as to prevent reoccurances of this section
                '//of code.
            DrawTriRect m_lngWindowCur, m_lngWorkingDC
                '//Draw the triple inversion (focus) rectangle to indicate it's valid
            UpdateInfo m_lngWindowCur
                '//Update control data
        End If '//(m_lngCurWindow <> m_lngWindowCur)
    End If '//[Left button valid state]
End Sub

Private Sub picFinder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//pFinder::MouseUp
    If ((Button And MouseButtonConstants.vbLeftButton) = MouseButtonConstants.vbLeftButton) Then
        '//If the left button is held down, then...
        DrawTriRect m_lngWindowCur, m_lngWorkingDC
            '//Draw triple inversion (focus) rectangle
        If Not m_lngWorkingDC = 0 Then
            '//If the working Device Context hasn't been released, then...
            ReleaseDC m_lngWindowCur, m_lngWorkingDC
                '//Release it
            m_lngWorkingDC = 0
                '//Release the working Device Context handle
        End If
        SelectedWindow = m_lngWindowCur
            '//Store the selected window to assume success if the user presses 'OK'
        m_lngWindowCur = 0
            '//Release the Selected window's Handle (copy)
        Set Screen.MouseIcon = Nothing
            '//Release the screen's mouseicon
        Screen.MousePointer = MousePointerConstants.vbDefault
            '//Restore the app-wide pointer
        picFinder.Visible = True
            '//Restore the visibility of the picturebox
        UpdateInfo SelectedWindow
            '//Update the control information of the selected window.
    End If '//[Left button valid state]
End Sub

Private Sub ScrRectToWndRect(Rectangle As FindWindow_Rectangle)
    With Rectangle
        '//Select the rectangle argument's namespace
        .Right = .Right - .Left
            '//Adjust the right
        .Bottom = .Bottom - .Top
            '//Adjust the bottom
        .Top = 0
            '//Adjust the top
        .Left = 0
            '//Adjust the left
    End With
End Sub

Private Sub DrawTriRect(Window As Long, DC As Long)
    Dim m_fwrRect As FindWindow_Rectangle
        '//Window Position rectangle
    GetWindowRect Window, m_fwrRect
        '//Get the window's rectangle
    ScrRectToWndRect m_fwrRect
        '//Adjust the rectangle to the actual position of the window (since it's a window based
        '//Device Context, it will draw shifted if we leave the left/top the way they are)
    DrawFocusRect DC, m_fwrRect
        '//Draw the First rectangle, the outter parameter
    InflateRect m_fwrRect, -1, -1
        '//Move it in a pixel.
    DrawFocusRect DC, m_fwrRect
        '//Draw it a bit smaller.
    InflateRect m_fwrRect, -1, -1
        '//Move it in another pixel.
    DrawFocusRect DC, m_fwrRect
        '//Draw the third and final rect.
End Sub

Private Sub UpdateInfo(Window As Long)
    Dim m_fwrRect As FindWindow_Rectangle
        '//Window Rectangle Variable
    If (IsWindow(Window) <> 0) Then
        '//If the window is valid, then...
        GetWindowRect Window, m_fwrRect
            '//Get the window's rectangle
        txtCaption = GetWindowCaption(Window)
            '//Show the window's caption
        txtClass = GetWindowClass(Window)
            '//Show the window's class
        With m_fwrRect
            '//Select the Rectangle Variable's namespace
            txtRect.Text = "({" & .Left & ", " & .Top & "}, {" & .Right & ", " & .Bottom & "})"
                '//Show the window's position
        End With
    Else
        txtCaption.Text = "[N/A]"
            '//Show that it's not available for this window, or that there isn't a window
        txtClass.Text = "[N/A]"
            '//...
        txtRect.Text = vbNullString
            '//Clear the rect text
    End If
End Sub

Private Function GetWindowCaption(Window As Long) As String
    Dim m_strBuffer As String
        '//Buffer variable
    m_strBuffer = Space(GetWindowTextLength(Window) + 1)
        '//Initialize the buffer
    GetWindowText Window, m_strBuffer, GetWindowTextLength(Window) + 1
        '//Fill the buffer
    m_strBuffer = TrimTerm(m_strBuffer)
        '//Get the actual text by removing the null character.
    GetWindowCaption = m_strBuffer
        '//Return
End Function

Private Function TrimTerm(Exp As String) As String
    Dim m_lngPos As Long
        '//Null character position variable
    m_lngPos = InStr(1, Exp, vbNullChar)
        'Get the null character's position
    TrimTerm = VBA.Left(Exp, m_lngPos - 1)
        '//Return a new string without the null character
End Function

Private Function GetWindowClass(Window As Long) As String
    Dim m_strBuffer As String
        '//Buffer variable
    m_strBuffer = Space(254) & vbNullChar
        '//Initialize the buffer
    GetClassName Window, m_strBuffer, Len(m_strBuffer)
        '//Fill the buffer
    m_strBuffer = TrimTerm(m_strBuffer)
        '//Get the actual class name by removing the null character
    GetWindowClass = m_strBuffer
        '//Return the window's class name
End Function

Private Sub txtCaption_Change()
    txtCaption.SelStart = 0
    txtCaption.SelLength = Len(txtCaption.Text)
End Sub

Private Sub txtCaption_GotFocus()
    txtCaption.SelStart = 0
    txtCaption.SelLength = Len(txtCaption.Text)
End Sub

Private Sub txtClass_GotFocus()
    txtClass.SelStart = 0
    txtClass.SelLength = Len(txtClass.Text)
End Sub

Private Sub txtRect_GotFocus()
    txtRect.SelStart = 0
    txtRect.SelLength = Len(txtRect.Text)
End Sub

Public Function GetMainWindow(Window As Long) As Long
    Dim m_lngParent As Long
        '//Parent window
    Dim m_lngLastParent As Long
        '//Last parent window handle
    m_lngParent = Window
        '//Store the window as the current window to begin the falling loop through the _
           generations.
    Do Until m_lngParent = 0
        '//Loop until the parent window is null
        m_lngLastParent = m_lngParent
            '//Store the last window, so when the end is reached, the top-level window will be
            '//stored
        m_lngParent = GetParent(m_lngParent)
            '//Get the window's parent
    Loop '//Until m_lngParent = 0
        '//Complete the loop
    GetMainWindow = m_lngLastParent
        '//Return the handle
End Function
