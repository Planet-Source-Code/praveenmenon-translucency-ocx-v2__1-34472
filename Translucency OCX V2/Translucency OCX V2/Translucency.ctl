VERSION 5.00
Begin VB.UserControl Translucency 
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "Translucency.ctx":0000
   ScaleHeight     =   6885
   ScaleWidth      =   6855
   Begin VB.PictureBox picBoard 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   4020
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Shape Shape1 
      Height          =   285
      Left            =   2460
      Top             =   2220
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "Translucency.ctx":0030
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "Translucency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'TranslucencyOCX V2
'==================================!=======!==================================================
'This is the second version of the much popular TranslucencyOCX. This time the TranslucencyOCX
'supports resizing and moving.
'==================================!=======!==================================================
'*********************************************************************
'************** Project         :  Translucency OCX     **************
'************** Version         :  2                    **************
'************** Author          :  Praveen Menon        **************
'************** Last Updated    :  30th April 2002      **************
'************** Copyright © 2002 Praveen Menon          **************
'*********************************************************************
'==================================!=======!==================================================
'Distribute freely the compled version of the project or the OCX
'But if you plan to distribute the source code as such, please contact me
'I can be reached at praveenmenon_in@yahoo.com
'==================================!=======!==================================================
'
Option Explicit

'==========================================================================================
'The API Declares for drawing the screenshot on the Form's Back
Private Declare Function CreateDC Lib "Gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function BitBlt Lib "Gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "Gdi32.dll" (ByVal hdc As Long) As Long
'==========================================================================================
'This API Blends the Forms' existing backcolor to the new picture. If you need to blend
'another picture to the present screenshot, just set the form's Picture property to that picture.
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal dreamAKA As Long) As Long
'==========================================================================================
'This API gives you measures of different parts of the window, so that you can apply the
'values while drawing. This is important, because, the part of the screen from where the
'screenshot is taken, is determined by using these values
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'Constants for retreiving info about the window
Private Const SM_CYCAPTION = 4
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
'==========================================================================================

'THe Messages to be trapped
Private Const WM_MYMSG = &H232
Private Const WM_PRAVEEN = &H5
'==========================================================================================

'mWndProcOrg holds the original address of the
'Window Procedure for this window. This is used to
'route messages to the original procedure after you
'process them.
Private mWndProcOrg As Long
'==========================================================================================

'Handle (hWnd) of the subclassed window.
Private mHWndSubClassed As Long
'==========================================================================================

Private bFirstTime As Boolean
'==========================================================================================

'Default Property Values:
Const m_def_BlendColor = 0
Private Const m_def_Subclassed = 0

'Property Variables:
Dim m_BlendColor As OLE_COLOR
Dim m_BlendPicture As Picture
Private m_Subclassed As Boolean
Private I_DiD_that As Boolean

Private Sub SubClass(f_hwnd As Long)

  '-------------------------------------------------------------
  'Initiates the subclassing of this UserControl's window (hwnd).
  'Records the original WinProc of the window in mWndProcOrg.
  'Places a pointer to the object in the window's UserData area.
  '-------------------------------------------------------------

  'Exit if the window is already subclassed.

    If mWndProcOrg Then Exit Sub

    'Redirect the window's messages from this control's default
    'Window Procedure to the SubWndProc function in your .BAS
    'module and record the address of the previous Window
    'Procedure for this window in mWndProcOrg.
    mWndProcOrg = SetWindowLong(f_hwnd, GWL_WNDPROC, AddressOf SubWndProc)

    'Record your window handle in case SetWindowLong gave you a
    'new one. You will need this handle so that you can unsubclass.
    mHWndSubClassed = f_hwnd

    'Store a pointer to this object in the UserData section of
    'this window that will be used later to get the pointer to
    'the control based on the handle (hwnd) of the window getting
    'the message.
    Call SetWindowLong(f_hwnd, GWL_USERDATA, ObjPtr(Me))

End Sub

Private Sub UnSubClass()

  '-----------------------------------------------------------
  'Unsubclasses this UserControl's window (hwnd), setting the
  'address of the Windows Procedure back to the address it was
  'at before it was subclassed.
  '-----------------------------------------------------------

  'Ensures that you don't try to unsubclass the window when
  'it is not subclassed.

    If mWndProcOrg = 0 Then Exit Sub

    'Reset the window's function back to the original address.
    SetWindowLong mHWndSubClassed, GWL_WNDPROC, mWndProcOrg
    '0 Indicates that you are no longer subclassed.
    mWndProcOrg = 0

End Sub

Friend Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
       ByVal lParam As Long) As Long

  '--------------------------------------------------------------
  'Process the window's messages that are sent to the UserControl.
  'The WindowProc function is declared as a "Friend" function so
  'that the .BAS module can call the function but the function
  'cannot be seen from outside the UserControl project.
  '--------------------------------------------------------------

    Select Case uMsg

      Case WM_MYMSG

        If bFirstTime Then

            bFirstTime = False
            Exit Function '>---> Bottom

        End If

        I_DiD_that = True
        With UserControl.Parent

            .Hide
            DoEvents
            Init .Left, .Top, .Width, .Height

        End With 'USERCONTROL.PARENT

      Case WM_PRAVEEN
        If Not I_DiD_that Then
        
            With UserControl.Parent

                .Hide
                DoEvents
                Init .Left, .Top, .Width, .Height

            End With 'USERCONTROL.PARENT

        End If

      Case Else
    I_DiD_that = False
        WindowProc = CallWindowProc(mWndProcOrg, hWnd, uMsg, wParam, ByVal lParam)
    End Select

End Function

Private Sub Init(iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer)

    On Error Resume Next

      With UserControl.Parent
          .Visible = False
          .ScaleMode = 3
          .AutoRedraw = True
          .Cls
          .BackColor = m_BlendColor
          .Picture = m_BlendPicture
          BlendIT UserControl.Parent, picBoard
          .Visible = True

      End With 'USERCONTROL.PARENT

      bFirstTime = False

End Sub

Private Sub BlendIT(Frm As Form, Pic As PictureBox)

Dim titleBarheight As Integer
Dim xDeviation As Integer
Dim yDeviation As Integer
Dim windowFrameHeight As Integer
Dim windowframewidth As Integer
Dim BlendVal As Long
Dim hDCscr As Long

    Pic.Move 0, 0, Frm.Width, Frm.Height
    Pic.Cls

    hDCscr = CreateDC("DISPLAY", "", "", 0)

    If Frm.BorderStyle <> 0 Then

        titleBarheight = GetSystemMetrics(SM_CYCAPTION)
        windowFrameHeight = GetSystemMetrics(SM_CYFRAME)
        windowframewidth = GetSystemMetrics(SM_CXFRAME)
        yDeviation = titleBarheight + windowFrameHeight
        xDeviation = windowframewidth

      Else 'NOT FRM.BORDERSTYLE...

        yDeviation = 0
        xDeviation = 0

    End If

    BitBlt Pic.hdc, 0, 0, Frm.ScaleWidth, Frm.ScaleHeight, hDCscr, Frm.Left / Screen.TwipsPerPixelX + xDeviation, Frm.Top / Screen.TwipsPerPixelY + yDeviation, vbSrcCopy

    BlendVal = 11796480
    'Blend the picture on the picturebox to the form
    AlphaBlend Frm.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, Pic.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, BlendVal

    Pic.Cls

    DeleteDC hDCscr

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    If Ambient.UserMode Then
        bFirstTime = True
    End If
    
    m_BlendColor = PropBag.ReadProperty("BlendColor", m_def_BlendColor)
    m_Subclassed = PropBag.ReadProperty("Subclassed", m_def_Subclassed)
    Set m_BlendPicture = PropBag.ReadProperty("BlendPicture", Nothing)

    UserControl.BackColor = m_BlendColor
    
End Sub

Private Sub UserControl_Resize()

    Image1.Move 60, 60
    Width = Image1.Width + 120
    Height = Image1.Height + 120
    Shape1.Move 0, 0, Width, Height
    
End Sub

Private Sub UserControl_Terminate()

    On Error Resume Next
      UnSubClass

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function drawTranslucency() As Variant

    Init UserControl.Parent.Left, UserControl.Parent.Top, UserControl.Parent.ScaleWidth, UserControl.Parent.ScaleHeight
    If m_Subclassed Then
        SubClass UserControl.Parent.hWnd
    End If

End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Subclassed() As Boolean
Attribute Subclassed.VB_ProcData.VB_Invoke_Property = "General"

    Subclassed = m_Subclassed

End Property

Public Property Let Subclassed(ByVal New_Subclassed As Boolean)

    m_Subclassed = New_Subclassed
    PropertyChanged "Subclassed"

End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_Subclassed = m_def_Subclassed

    m_BlendColor = m_def_BlendColor
    Set m_BlendPicture = LoadPicture("")
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Subclassed", m_Subclassed, m_def_Subclassed)
    Call PropBag.WriteProperty("BlendColor", m_BlendColor, m_def_BlendColor)
    Call PropBag.WriteProperty("BlendPicture", m_BlendPicture, Nothing)
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BlendColor() As OLE_COLOR
    BlendColor = m_BlendColor
End Property

Public Property Let BlendColor(ByVal New_BlendColor As OLE_COLOR)
    m_BlendColor = New_BlendColor
    PropertyChanged "BlendColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get BlendPicture() As Picture
    Set BlendPicture = m_BlendPicture
End Property

Public Property Set BlendPicture(ByVal New_BlendPicture As Picture)
    Set m_BlendPicture = New_BlendPicture
    PropertyChanged "BlendPicture"
End Property

