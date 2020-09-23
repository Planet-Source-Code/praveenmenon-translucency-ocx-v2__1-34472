Attribute VB_Name = "MdlSubClass"
Option Explicit

'API Declarations used for subclassing.
Public Declare Sub CopyMemory _
                        Lib "kernel32" Alias "RtlMoveMemory" _
                        (pDest As Any, _
                        pSrc As Any, _
                        ByVal ByteLen As Long)

Public Declare Function SetWindowLong _
                        Lib "user32" Alias "SetWindowLongA" _
                        (ByVal hWnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long

Public Declare Function GetWindowLong _
                        Lib "user32" Alias "GetWindowLongA" _
                        (ByVal hWnd As Long, _
                        ByVal nIndex As Long) As Long

Public Declare Function CallWindowProc _
                        Lib "user32" Alias "CallWindowProcA" _
                        (ByVal lpPrevWndFunc As Long, _
                        ByVal hWnd As Long, _
                        ByVal Msg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

'Constants for GetWindowLong() and SetWindowLong() APIs.
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

'Used to hold a reference to the control to call its procedure.
'NOTE: "Translucency" is the UserControl.Name Property at
'      design-time of the .CTL file.
'      ('As Object' or 'As Control' does not work)
Private ctlTranslucency As Translucency

'Used as a pointer to the UserData section of a window.
Private ptrObject As Long

'The address of this function is used for subclassing.
'Messages will be sent here and then forwarded to the
'UserControl's WindowProc function. The HWND determines
'to which control the message is sent.
Public Function SubWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error Resume Next

      'Get pointer to the control's VTable from the
      'window's UserData section. The VTable is an internal
      'structure that contains pointers to the methods and
      'properties of the control.
      ptrObject = GetWindowLong(hWnd, GWL_USERDATA)

      'Copy the memory that points to the VTable of our original
      'control to the shadow copy of the control you use to
      'call the original control's WindowProc Function.
      'This way, when you call the method of the shadow control,
      'you are actually calling the original controls' method.
      CopyMemory ctlTranslucency, ptrObject, 4

      'Call the WindowProc function in the instance of the UserControl.
      SubWndProc = ctlTranslucency.WindowProc(hWnd, Msg, wParam, lParam)

      'Destroy the Translucency Control Copy
      CopyMemory ctlTranslucency, 0&, 4
      Set ctlTranslucency = Nothing
      
End Function
