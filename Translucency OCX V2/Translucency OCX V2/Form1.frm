VERSION 5.00
Object = "*\A..\TRANSL~1\TranslucencyOCX.vbp"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   705
      Left            =   6600
      TabIndex        =   0
      Top             =   2430
      Width           =   1065
   End
   Begin TranslucencyOCX.Translucency Translucency1 
      Left            =   2520
      Top             =   1110
      _ExtentX        =   1058
      _ExtentY        =   1058
      Subclassed      =   -1  'True
      BlendColor      =   16744576
      BlendPicture    =   "Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================
'The code included with this pack is written by Praveen Menon.
'The rights for the translucency OCX goes entirely to Praveen Menon.
'==========================================================================================
'If someone want to use the control in their project and needs to distribute
'the complied version of the OCX along with the application, please do it with pleasure
'But if you plan to reproduce or distribute the source code as such or any part of it,
'modified or original, please contact the author. He can be reached at
'praveenmenon_in@yahoo.com
'==========================================================================================
'lastly, Please refrain from boasting that the code is yours'
'==========================================================================================

Option Explicit


Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Form_Load()

  '=========================================================
  'Tell the control to start drawing the Translucency
  '=========================================================
  'if you need the form to be resized and moved, you
  'need to set the Subclassed property to true.
  'Otherwise it remains unSubclassed and doesn't redraw
  'everytime you move it or resize it.
  '=========================================================

    Translucency1.drawTranslucency

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Form3.Show
End Sub
