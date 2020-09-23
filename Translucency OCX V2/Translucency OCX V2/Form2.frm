VERSION 5.00
Object = "*\A..\TRANSL~1\TranslucencyOCX.vbp"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2235
   LinkTopic       =   "Form2"
   ScaleHeight     =   3645
   ScaleWidth      =   2235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   405
      Left            =   1290
      TabIndex        =   0
      Top             =   3180
      Width           =   885
   End
   Begin TranslucencyOCX.Translucency Translucency1 
      Left            =   570
      Top             =   780
      _ExtentX        =   1058
      _ExtentY        =   1058
      Subclassed      =   -1  'True
      BlendColor      =   16777215
      BlendPicture    =   "Form2.frx":0000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
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
