VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo Form"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Praveen"
      Height          =   405
      Left            =   60
      TabIndex        =   4
      Top             =   3540
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   405
      Left            =   5370
      TabIndex        =   3
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   30
      TabIndex        =   2
      Top             =   3330
      Width           =   6585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Demo 2 -- Translucent form with Picture"
      Height          =   525
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   2490
      Width           =   3345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Demo 1 -- Translucent form without Picture"
      Height          =   525
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   1920
      Width           =   3345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   1635
      Left            =   2970
      TabIndex        =   5
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   60
      Picture         =   "Form3.frx":0000
      Top             =   60
      Width           =   2250
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

Me.Hide
DoEvents

Select Case Index

    Case 0
        Form1.Show
        
    Case 1
        Form2.Show
        
End Select

End Sub

Private Sub Command2_Click()
        Unload Me
End Sub

Private Sub Command3_Click()

        MsgBox "Thanx all for the votes showered on me for the TranslucencyOCX V1" & vbCrLf & vbCrLf & _
        "If the source has helped you in some way, or if it taught you something, I consider the work was not futile" & vbCrLf & _
        "Also, If you think this is a novel idea for implementing Translucency and if you feel this ain't that bad for a beginner, please vote"
End Sub

Private Sub Form_Load()

        Label1.Caption = "Guys... Thanx all for the votes showered on me for the first version of TranslucencyOCX" & vbCrLf & "This time, the control has evolved better with some surprise features" & vbCrLf & "With the form Resizing and moveing and with a new picture property, this is too much you can ask for"

End Sub
