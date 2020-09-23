VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00D0F4F1&
   BorderStyle     =   0  'None
   Caption         =   "Caption"
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   1
      X1              =   2475
      X2              =   2475
      Y1              =   0
      Y2              =   780
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   780
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   2460
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   2460
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00D0F4F1&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   855
      TabIndex        =   0
      Top             =   315
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Const HWND_TOPMOST = -1

Private Const HWND_BOTTOM = 1


Private Sub Form_Activate()
    Line1(0).X1 = 0
    Line1(0).X2 = Me.Width
    Line1(0).Y1 = 0
    Line1(0).Y2 = 0
        
    Line1(1).X1 = 0
    Line1(1).X2 = Me.Width
    Line1(1).Y1 = Me.Height - (Line1(1).BorderWidth * 8)
    Line1(1).Y2 = Me.Height - (Line1(1).BorderWidth * 8)

    Line2(0).X1 = 0
    Line2(0).X2 = 0
    Line2(0).Y1 = 0
    Line2(0).Y2 = Me.Height
    
    Line2(1).X1 = Me.Width - (Line1(1).BorderWidth * 8)
    Line2(1).X2 = Me.Width - (Line1(1).BorderWidth * 8)
    Line2(1).Y1 = 0
    Line2(1).Y2 = Me.Height - (Line1(1).BorderWidth * 8)

End Sub

Private Sub Form_Click()
Unload Me

End Sub


Private Sub Form_Load()
    Label1.Caption = "String: " & form2.ListView1.SelectedItem & vbCrLf & "Number: " & form2.ListView1.SelectedItem.SubItems(1) & vbCrLf & "Time: " & form2.ListView1.SelectedItem.SubItems(2)
    

    Label1.Left = 100
    Label1.Top = 100
    Me.Height = Label1.Height + 200
    Me.Width = Label1.Width + 200
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
