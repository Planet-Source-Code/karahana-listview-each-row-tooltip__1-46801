VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4515
   LinkTopic       =   "Form3"
   ScaleHeight     =   1980
   ScaleWidth      =   4515
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   765
      Width           =   1725
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1260
      Width           =   1680
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   270
      Width           =   1725
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub open_dlg(par1, par2, par3, frm As Form)
  Text3(0) = par1
  Text3(1) = par2
  Text3(2) = par3
  Me.Show vbModal, frm
End Sub

