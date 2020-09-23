VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form2 
   Caption         =   "ListView Functions Example"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1320
      Top             =   4740
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4095
      Left            =   420
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7223
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "imagelist2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imagelist2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   9
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":005C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub Form_Load()
    
    Dim i As Integer
    
    With ListView1
    
        ' Set ListView Properties
        
        .View = lvwReport
        .FullRowSelect = True
        
        .ColumnHeaders.Add , , "String"
        .ColumnHeaders.Item(1).Tag = str(ldtString)
        
        .ColumnHeaders.Add , , "Number"
        .ColumnHeaders.Item(2).Tag = str(ldtNumber)
        
        .ColumnHeaders.Add , , "Time"
        .ColumnHeaders.Item(3).Tag = str(ldtTime)
        
        
        ' Populate the ListView with Junk

        For i = 1 To 100
            With .ListItems.Add(, , RandomString)
                .ListSubItems.Add , , RandomNumber
                .ListSubItems.Add , , RandomHour
            End With
        Next
        
    End With
  
 
End Sub


Private Sub Form_Resize()
   ListView1.Left = 0
   ListView1.Width = Me.ScaleWidth
End Sub



Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call SortListView(ListView1, ColumnHeader.Index, Val(ColumnHeader.Tag), ListView1.SortOrder)
End Sub

Private Function RandomString() As String
    Dim CHARS As Integer
    CHARS = Int(Rnd * 15)
    Dim i As Integer, str As String
    For i = 1 To CHARS
        str = str & Chr$(Asc("A") + CInt(Rnd * 25))
    Next
    RandomString = str
End Function

Private Function RandomNumber() As String
    Const RANGE As Integer = 200
    RandomNumber = Format$((Rnd * RANGE) - (RANGE / 2), "0.00")
End Function

Private Function RandomDate() As String
    Const RANGE As Integer = 200
    RandomDate = Format$(DateAdd("d", CInt(Rnd * RANGE) - (RANGE / 2), Date), _
                                                                "DD/MM/YYYY")
End Function

Private Function RandomHour() As String
    RandomHour = Format$(Int(Rnd * 24) & ":" & Int(Rnd * 60) & ":" & Int(Rnd * 60), "hh:mm:ss")
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
Unload Form1
End Sub


Private Sub ListView1_DblClick()
  Timer1.Enabled = False
  Form3.open_dlg ListView1.SelectedItem, ListView1.SelectedItem.SubItems(1), ListView1.SelectedItem.SubItems(2), Me
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    i = 0
    Timer1.Enabled = True
 If IsLoaded("Form1") Then
    Unload Form1
    Me.SetFocus
 End If
End Sub

Private Sub Timer1_Timer()
   i = i + 1
   If i > 3 Then
     If Form1.Visible = False Then
        Form1.Left = Me.Left + ListView1.SelectedItem.Left + ListView1.SelectedItem.Width
        Form1.Top = Me.Top + ListView1.SelectedItem.Top + ListView1.SelectedItem.Height + Me.Height - Me.ScaleHeight
        Form1.Show
        Me.SetFocus
        i = 0
        Timer1.Enabled = False
     End If
   End If
End Sub


