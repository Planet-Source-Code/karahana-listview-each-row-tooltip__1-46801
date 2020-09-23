Attribute VB_Name = "ListViewFuncs"
Option Explicit

Public Const HDS_HOTTRACK = &H4

Public Const HDI_BITMAP = &H10
Public Const HDI_IMAGE = &H20
Public Const HDI_ORDER = &H80
Public Const HDI_FORMAT = &H4
Public Const HDI_TEXT = &H2
Public Const HDI_WIDTH = &H1
Public Const HDI_HEIGHT = HDI_WIDTH

Public Const HDF_LEFT = 0
Public Const HDF_RIGHT = 1
Public Const HDF_IMAGE = &H800
Public Const HDF_BITMAP_ON_RIGHT = &H1000
Public Const HDF_BITMAP = &H2000
Public Const HDF_STRING = &H4000

Public Const HDM_FIRST = &H1200
Public Const HDM_SETITEM = (HDM_FIRST + 4)

Public Const LVM_FIRST = &H1000
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const LVM_GETBKCOLOR = (LVM_FIRST + 0)
Public Const LVM_SETBKCOLOR = (LVM_FIRST + 1)
Public Const LVM_GETIMAGELIST = (LVM_FIRST + 2)
Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Public Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Public Const LVM_GETITEMA = (LVM_FIRST + 5)
Public Const LVM_GETITEM = LVM_GETITEMA
Public Const LVM_SETITEMA = (LVM_FIRST + 6)
Public Const LVM_SETITEM = LVM_SETITEMA
Public Const LVM_INSERTITEMA = (LVM_FIRST + 7)
Public Const LVM_INSERTITEM = LVM_INSERTITEMA
Public Const LVM_DELETEITEM = (LVM_FIRST + 8)
Public Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
Public Const LVM_GETCALLBACKMASK = (LVM_FIRST + 10)
Public Const LVM_SETCALLBACKMASK = (LVM_FIRST + 11)
Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Public Const LVM_FINDITEMA = (LVM_FIRST + 13)
Public Const LVM_FINDITEM = LVM_FINDITEMA
Public Const LVM_GETITEMRECT = (LVM_FIRST + 14)
Public Const LVM_SETITEMPOSITION = (LVM_FIRST + 15)
Public Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
Public Const LVM_GETSTRINGWIDTHA = (LVM_FIRST + 17)
Public Const LVM_GETSTRINGWIDTH = LVM_GETSTRINGWIDTHA
Public Const LVM_HITTEST = (LVM_FIRST + 18)
Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Public Const LVM_SCROLL = (LVM_FIRST + 20)
Public Const LVM_REDRAWITEMS = (LVM_FIRST + 21)
Public Const LVM_ARRANGE = (LVM_FIRST + 22)
Public Const LVM_EDITLABELA = (LVM_FIRST + 23)
Public Const LVM_EDITLABEL = LVM_EDITLABELA
Public Const LVM_GETEDITCONTROL = (LVM_FIRST + 24)
Public Const LVM_GETCOLUMNA = (LVM_FIRST + 25)
Public Const LVM_GETCOLUMN = LVM_GETCOLUMNA
Public Const LVM_SETCOLUMNA = (LVM_FIRST + 26)
Public Const LVM_SETCOLUMN = LVM_SETCOLUMNA
Public Const LVM_INSERTCOLUMNA = (LVM_FIRST + 27)
Public Const LVM_INSERTCOLUMN = LVM_INSERTCOLUMNA
Public Const LVM_DELETECOLUMN = (LVM_FIRST + 28)
Public Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVM_CREATEDRAGIMAGE = (LVM_FIRST + 33)
Public Const LVM_GETVIEWRECT = (LVM_FIRST + 34)
Public Const LVM_GETTEXTCOLOR = (LVM_FIRST + 35)
Public Const LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
Public Const LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Public Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
Public Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
Public Const LVM_GETORIGIN = (LVM_FIRST + 41)
Public Const LVM_UPDATE = (LVM_FIRST + 42)
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXTA = (LVM_FIRST + 45)
Public Const LVM_GETITEMTEXT = LVM_GETITEMTEXTA
Public Const LVM_SETITEMTEXTA = (LVM_FIRST + 46)
Public Const LVM_SETITEMTEXT = LVM_SETITEMTEXTA
Public Const LVM_SETITEMCOUNT = (LVM_FIRST + 47)
Public Const LVM_SORTITEMS = (LVM_FIRST + 48)
Public Const LVM_SETITEMPOSITION32 = (LVM_FIRST + 49)
Public Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Public Const LVM_GETITEMSPACING = (LVM_FIRST + 51)
Public Const LVM_GETISEARCHSTRINGA = (LVM_FIRST + 52)
Public Const LVM_GETISEARCHSTRING = LVM_GETISEARCHSTRINGA
Public Const LVM_SETICONSPACING = (LVM_FIRST + 53)
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Public Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Public Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
Public Const LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)
Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Public Const LVM_SETHOTITEM = (LVM_FIRST + 60)
Public Const LVM_GETHOTITEM = (LVM_FIRST + 61)
Public Const LVM_SETHOTCURSOR = (LVM_FIRST + 62)
Public Const LVM_GETHOTCURSOR = (LVM_FIRST + 63)
Public Const LVM_APPROXIMATEVIEWRECT = (LVM_FIRST + 64)

Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVSCW_AUTOSIZE As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Public Enum ListDataType
    ldtString = 0
    ldtTime = 0
    ldtNumber = 1
    ldtDate = 2
End Enum

Type POINTAPI
        x As Long
        y As Long
End Type

Public Type LV_FINDINFO
      flags As Long
      psz As String
      lParam As Long
      pt As POINTAPI
      vkDirection As Long
End Type

Public Type LV_ITEM
      mask As Long
      iItem As Long
      iSubItem As Long
      State As Long
      stateMask As Long
      pszText As Long
      cchTextMax As Long
      iImage As Long
      lParam As Long
      iIndent As Long
End Type

Public Type RECT
      Left As Long
      Top As Long
      Right As Long
      Bottom As Long
End Type

Public Type HD_ITEM
      mask        As Long
      cxy         As Long
      pszText     As String
      hbm         As Long
      cchTextMax  As Long
      fmt         As Long
      lParam      As Long
      iImage      As Long
      iOrder      As Long
End Type


Public Const LVFI_PARAM = &H1
Public Const LVFI_STRING = &H2
Public Const LVFI_PARTIAL = &H8
Public Const LVFI_WRAP = &H20
Public Const LVFI_NEARESTXY = &H40

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As String) As Long
Public Declare Function GetListViewItemHeight Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As RECT) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Public Sub ShowHeaderIcon(ByVal myListView As ListView, ByVal colNumber As Long, ByVal showImage As Long)

  Dim r As Long
  Dim hHeader As Long
  Dim HD As HD_ITEM
   
  ' Get a handle to the listview header component '
   hHeader = SendMessageLong(myListView.hwnd, LVM_GETHEADER, 0, 0)
   
  ' Set up the required structure members '
  HD.mask = HDI_IMAGE Or HDI_FORMAT
  HD.fmt = HDF_LEFT Or HDF_STRING Or HDF_BITMAP_ON_RIGHT Or showImage
  HD.pszText = myListView.ColumnHeaders(myListView.SortKey + 1).Text
  
  If showImage Then
    
    HD.iImage = myListView.SortOrder
    
  End If
   
  ' Modify the header '
  r = SendMessageAny(hHeader, HDM_SETITEM, colNumber, HD)
   
End Sub
Public Sub SortListView(ByVal myListView As ListView, ByVal colNumber As Long, ByVal DataType As ListDataType, ByVal SortOrder As Boolean)
  
   Dim Counter As Long

    Dim i As Integer
    Dim l As Long
    Dim strFormat As String
    
    ' Display the hourglass cursor whilst sorting
    
    Dim lngCursor As Long
    lngCursor = myListView.MousePointer
    myListView.MousePointer = vbHourglass
    
    ' Prevent the myListView control from updating on screen - this is to hide
    ' the changes being made to the listitems, and also to speed up the sort
    
    LockWindowUpdate myListView.hwnd
    
    Dim blnRestoreFromTag As Boolean
    
    Select Case DataType
    Case ldtString Or ldtTime
        
        ' Sort alphabetically. This is the only sort provided by the
        ' MS myListView control (at this time), and as such we don't really
        ' need to do much here
    
        blnRestoreFromTag = False
        
    Case ldtNumber
    
        ' Sort Numerically
    
        strFormat = String$(20, "0") & "." & String$(10, "0")
        
        ' Loop through the values in this column. Re-format the values so
        ' as they can be sorted alphabetically, having already stored their
        ' text values in the tag, along with the tag's original value
    
        With myListView.ListItems
            If (colNumber = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        .Tag = .Text & Chr$(0) & .Tag
                        If IsNumeric(.Text) Then
                            If CDbl(.Text) >= 0 Then
                                .Text = Format(CDbl(.Text), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(.Text), strFormat))
                            End If
                        Else
                            .Text = ""
                        End If
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .Item(l).ListSubItems(colNumber - 1)
                        .Tag = .Text & Chr$(0) & .Tag
                        If IsNumeric(.Text) Then
                            If CDbl(.Text) >= 0 Then
                                .Text = Format(CDbl(.Text), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(.Text), strFormat))
                            End If
                        Else
                            .Text = ""
                        End If
                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
    
    Case ldtDate
    
        ' Sort by date.
        
        strFormat = "YYYYMMDDHhNnSs"
        
        Dim dte As Date
    
        ' Loop through the values in this column. Re-format the dates so as they
        ' can be sorted alphabetically, having already stored their visible
        ' values in the tag, along with the tag's original value
    
        With myListView.ListItems
            If (colNumber = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .Item(l).ListSubItems(colNumber - 1)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
        
    End Select
    
    ' Sort the myListView Alphabetically
    
    myListView.SortOrder = IIf(SortOrder, lvwAscending, lvwDescending)
    myListView.SortKey = colNumber - 1
    myListView.Sorted = True
    
    ' Restore the Text Values if required
    
    If blnRestoreFromTag Then
        
        ' Restore the previous values to the 'cells' in this column of the list
        ' from the tags, and also restore the tags to their original values
        
        With myListView.ListItems
            If (colNumber = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .Item(l).ListSubItems(colNumber - 1)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End With
                Next l
            End If
        End With
    End If
    
    ' Unlock the list window so that the OCX can update it
    
    LockWindowUpdate 0&
    
    ' Restore the previous cursor
    
    myListView.MousePointer = lngCursor
  
   For Counter = 0 To myListView.ColumnHeaders.Count - 1
     If Counter = myListView.SortKey Then
        Call ShowHeaderIcon(myListView, myListView.SortKey, HDF_IMAGE)
     Else
        Call ShowHeaderIcon(myListView, Counter, 0)
     End If
   Next Counter

End Sub

Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function
Public Function IsLoaded(strFormName As String) As Boolean
    Dim i As Integer


    For i = 0 To Forms.Count - 1


        If (Forms(i).Name = strFormName) Then
            IsLoaded = True
            Exit For
        End If
    Next
End Function
