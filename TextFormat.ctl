VERSION 5.00
Begin VB.UserControl TextFormat 
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "TextFormat.ctx":0000
   Begin VB.PictureBox Picture2 
      HasDC           =   0   'False
      Height          =   3435
      Left            =   120
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   1
      Top             =   60
      Width           =   4635
      Begin VB.VScrollBar VScroll1 
         Height          =   2835
         LargeChange     =   30
         Left            =   4200
         Max             =   11
         SmallChange     =   10
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   2055
         Left            =   210
         ScaleHeight     =   133
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   277
         TabIndex        =   0
         Top             =   645
         Width           =   4215
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuCopyFormated 
         Caption         =   "Copy Formated"
      End
   End
   Begin VB.Menu mnuHiddenLink 
      Caption         =   "HiddenLink"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyLink 
         Caption         =   "Copy link"
      End
   End
End
Attribute VB_Name = "TextFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum FormatMethods
    UseBrackets
    UseSlash
End Enum

Private Const BLACKNESS As Long = &H42
Private Const DSTINVERT As Long = &H550009
Private Const MERGECOPY As Long = &HC000CA
Private Const MERGEPAINT As Long = &HBB0226
Private Const NOTSRCCOPY As Long = &H330008
Private Const NOTSRCERASE As Long = &H1100A6
Private Const PATCOPY As Long = &HF00021
Private Const PATINVERT As Long = &H5A0049
Private Const PATPAINT As Long = &HFB0A09
Private Const SRCAND As Long = &H8800C6
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCERASE As Long = &H440328
Private Const SRCINVERT As Long = &H660046
Private Const SRCPAINT As Long = &HEE0086
Private Const WHITENESS As Long = &HFF0062
Private Declare Function StretchBlt Lib "gdi32" ( _
     ByVal hdc As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal hSrcDC As Long, _
     ByVal xSrc As Long, _
     ByVal ySrc As Long, _
     ByVal nSrcWidth As Long, _
     ByVal nSrcHeight As Long, _
     ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" ( _
     ByVal hDestDC As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal hSrcDC As Long, _
     ByVal xSrc As Long, _
     ByVal ySrc As Long, _
     ByVal dwRop As Long) As Long

Enum aBorderStyleConstants
    None = 0
    FixedSingle = 1
End Enum
Private Const DT_ACCEPT_DBCS As Long = (&H20)
Private Const DT_AGENT As Long = (&H3)
Private Const DT_BOTTOM As Long = &H8
Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_CHARSTREAM As Long = 4
Private Const DT_DISPFILE As Long = 6
Private Const DT_DISTLIST As Long = (&H1)
Private Const DT_EDITABLE As Long = (&H2)
Private Const DT_EDITCONTROL As Long = &H2000
Private Const DT_END_ELLIPSIS As Long = &H8000
Private Const DT_EXPANDTABS As Long = &H40
Private Const DT_EXTERNALLEADING As Long = &H200
Private Const DT_FOLDER As Long = (&H1000000)
Private Const DT_FOLDER_LINK As Long = (&H2000000)
Private Const DT_FOLDER_SPECIAL As Long = (&H4000000)
Private Const DT_FORUM As Long = (&H2)
Private Const DT_GLOBAL As Long = (&H20000)
Private Const DT_HIDEPREFIX As Long = &H100000
Private Const DT_INTERNAL As Long = &H1000
Private Const DT_LEFT As Long = &H0
Private Const DT_LOCAL As Long = (&H30000)
Private Const DT_MAILUSER As Long = (&H0)
Private Const DT_METAFILE As Long = 5
Private Const DT_MODIFIABLE As Long = (&H10000)
Private Const DT_MODIFYSTRING As Long = &H10000
Private Const DT_MULTILINE As Long = (&H1)
Private Const DT_NOCLIP As Long = &H100
Private Const DT_NOFULLWIDTHCHARBREAK As Long = &H80000
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_NOT_SPECIFIC As Long = (&H50000)
Private Const DT_ORGANIZATION As Long = (&H4)
Private Const DT_PASSWORD_EDIT As Long = (&H10)
Private Const DT_PATH_ELLIPSIS As Long = &H4000
Private Const DT_PLOTTER As Long = 0
Private Const DT_PREFIXONLY As Long = &H200000
Private Const DT_PRIVATE_DISTLIST As Long = (&H5)
Private Const DT_RASCAMERA As Long = 3
Private Const DT_RASDISPLAY As Long = 1
Private Const DT_RASPRINTER As Long = 2
Private Const DT_REMOTE_MAILUSER As Long = (&H6)
Private Const DT_REQUIRED As Long = (&H4)
Private Const DT_RIGHT As Long = &H2
Private Const DT_RTLREADING As Long = &H20000
Private Const DT_SET_IMMEDIATE As Long = (&H8)
Private Const DT_SET_SELECTION As Long = (&H40)
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_TABSTOP As Long = &H80
Private Const DT_TOP As Long = &H0
Private Const DT_VCENTER As Long = &H4
Private Const DT_WAN As Long = (&H40000)
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Private Const DT_WORDBREAK As Long = &H10
Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type LinkAreaType
    r(1 To 3) As RECT
    Link As String
End Type

Private Declare Function SetBkMode Lib "gdi32" ( _
     ByVal hdc As Long, _
     ByVal nBkMode As Long) As Long
Private Const OPAQUE As Long = 2
Private Const TRANSPARENT As Long = 1
Private Declare Function SetBkColor Lib "gdi32" ( _
     ByVal hdc As Long, _
     ByVal crColor As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" ( _
     ByVal hdc As Long, _
     ByVal lpStr As String, _
     ByVal nCount As Long, _
     lpRect As RECT, _
     ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" ( _
     ByVal hdc As Long, _
     ByVal lpsz As String, _
     ByVal n As Long, _
     lpRect As RECT, _
     ByVal un As Long, _
     lpDrawTextParams As DRAWTEXTPARAMS) As Long

Private Declare Function GetSystemMetrics Lib "user32" ( _
     ByVal nIndex As Long) As Long
Private Const SM_CXVSCROLL As Long = 2

Private Const IDC_HAND As Long = (32649)
Private Const IDC_ARROW As Long = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
     ByVal hInstance As Long, _
     ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" ( _
     ByVal hCursor As Long) As Long

Private Declare Function RegisterWindowMessage Lib "user32" _
   Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private TheText As String

Private stackFormat As New clsStack
Private stackFrom As New clsStack
Private stackLen As New clsStack

Private Links() As LinkAreaType

Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Private UseFormat As FormatMethods
Private MaxHeight As Long, mRightMargin As Long

Private FontCol As New Collection

Public Property Get Fonts() As Collection
    Set Fonts = FontCol
End Property

Public Property Get RightMargin() As Long
    RightMargin = mRightMargin
End Property

Public Property Let RightMargin(NewRightMargin As Long)
    mRightMargin = NewRightMargin
    UserControl_Resize
End Property

Public Property Get PrintAreaMaxHeight() As Long
    PrintAreaMaxHeight = MaxHeight
End Property

Public Property Let PrintAreaMaxHeight(NewPrintAreaMaxHeight As Long)
    MaxHeight = NewPrintAreaMaxHeight
End Property

Public Property Get PointerForLink() As IPictureDisp
    Set PointerForLink = Picture1.MouseIcon
End Property

Public Property Set PointerForLink(NewPointerForLink As IPictureDisp)
    Set Picture1.MouseIcon = NewPointerForLink
End Property

Public Property Get FormatMethod() As FormatMethods
    FormatMethod = UseFormat
End Property

Public Property Let FormatMethod(NewFormatMethod As FormatMethods)
    UseFormat = NewFormatMethod
End Property

Private Sub SubClassHookForm()
   'MSWHEEL_ROLLMSG = RegisterWindowMessage("MSWHEEL_ROLLMSG")
   ' On Windows NT 4.0, Windows 98, and Windows Me, change the above line to
   MSWHEEL_ROLLMSG = &H20A
   m_PrevWndProc = SetWindowLong(Picture1.hwnd, GWL_WNDPROC, _
                                 AddressOf WindowProc)
End Sub

Private Sub SubClassUnHookForm()
   Call SetWindowLong(Picture1.hwnd, GWL_WNDPROC, m_PrevWndProc)
End Sub

Public Property Get BorderStyle() As aBorderStyleConstants
    BorderStyle = Picture2.BorderStyle
End Property

Public Property Let BorderStyle(aVal As aBorderStyleConstants)
    Picture2.BorderStyle = aVal
    UserControl_Resize
End Property

Public Property Get Text() As String
Text = TheText
End Property

Public Property Let Text(new_text As String)
TheText = new_text
DrawTheText
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(IsEnabled As Boolean)
    UserControl.Enabled = IsEnabled
    VScroll1.Enabled = IsEnabled
    DrawTheText
End Property

Public Property Get MenuCopyVisible() As Boolean
    MenuCopyVisible = mnuCopy.Visible
End Property

Public Property Let MenuCopyVisible(IsVisible As Boolean)
    mnuCopy.Visible = IsVisible
End Property

Public Property Get MenuCopyFormatedVisible() As Boolean
    MenuCopyFormatedVisible = mnuCopyFormated.Visible
End Property

Public Property Let MenuCopyFormatedVisible(IsVisible As Boolean)
    mnuCopyFormated.Visible = IsVisible
End Property

Public Sub Refresh()
DrawTheText
End Sub

Private Sub mnuCopyLink_Click()
    Clipboard.Clear
    Clipboard.SetText mnuCopyLink.Tag
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If Not VScroll1.Enabled Then Exit Sub
If KeyCode = vbKeyUp Then
    If VScroll1.Value - VScroll1.SmallChange < 0 Then
        VScroll1.Value = 0
    Else
        VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
    End If
ElseIf KeyCode = vbKeyDown Then
    If VScroll1.Value + VScroll1.SmallChange > VScroll1.Max Then
        VScroll1.Value = VScroll1.Max
    Else
        VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
    End If
ElseIf KeyCode = vbKeyPageUp Then
    If VScroll1.Value - VScroll1.LargeChange < 0 Then
        VScroll1.Value = 0
    Else
        VScroll1.Value = VScroll1.Value - VScroll1.LargeChange
    End If
ElseIf KeyCode = vbKeyPageDown Then
    If VScroll1.Value + VScroll1.LargeChange > VScroll1.Max Then
        VScroll1.Value = VScroll1.Max
    Else
        VScroll1.Value = VScroll1.Value + VScroll1.LargeChange
    End If
ElseIf KeyCode = vbKeyHome Then
    VScroll1.Value = 0
ElseIf KeyCode = vbKeyEnd Then
    VScroll1.Value = VScroll1.Max
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)

If Button = 0 Then
    If MouseAtALinkArea(x, y) = 0 Then
        Call SetCursor(LoadCursor(0, IDC_ARROW))
    Else
        Call SetCursor(LoadCursor(0, IDC_HAND))
    End If
End If
End Sub

Private Function MouseAtALinkArea(x As Single, y As Single) As Long
Dim i As Long, j As Long

For i = 1 To UBound(Links)
    For j = 1 To 3
        If x > Links(i).r(j).Left And x < Links(i).r(j).Right And _
            y > Links(i).r(j).Top And y < Links(i).r(j).Bottom Then
                MouseAtALinkArea = i
                Exit Function
        End If
    Next j
Next i

End Function

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
RaiseEvent MouseUp(Button, Shift, x, y)
If Button = 2 Then
    i = MouseAtALinkArea(x, y)
    If i = 0 Then
        UserControl.PopupMenu mnuHidden
    Else
        mnuCopyLink.Tag = Links(i).Link
        UserControl.PopupMenu mnuHiddenLink
    End If
ElseIf Button = 1 Then
    i = MouseAtALinkArea(x, y)
    If i <> 0 Then
        Dim iret As Long
        ' open URL into the default internet browser
        Const SW_SHOWNORMAL = 1
        iret = ShellExecute(UserControl.Parent.hwnd, vbNullString, Links(i).Link, _
            vbNullString, "", SW_SHOWNORMAL)
    End If
End If
End Sub

Private Sub UserControl_GotFocus()
Picture1.SetFocus
End Sub

Private Sub UserControl_Initialize()
VScroll1.Width = GetSystemMetrics(SM_CXVSCROLL)
Picture1.BorderStyle = 0
Picture2.Left = 0
Picture2.Top = 0
VScroll1.Top = 0
Set aControl = VScroll1
'SubClassHookForm
End Sub

Private Sub UserControl_InitProperties()
MaxHeight = 5000
End Sub

Private Sub UserControl_Resize()
    Picture2.Width = UserControl.ScaleWidth
    Picture2.Height = UserControl.ScaleHeight
    
    Picture1.Left = 0 'IIf(Picture2.BorderStyle = 0, 0, 2)
    Picture1.Top = 0 'IIf(Picture2.BorderStyle = 0, 0, 2)
    If UserControl.ScaleWidth - VScroll1.Width - mRightMargin - IIf(Picture2.BorderStyle = 0, 0, 4) < 0 Then
        Picture1.Width = 0
    Else
        Picture1.Width = UserControl.ScaleWidth - VScroll1.Width - mRightMargin - IIf(Picture2.BorderStyle = 0, 0, 4)
    End If
    VScroll1.Top = 0 'IIf(Picture2.BorderStyle = 0, 0, 2)
    VScroll1.Left = Picture1.Width + mRightMargin
    VScroll1.Height = UserControl.ScaleHeight - IIf(Picture2.BorderStyle = 0, 0, 4)
    DrawTheText
End Sub

Private Sub UserControl_Terminate()
SubClassUnHookForm
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mnuCopyFormated.Visible = PropBag.ReadProperty("MenuCopyFormatedVisible", True)
    mnuCopy.Visible = PropBag.ReadProperty("MenuCopyVisible", True)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    FormatMethod = PropBag.ReadProperty("FormatMethod", 0)
    PrintAreaMaxHeight = PropBag.ReadProperty("PrintAreaMaxHeight", 5000)
    Set PointerForLink = PropBag.ReadProperty("PointerForLink", Nothing)
    RightMargin = PropBag.ReadProperty("RightMargin", 0)
    Text = PropBag.ReadProperty("text", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "MenuCopyFormatedVisible", mnuCopyFormated.Visible, True
    PropBag.WriteProperty "MenuCopyVisible", mnuCopy.Visible, True
    PropBag.WriteProperty "Enabled", UserControl.Enabled, True
    PropBag.WriteProperty "BorderStyle", Picture2.BorderStyle, 1
    PropBag.WriteProperty "FormatMethod", FormatMethod, ""
    PropBag.WriteProperty "PrintAreaMaxHeight", PrintAreaMaxHeight, ""
    PropBag.WriteProperty "PointerForLink", PointerForLink, ""
    PropBag.WriteProperty "RightMargin", RightMargin, ""
    PropBag.WriteProperty "Text", TheText, ""
End Sub

Private Sub VScroll1_Change()
Picture1.Top = -VScroll1.Value
End Sub

Private Sub mnuCopy_Click()
Dim i As Long, printThis As String, aStr As String
Dim nextText As String
    
    If UseFormat = 0 Then
        aStr = Replace(TheText, vbNewLine, "[l]")
    Else
        aStr = Replace(TheText, vbNewLine, "/l ")
    End If
    
    SplitFormat aStr
    For i = 1 To stackLen.stackLevel
        nextText = Mid(aStr, stackFrom.pop, stackLen.pop)
        If UseFormat = 0 Then
            nextText = Replace(nextText, "[[", "[")
            nextText = Replace(nextText, "]]", "]")
        Else
            nextText = Replace(nextText, "//", "/")
        End If
        If stackFormat.pop = "l" Then printThis = printThis & vbNewLine
        printThis = printThis & nextText
    Next i
        
    Clipboard.Clear
    Clipboard.SetText printThis
    
End Sub

Private Sub mnuCopyFormated_Click()
    Clipboard.Clear
    Clipboard.SetText TheText
End Sub

' Use of [ ]
Private Sub SplitFormat(ByRef aStr As String)
Dim pl As Long, pl2 As Long, lastStart As Long
Dim stackFormatNo As New clsStack
Dim stackFromNo As New clsStack
Dim stackLenNo As New clsStack

If UseFormat = UseSlash Then
    SplitFormat2 aStr
    Exit Sub
End If

stackFormat.Clear
stackFrom.Clear
stackLen.Clear
pl = InStr(aStr, "[")
If pl > 1 Then
    While Mid(aStr, pl + 1, 1) = "[" And pl <> 0
        pl = InStr(pl + 2, aStr, "[")
    Wend
    stackFormatNo.push ""
    stackFromNo.push 1
    If pl = 0 Then
        stackLenNo.push Len(aStr)
        pl2 = Len(aStr)
    Else
        stackLenNo.push pl - 1
    End If
    lastStart = 0
End If
Do While pl <> 0
    If Mid(aStr, pl + 1, 1) <> "[" Then
        pl2 = InStr(pl + 1, aStr, "]")
        If Mid(aStr, pl2 + 1, 1) <> "]" Then
            If pl2 <> 0 Then
                If Mid(aStr, pl + 1, 1) = "/" Then
                    'END format
                    stackFormatNo.push Mid(aStr, pl + 2, pl2 - pl - 2)
                Else
                    'START format
                    stackFormatNo.push Mid(aStr, pl + 1, pl2 - pl - 1)
                End If
                stackFromNo.push pl2 + 1
                If lastStart <> 0 Then stackLenNo.push pl - lastStart
                lastStart = pl2 + 1
            End If
        End If
    End If
    pl = InStr(pl + 1, aStr, "[")
    While Mid(aStr, pl + 1, 1) = "[" And pl <> 0
        pl = InStr(pl + 2, aStr, "[")
    Wend
Loop
If pl2 <> Len(aStr) Then ' More text
    stackFormatNo.push ""
    stackFromNo.push pl2 + 1
    stackLenNo.push Len(aStr) - pl2
End If

For pl = 1 To stackFormatNo.stackLevel
    stackFormat.push stackFormatNo.pop
Next pl
For pl = 1 To stackFromNo.stackLevel
    stackFrom.push stackFromNo.pop
Next pl
For pl = 1 To stackLenNo.stackLevel
    stackLen.push stackLenNo.pop
Next pl

End Sub

' Use of /
Private Sub SplitFormat2(ByRef aStr As String)
Dim pl As Long, pl2 As Long, pl3 As Long, lastStart As Long
Dim MustAdd1 As Long
Dim stackFormatNo As New clsStack
Dim stackFromNo As New clsStack
Dim stackLenNo As New clsStack

stackFormat.Clear
stackFrom.Clear
stackLen.Clear
pl = InStr(aStr, "/")
If pl > 1 Then
    While Mid(aStr, pl + 1, 1) = "/" And pl <> 0
        pl = InStr(pl + 2, aStr, "/")
    Wend
    stackFormatNo.push ""
    stackFromNo.push 1
    stackLenNo.push pl - 1
    lastStart = 0
End If
Do While pl <> 0
    pl2 = InStr(pl + 1, aStr, " ")
    MustAdd1 = 1
    pl3 = InStr(pl + 1, aStr, "/")
    If pl3 < pl2 And pl3 <> 0 Then
        pl2 = pl3: MustAdd1 = 0
    End If
    pl3 = InStr(pl + 1, aStr, vbNewLine)
    If pl3 < pl2 And pl3 <> 0 Then pl2 = pl3: MustAdd1 = 0
    
    If pl2 <> 0 Then
        stackFormatNo.push Mid(aStr, pl + 1, pl2 - pl - 1)
        stackFromNo.push pl2 + 1
        If lastStart <> 0 Then stackLenNo.push pl - lastStart
        lastStart = pl2 + MustAdd1
        
        pl = InStr(pl2, aStr, "/")
    Else
        pl = InStr(pl + 1, aStr, "/")
    End If
    
    While Mid(aStr, pl + 1, 1) = "/" And pl <> 0
        pl = InStr(pl + 2, aStr, "/")
    Wend
Loop
If pl2 <> Len(aStr) Then ' More text
    stackFormatNo.push ""
    stackFromNo.push pl2 + 1
    stackLenNo.push Len(aStr) - pl2
End If

For pl = 1 To stackFormatNo.stackLevel
    stackFormat.push stackFormatNo.pop
Next pl
For pl = 1 To stackFromNo.stackLevel
    stackFrom.push stackFromNo.pop
Next pl
For pl = 1 To stackLenNo.stackLevel
    stackLen.push stackLenNo.pop
Next pl

End Sub

Private Sub DrawTheTextOLD()
Dim aStr As String, printThis As String, i As Long, Lines As Long
Dim HeightOf1Line As Long, cHeight As Long, r As RECT, pl As Long
Dim NewprintThis As String, LastprintThis As String, WhatIs As String
Dim OldFontSize As Single, MaxHeightOf1Line As Long, NextMargin As Long
Dim HasPrintSomething As Boolean, textParams As DRAWTEXTPARAMS, LinkAreaNumber As Long
Dim OldForeColor As Long, LeftMargin As Long, LeftMarginNext As Long, WasLink As Boolean
Dim TheFrom As Long, TheLen As Long, aSingle As Single

On Error GoTo ErrHandle

ReDim Links(0) As LinkAreaType

Picture1.Height = MaxHeight

textParams.cbSize = Len(textParams)

If UseFormat = 0 Then
    aStr = Replace(TheText, vbNewLine, "[l]")
Else
    aStr = Replace(TheText, vbNewLine, "/l ")
End If

SplitFormat aStr

Picture1.FontBold = False
Picture1.FontItalic = False
Picture1.FontUnderline = False
OldFontSize = Picture1.FontSize
OldForeColor = Picture1.ForeColor
If Not UserControl.Enabled Then
    Picture1.ForeColor = &H80000011
End If

HeightOf1Line = Picture1.TextHeight("astr")
MaxHeightOf1Line = HeightOf1Line
Picture1.Cls
For i = 1 To stackLen.stackLevel
    TheFrom = stackFrom.pop
    TheLen = stackLen.pop
    printThis = Mid(aStr, TheFrom, TheLen)
    If UseFormat = 0 Then
        printThis = Replace(printThis, "[[", "[")
        printThis = Replace(printThis, "]]", "]")
    Else
        printThis = Replace(printThis, "//", "/")
    End If
    
    WhatIs = stackFormat.pop
    Select Case Left(WhatIs, 1)
        Case "":
        Case "b": Picture1.FontBold = Not Picture1.FontBold
        Case "i": Picture1.FontItalic = Not Picture1.FontItalic
        Case "u": Picture1.FontUnderline = Not Picture1.FontUnderline
        Case "l":
            Picture1.Print ""
            Picture1.CurrentX = LeftMargin
            
            NextMargin = 0
            HasPrintSomething = False
            MaxHeightOf1Line = HeightOf1Line
        Case "m":
            If Mid(WhatIs, 2, 1) = "+" Then
                LeftMarginNext = LeftMargin + Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "-" Then
                LeftMarginNext = LeftMargin + Val(Mid(WhatIs, 2))
            Else
                LeftMarginNext = Val(Mid(WhatIs, 2))
            End If
            NextMargin = 0
            If i = 1 Then
                Picture1.CurrentX = LeftMarginNext
            End If
        Case "n":
            NextMargin = Val(Mid(WhatIs, 2))
        Case "s":
            Picture1.FontSize = Val(Mid(WhatIs, 2))
            HeightOf1Line = Picture1.TextHeight("astr")
            If MaxHeightOf1Line < HeightOf1Line Then MaxHeightOf1Line = HeightOf1Line
        Case "e": ' print a line
            If Val(Mid(WhatIs, 2)) < 2 Then
                Picture1.Line -(Picture1.Width, Picture1.CurrentY)
            Else
                Picture1.Line -(Picture1.Width, Picture1.CurrentY + Val(Mid(WhatIs, 2)) - 1), , BF
            End If
            Picture1.CurrentY = Picture1.CurrentY + 2
            Picture1.CurrentX = LeftMargin
        Case "w": 'web link
            WasLink = Not WasLink
            If WasLink Then
                ReDim Preserve Links(UBound(Links) + 1) As LinkAreaType
                Links(UBound(Links)).Link = printThis
                If UserControl.Enabled Then Picture1.ForeColor = vbBlue
                GoSub DoJodWithLink
                GoTo Nexti
            Else
                If UserControl.Enabled Then Picture1.ForeColor = OldForeColor
                LinkAreaNumber = 0
            End If
        Case "y": ' Change the CurrentY
            Picture1.CurrentY = Picture1.CurrentY + Val(Mid(WhatIs, 2))
        Case "c": ' Color
            If UserControl.Enabled Then Picture1.ForeColor = Val(Mid(WhatIs, 2))
        Case "t": 'Bullet
            aSingle = Picture1.CurrentY
            If Mid(WhatIs, 2, 1) = "2" Then
                Picture1.CurrentY = aSingle + HeightOf1Line / 2.5
                Picture1.CurrentX = Picture1.CurrentX + 5
                Picture1.DrawWidth = 3
                Picture1.Line -(Picture1.CurrentX + HeightOf1Line / 5, Picture1.CurrentY + HeightOf1Line / 5), , BF
            Else
                Picture1.CurrentY = aSingle + HeightOf1Line / 2
                Picture1.CurrentX = Picture1.CurrentX + 6
                Picture1.DrawWidth = 5
                Picture1.Circle (Picture1.CurrentX, Picture1.CurrentY), HeightOf1Line \ 13
            End If
            Picture1.DrawWidth = 1
            Picture1.CurrentX = Picture1.CurrentX + 7
            Picture1.CurrentY = aSingle
    End Select
    If printThis = "" Then GoTo Nexti
    
    r.Left = Picture1.CurrentX
    r.Top = Picture1.CurrentY
    r.Right = Picture1.Width
    r.Bottom = r.Top + HeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (r.Right < Picture1.Width) And (r.Bottom = r.Top + HeightOf1Line) Then
        Picture1.Print printThis;
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            r.Right = Picture1.Width
            r.Bottom = r.Top + HeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((r.Right < Picture1.Width) And (r.Bottom = r.Top + HeightOf1Line)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            Picture1.Print LastprintThis;
            HasPrintSomething = False
            Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
            Picture1.CurrentX = LeftMargin + NextMargin
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            If printThis <> "" Then GoTo here
        Else
here:
            If HasPrintSomething Then
                Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
                Picture1.CurrentX = LeftMargin + NextMargin
            Else
                HasPrintSomething = True
            End If
            r.Left = LeftMargin + NextMargin
            r.Right = Picture1.Width
            r.Top = Picture1.CurrentY
            Call DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                        
            If r.Top = r.Bottom - HeightOf1Line Then
                Picture1.Print printThis;
            Else
                r.Bottom = r.Bottom - HeightOf1Line
                Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                Picture1.CurrentX = LeftMargin + NextMargin
                Picture1.CurrentY = r.Bottom
                Picture1.Print Mid(printThis, textParams.uiLengthDrawn + 1);
            End If
        End If
    End If
Nexti:
    LeftMargin = LeftMarginNext
Next i

If Picture1.CurrentY + HeightOf1Line > r.Bottom Then
    Picture1.Height = Picture1.CurrentY + HeightOf1Line
Else
    Picture1.Height = r.Bottom
End If

If Picture1.Height > UserControl.ScaleHeight Then
    VScroll1.Enabled = True
    VScroll1.Max = Picture1.Height - Picture2.Height + IIf(Picture2.BorderStyle = 0, 0, 4)
    VScroll1.SmallChange = HeightOf1Line
    VScroll1.LargeChange = UserControl.ScaleHeight - IIf(Picture2.BorderStyle = 0, 5, 7)
    If VScroll1.Value >= VScroll1.Min And VScroll1.Value <= VScroll1.Max Then
        VScroll1_Change
    Else
        VScroll1.Value = 0
        Picture1.Top = 0
    End If
Else
    VScroll1.Enabled = False
    Picture1.Top = 0
End If

Picture1.FontSize = OldFontSize
Picture1.ForeColor = OldForeColor
If Not UserControl.Enabled Then
    VScroll1.Enabled = False
End If

Exit Sub

ErrHandle:
Beep
MsgBox Error
Resume Next
Exit Sub

DoJodWithLink:
    If printThis = "" Then Return
    
    r.Left = Picture1.CurrentX
    r.Top = Picture1.CurrentY
    r.Right = Picture1.Width
    r.Bottom = r.Top + HeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (r.Right < Picture1.Width) And (r.Bottom = r.Top + HeightOf1Line) Then
        'If WasLink Then
            LinkAreaNumber = LinkAreaNumber + 1
            Links(UBound(Links)).r(LinkAreaNumber) = r
        'End If
        Picture1.Print printThis;
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            r.Right = Picture1.Width
            r.Bottom = r.Top + HeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((r.Right < Picture1.Width) And (r.Bottom = r.Top + HeightOf1Line)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            'If WasLink Then
                'Calculate rect
                r.Left = Picture1.CurrentX
                r.Top = Picture1.CurrentY
                r.Right = Picture1.Width
                r.Bottom = Picture1.CurrentY + MaxHeightOf1Line
                Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                
                LinkAreaNumber = LinkAreaNumber + 1
                Links(UBound(Links)).r(LinkAreaNumber) = r
            'End If
            Picture1.Print LastprintThis;
            HasPrintSomething = False
            Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
            Picture1.CurrentX = LeftMargin + NextMargin
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            If printThis <> "" Then GoTo hereWithLink
        Else
hereWithLink:
            If HasPrintSomething Then
                Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
                Picture1.CurrentX = LeftMargin + NextMargin
            Else
                HasPrintSomething = True
            End If
            r.Left = LeftMargin + NextMargin
            r.Top = Picture1.CurrentY
            Call DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                        
            If r.Top = r.Bottom - HeightOf1Line Then
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).r(LinkAreaNumber) = r
                'End If
                Picture1.Print printThis;
            Else
                r.Bottom = r.Bottom - HeightOf1Line
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).r(LinkAreaNumber) = r
                'End If
                Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                Picture1.CurrentX = LeftMargin + NextMargin
                Picture1.CurrentY = r.Bottom
                'If WasLink Then
                    'Calculate rect
                    r.Left = Picture1.CurrentX
                    r.Top = Picture1.CurrentY
                    r.Right = Picture1.Width
                    r.Bottom = Picture1.CurrentY + MaxHeightOf1Line
                    Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).r(LinkAreaNumber) = r
                'End If
                Picture1.Print Mid(printThis, textParams.uiLengthDrawn + 1);
            End If
        End If
    End If

Return

End Sub



Private Sub DrawTheText()
Dim aStr As String, printThis As String, i As Long, Lines As Long
Dim HeightOf1Line As Long, cHeight As Long, r As RECT, pl As Long
Dim NewprintThis As String, LastprintThis As String, WhatIs As String
Dim OldFontSize As Single, MaxHeightOf1Line As Long, NextMargin As Long
Dim HasPrintSomething As Boolean, textParams As DRAWTEXTPARAMS, LinkAreaNumber As Long
Dim OldForeColor As Long, LeftMargin As Long, LeftMarginNext As Long, WasLink As Boolean
Dim TheFrom As Long, TheLen As Long, aSingle As Single
Dim OldFontName As String

On Error GoTo ErrHandle

ReDim Links(0) As LinkAreaType

Picture1.Height = MaxHeight

textParams.cbSize = Len(textParams)

If UseFormat = 0 Then
    aStr = Replace(TheText, vbNewLine, "[l]")
Else
    aStr = Replace(TheText, vbNewLine, "/l ")
End If

SplitFormat aStr

Picture1.FontBold = False
Picture1.FontItalic = False
Picture1.FontUnderline = False
OldFontSize = Picture1.FontSize
OldForeColor = Picture1.ForeColor
OldFontName = Picture1.FontName
If Not UserControl.Enabled Then
    Picture1.ForeColor = &H80000011
End If

HeightOf1Line = Picture1.TextHeight("astr")
MaxHeightOf1Line = HeightOf1Line
Picture1.Cls
For i = 1 To stackLen.stackLevel
    TheFrom = stackFrom.pop
    TheLen = stackLen.pop
    printThis = Mid(aStr, TheFrom, TheLen)
    If UseFormat = 0 Then
        printThis = Replace(printThis, "[[", "[")
        printThis = Replace(printThis, "]]", "]")
    Else
        printThis = Replace(printThis, "//", "/")
    End If
    
    WhatIs = stackFormat.pop
    Select Case Left(WhatIs, 1)
        Case "":
        Case "b": Picture1.FontBold = Not Picture1.FontBold
        Case "i": Picture1.FontItalic = Not Picture1.FontItalic
        Case "u": Picture1.FontUnderline = Not Picture1.FontUnderline
        Case "l":
            r.Top = r.Top + MaxHeightOf1Line
            r.Left = LeftMargin
            r.Right = r.Left
            
            NextMargin = 0
            HasPrintSomething = False
            MaxHeightOf1Line = HeightOf1Line
        Case "m":
            If Mid(WhatIs, 2, 1) = "+" Then
                LeftMarginNext = LeftMargin + Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "-" Then
                LeftMarginNext = LeftMargin + Val(Mid(WhatIs, 2))
            Else
                LeftMarginNext = Val(Mid(WhatIs, 2))
            End If
            NextMargin = 0
            If i = 1 Then
                r.Left = LeftMarginNext
                r.Right = LeftMarginNext
            End If
        Case "n":
            NextMargin = Val(Mid(WhatIs, 2))
        Case "s":
            Picture1.FontSize = Val(Mid(WhatIs, 2))
            HeightOf1Line = Picture1.TextHeight("astr")
            If MaxHeightOf1Line < HeightOf1Line Then MaxHeightOf1Line = HeightOf1Line
        Case "e": ' print a line
            Picture1.CurrentY = r.Top
            Picture1.CurrentX = r.Left
            If Val(Mid(WhatIs, 2)) < 2 Then
                Picture1.Line -(Picture1.Width, Picture1.CurrentY)
                r.Top = r.Top + 2
            Else
                Picture1.Line -(Picture1.Width, Picture1.CurrentY + Val(Mid(WhatIs, 2)) - 1), , BF
                r.Top = r.Top + Val(Mid(WhatIs, 2)) + 1
            End If
            r.Left = LeftMargin
        Case "f": 'font
            On Error Resume Next
            If Val(Mid(WhatIs, 2)) < 1 Then
                Picture1.FontName = OldFontName
            Else
                Picture1.FontName = FontCol.Item(Val(Mid(WhatIs, 2)))
            End If
            On Error GoTo ErrHandle
        Case "w": 'web link
            WasLink = Not WasLink
            If WasLink Then
                ReDim Preserve Links(UBound(Links) + 1) As LinkAreaType
                Links(UBound(Links)).Link = printThis
                If UserControl.Enabled Then Picture1.ForeColor = vbBlue
                GoSub DoJodWithLink
                GoTo Nexti
            Else
                If UserControl.Enabled Then Picture1.ForeColor = OldForeColor
                LinkAreaNumber = 0
            End If
        Case "y": ' Change the CurrentY
            r.Top = r.Top + Val(Mid(WhatIs, 2))
        Case "c": ' Color
            If UserControl.Enabled Then Picture1.ForeColor = Val(Mid(WhatIs, 2))
        Case "t": 'Bullet
            aSingle = r.Top
            If Mid(WhatIs, 2, 1) = "2" Then
                Picture1.CurrentY = aSingle + HeightOf1Line / 2.5
                Picture1.CurrentX = r.Left + 5
                Picture1.DrawWidth = 3
                Picture1.Line -(Picture1.CurrentX + HeightOf1Line / 5, Picture1.CurrentY + HeightOf1Line / 5), , BF
            Else
                Picture1.CurrentY = aSingle + HeightOf1Line / 2
                Picture1.CurrentX = r.Left + 6
                Picture1.DrawWidth = 5
                Picture1.Circle (Picture1.CurrentX, Picture1.CurrentY), HeightOf1Line \ 13
            End If
            Picture1.DrawWidth = 1
            r.Left = Picture1.CurrentX + 7
            r.Right = r.Left
            r.Top = aSingle
    End Select
    If printThis = "" Then GoTo Nexti
    
    r.Left = r.Right
    r.Right = Picture1.Width
    r.Bottom = r.Top + HeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (r.Right < Picture1.Width) And (r.Bottom <= r.Top + HeightOf1Line + 2) Then
        Call DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_EDITCONTROL)
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            r.Right = Picture1.Width
            r.Bottom = r.Top + MaxHeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((r.Right < Picture1.Width) And (r.Bottom <= r.Top + MaxHeightOf1Line + 2)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), r, DT_WORDBREAK Or DT_EDITCONTROL)
            HasPrintSomething = False
            r.Top = r.Top + MaxHeightOf1Line
            r.Left = LeftMargin + NextMargin
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            If printThis <> "" Then GoTo here
        Else
here:
            If HasPrintSomething Then
                r.Top = r.Top + MaxHeightOf1Line
                r.Left = LeftMargin + NextMargin
                printThis = LTrim(printThis)
            Else
                HasPrintSomething = True
            End If
            r.Left = LeftMargin + NextMargin
            r.Right = Picture1.Width
            Call DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                        
            If Abs(r.Top - (r.Bottom - HeightOf1Line)) <= 2 Then
                Call DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_EDITCONTROL)
            Else
                r.Bottom = r.Bottom - HeightOf1Line
                Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                r.Left = LeftMargin + NextMargin
                r.Top = r.Bottom
                r.Bottom = r.Bottom + HeightOf1Line
                Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), r, DT_WORDBREAK Or DT_EDITCONTROL)
            End If
        End If
    End If
Nexti:
    LeftMargin = LeftMarginNext
Next i

If Picture1.CurrentY + HeightOf1Line > r.Bottom Then
    Picture1.Height = Picture1.CurrentY + HeightOf1Line
Else
    Picture1.Height = r.Bottom
End If

If Picture1.Height > UserControl.ScaleHeight Then
    VScroll1.Enabled = True
    VScroll1.Max = Picture1.Height - Picture2.Height + IIf(Picture2.BorderStyle = 0, 0, 4)
    VScroll1.SmallChange = HeightOf1Line
    VScroll1.LargeChange = UserControl.ScaleHeight - IIf(Picture2.BorderStyle = 0, 5, 7)
    If VScroll1.Value >= VScroll1.Min And VScroll1.Value <= VScroll1.Max Then
        VScroll1_Change
    Else
        VScroll1.Value = 0
        Picture1.Top = 0
    End If
Else
    VScroll1.Enabled = False
    Picture1.Top = 0
End If

Picture1.FontSize = OldFontSize
Picture1.ForeColor = OldForeColor
Picture1.FontName = OldFontName
If Not UserControl.Enabled Then
    VScroll1.Enabled = False
End If

Exit Sub

ErrHandle:
Beep
MsgBox Error
Resume Next
Exit Sub

DoJodWithLink:
    If printThis = "" Then Return
    
    r.Left = r.Right
    r.Right = Picture1.Width
    r.Bottom = r.Top + MaxHeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (r.Right < Picture1.Width) And (r.Bottom <= r.Top + MaxHeightOf1Line + 2) Then
        'If WasLink Then
            LinkAreaNumber = LinkAreaNumber + 1
            Links(UBound(Links)).r(LinkAreaNumber) = r
        'End If
        Call DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_EDITCONTROL)
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            r.Right = Picture1.Width
            r.Bottom = r.Top + MaxHeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((r.Right < Picture1.Width) And (r.Bottom <= r.Top + MaxHeightOf1Line + 2)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            'If WasLink Then
                'Calculate rect
                r.Right = Picture1.Width
                Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                
                LinkAreaNumber = LinkAreaNumber + 1
                Links(UBound(Links)).r(LinkAreaNumber) = r
            'End If
            Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), r, DT_WORDBREAK Or DT_EDITCONTROL)
            HasPrintSomething = False
            r.Top = r.Top + MaxHeightOf1Line
            r.Left = LeftMargin + NextMargin
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            If printThis <> "" Then GoTo hereWithLink
        Else
hereWithLink:
            If HasPrintSomething Then
                r.Top = r.Top + MaxHeightOf1Line
                r.Left = LeftMargin + NextMargin
                printThis = LTrim(printThis)
            Else
                HasPrintSomething = True
            End If
            r.Left = LeftMargin + NextMargin
            r.Right = Picture1.Width
            Call DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                        
            If Abs(r.Top - (r.Bottom - HeightOf1Line)) <= 2 Then
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).r(LinkAreaNumber) = r
                'End If
                Call DrawText(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_EDITCONTROL)
            Else
                r.Bottom = r.Bottom - HeightOf1Line
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).r(LinkAreaNumber) = r
                'End If
                Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), r, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                r.Left = LeftMargin + NextMargin
                r.Top = r.Bottom
                r.Bottom = r.Top + MaxHeightOf1Line
                Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), r, DT_WORDBREAK Or DT_EDITCONTROL)
                'If WasLink Then
                    'Calculate rect
                    r.Right = Picture1.Width
                    Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), r, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).r(LinkAreaNumber) = r
                'End If
            End If
        End If
    End If

Return

End Sub

