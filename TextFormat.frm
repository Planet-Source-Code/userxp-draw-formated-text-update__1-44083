VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Draw formatted text"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin Project1.TextFormat TextFormat1 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5895
      _ExtentX        =   9869
      _ExtentY        =   4366
      FormatMethod    =   0
      PrintAreaMaxHeight=   5000
      PointerForLink  =   "TextFormat.frx":0000
      RightMargin     =   10
   End
   Begin VB.TextBox Text1 
      Height          =   1500
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "TextFormat.frx":031A
      Top             =   3720
      Width           =   5940
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set text"
      Height          =   390
      Left            =   60
      TabIndex        =   3
      Top             =   5280
      Width           =   1620
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   3180
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Has border"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   3180
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   4260
      TabIndex        =   5
      Top             =   3120
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private Declare Function GetDC Lib "user32" ( _
     ByVal hwnd As Long) As Long
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

Dim TheX As Single, TheY As Single

Private Sub Check1_Click()
TextFormat1.BorderStyle = IIf(Check1.Value = vbChecked, FixedSingle, None)
End Sub

Private Sub Check2_Click()
TextFormat1.Enabled = Check2.Value = vbChecked
End Sub

Private Sub Command2_Click()
TextFormat1.Text = Text1.Text
End Sub

Private Sub Form_Load()
Dim aTxt As String
aTxt = "Nothing [b] [s12]A bold [s8]" & "string [i]bold italic[/i] bold only[/b] [u] underline [/u] ok format "
aTxt = aTxt & vbNewLine & vbNewLine & "1  " & aTxt & "2 " & aTxt & aTxt & aTxt & aTxt & aTxt & " END" & vbNewLine & vbNewLine & "Some things more to print here, just to check this control. Very good!"

'UserControl11.Text = aTxt

TextFormat1.Text = Text1.Text

TextFormat1.Fonts.Add "Arial"
TextFormat1.Fonts.Add "Times New Roman"
TextFormat1.Fonts.Add "Tahoma"

End Sub

Private Sub Form_Resize()

    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Height < 2900 Then Me.Height = 2900
    If Me.Width < 3200 Then Me.Width = 3200
    
    TextFormat1.Width = Me.ScaleWidth - 2 * TextFormat1.Left
    TextFormat1.Height = (Me.ScaleHeight - 2 * TextFormat1.Top) / 2
    Check2.Top = 2 * TextFormat1.Top + TextFormat1.Height
    Check1.Top = 2 * TextFormat1.Top + TextFormat1.Height
    
    Label1.Top = Check2.Top
    
    Label1.Left = TextFormat1.Left + TextFormat1.Width - Label1.Width
        
    Text1.Top = Check1.Top + Check1.Height + TextFormat1.Top
    Text1.Width = TextFormat1.Width
    Text1.Height = Me.ScaleHeight - TextFormat1.Height - 6 * TextFormat1.Top - Command2.Height - Check2.Height
    
    Command2.Top = Me.ScaleHeight - Command2.Height - TextFormat1.Top
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 65 Then
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End If
End Sub

Private Sub TextFormat1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
TheX = x
TheY = y
End Sub

Private Sub TextFormat1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1 = x & ", " & y
End Sub
