Attribute VB_Name = "modTextFormat"
Option Explicit

Public MSWHEEL_ROLLMSG     As Long
Public m_PrevWndProc       As Long
Public Const GWL_WNDPROC = (-4)

Public aControl As VScrollBar

Private Const SB_HORZ As Long = 0
Private Const SB_VERT As Long = 1
Private Const SB_CTL As Long = 2
Private Declare Function SetScrollPos Lib "user32" ( _
     ByVal hwnd As Long, _
     ByVal nBar As Long, _
     ByVal nPos As Long, _
     ByVal bRedraw As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias _
   "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
   ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias _
   "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal msg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long

   If msg = MSWHEEL_ROLLMSG Then
      ' Return if the mouse wheel was rolled up or down
      If wParam > 0 Then 'Mouse_RollUp
        If aControl.Value - aControl.SmallChange < aControl.Min Then
            aControl.Value = aControl.Min
        Else
            aControl.Value = aControl.Value - aControl.SmallChange
        End If
      Else 'Mouse_RollUp
        If aControl.Value + aControl.SmallChange > aControl.Max Then
            aControl.Value = aControl.Max
        Else
            aControl.Value = aControl.Value + aControl.SmallChange
        End If
      End If
      
      
   End If
   WindowProc = CallWindowProc(m_PrevWndProc, hwnd, msg, wParam, lParam)
End Function


