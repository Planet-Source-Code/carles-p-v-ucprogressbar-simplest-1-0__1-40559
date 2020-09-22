VERSION 5.00
Begin VB.UserControl ucProgress 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ClipControls    =   0   'False
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   129
End
Attribute VB_Name = "ucProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// Simple Progress bar
'// -------------------
'// Carles P.V. - 2002
'//
'//
'//
'// I've tried to build this progress bar as simple as possible.
'// Control paintings have tried to get the same. Min value has been
'// removed to simplify calcs.
'//
'// About border styles, see:
'// http://www.vbaccelerator.com/codelib/winstyle/ucbstyle.htm
'// http://www.vbsmart.com/library/smartedge/smartedge.htm

Option Explicit

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOREDRAW      As Long = &H8
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_SHOWWINDOW    As Long = &H40

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE         As Long = (-16)
Private Const WS_THICKFRAME     As Long = &H40000
Private Const WS_BORDER         As Long = &H800000
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const WS_EX_WINDOWEDGE  As Long = &H100&
Private Const WS_EX_CLIENTEDGE  As Long = &H200&
Private Const WS_EX_STATICEDGE  As Long = &H20000

'// Public Enums.:
Public Enum pbBorderStyleConstants
    pbNone = 0
    pbThin
    pbThick
End Enum
Public Enum pbOrientationConstants
    pbHorizontal = 0
    pbVertical
End Enum

'// Event Declarations:
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'// Default Property Values:
Private Const m_def_Orientation = [pbHorizontal]
Private Const m_def_BorderStyle = [pbThick]
Private Const m_def_BackColor = vbButtonFace
Private Const m_def_ForeColor = vbHighlight
Private Const m_def_Max = 100

'// Property Variables:
Private m_Orientation As pbOrientationConstants
Private m_BorderStyle As pbBorderStyleConstants
Private m_BackColor   As OLE_COLOR
Private m_ForeColor   As OLE_COLOR
Private m_Max         As Long

'// Private Variables:
Private m_Value       As Long
Private m_PrgForeRect As RECT
Private m_PrgBackRect As RECT
Private m_PrgPos      As Long
Private m_LastPrgPos  As Long
Private m_hForeBrush  As Long
Private m_hBackBrush  As Long



'// -----------------------------------------------------------------------------------
'// Control appearance: Resize / Paint
'// -----------------------------------------------------------------------------------

Private Sub UserControl_Resize()
    pGetProgress
    pCalcRects
    UserControl_Paint
End Sub

Private Sub UserControl_Paint()
    FillRect hdc, m_PrgForeRect, m_hForeBrush
    FillRect hdc, m_PrgBackRect, m_hBackBrush
End Sub

'// -----------------------------------------------------------------------------------
'// Events
'// -----------------------------------------------------------------------------------

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'// -----------------------------------------------------------------------------------
'// Properties
'// -----------------------------------------------------------------------------------

Public Property Get BorderStyle() As pbBorderStyleConstants
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As pbBorderStyleConstants)
    m_BorderStyle = New_BorderStyle
    pSetBorder
    pGetProgress
    pCalcRects
    UserControl_Paint
End Property

Public Property Get Orientation() As pbOrientationConstants
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As pbOrientationConstants)

    m_Orientation = New_Orientation
    
    With Extender
        .Move .Left, .Top, .Height, .Width
    End With
    pGetProgress
    pCalcRects
    UserControl_Paint
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    
    m_BackColor = New_BackColor
    pCreateBackBrush
    UserControl_Paint
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)

    m_ForeColor = New_ForeColor
    pCreateForeBrush
    UserControl_Paint
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    If (New_Max < 1) Then Err.Raise 380
    m_Max = New_Max
End Property

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "400"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)

    m_Value = New_Value
    
    pGetProgress
    If (m_PrgPos <> m_LastPrgPos) Then
        m_LastPrgPos = m_PrgPos
        pCalcRects
        UserControl_Paint
    End If
End Property

'// -----------------------------------------------------------------------------------
'// Init/Read/Write properties / Termination
'// -----------------------------------------------------------------------------------

Private Sub UserControl_InitProperties()

    m_BorderStyle = m_def_BorderStyle
    m_Orientation = m_def_Orientation
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Max = m_def_Max
    
    pSetBorder
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", -1)

    pSetBorder
    pCalcRects
    pCreateForeBrush
    pCreateBackBrush
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, -1)
End Sub

Private Sub UserControl_Terminate()
    If (m_hForeBrush <> 0) Then DeleteObject m_hForeBrush
    If (m_hBackBrush <> 0) Then DeleteObject m_hBackBrush
End Sub

'// -----------------------------------------------------------------------------------
'// Private
'// -----------------------------------------------------------------------------------

Private Sub pCreateForeBrush()
    
  Dim lColor As Long
    
    If (m_hForeBrush <> 0) Then
        DeleteObject m_hForeBrush
        m_hForeBrush = 0
    End If
    TranslateColor ForeColor, 0, lColor
    m_hForeBrush = CreateSolidBrush(lColor)
End Sub

Private Sub pCreateBackBrush()
    
  Dim lColor As Long
  
    If (m_hBackBrush <> 0) Then
        DeleteObject m_hBackBrush
        m_hBackBrush = 0
    End If
    TranslateColor BackColor, 0, lColor
    m_hBackBrush = CreateSolidBrush(lColor)
End Sub

Private Sub pGetProgress()
    
    On Error Resume Next
  
    Select Case m_Orientation
      Case [pbHorizontal]
        m_PrgPos = (m_Value * ScaleWidth) \ m_Max
      Case [pbVertical]
        m_PrgPos = (m_Value * ScaleHeight) \ m_Max
    End Select
    
    On Error GoTo 0
End Sub

Private Sub pCalcRects()
    
    Select Case m_Orientation
      Case [pbHorizontal]
        SetRect m_PrgForeRect, 0, 0, m_PrgPos, ScaleHeight
        SetRect m_PrgBackRect, m_PrgPos, 0, ScaleWidth, ScaleHeight
      Case [pbVertical]
        SetRect m_PrgForeRect, 0, ScaleHeight - m_PrgPos, ScaleWidth, ScaleHeight
        SetRect m_PrgBackRect, 0, 0, ScaleWidth, ScaleHeight - m_PrgPos
    End Select
End Sub

Private Sub pSetBorder()

    Select Case m_BorderStyle
      Case [pbNone]
        pSetWinStyle GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        pSetWinStyle GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
      Case [pbThin]
        pSetWinStyle GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        pSetWinStyle GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
      Case [pbThick]
        pSetWinStyle GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        pSetWinStyle GWL_EXSTYLE, WS_EX_CLIENTEDGE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE
    End Select
End Sub

Private Sub pSetWinStyle(ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)

  Dim lS As Long
    
    lS = GetWindowLong(hWnd, lType)
    lS = (lS And Not lStyleNot) Or lStyle
    SetWindowLong hWnd, lType, lS
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub
