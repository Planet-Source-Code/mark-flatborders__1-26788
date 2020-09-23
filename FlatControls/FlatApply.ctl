VERSION 5.00
Begin VB.UserControl FlatApply 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ScaleHeight     =   1260
   ScaleWidth      =   2940
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2880
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   2880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   2880
      X2              =   2880
      Y1              =   0
      Y2              =   1200
   End
End
Attribute VB_Name = "FlatApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit
'Costruisce la maschera sulla combobox
 Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
 Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
 Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

 Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
 End Type

 Private Type POINTAPI
   X As Long
   Y As Long
 End Type

 Private comboHDC As Long
 Private tpLine As Integer

' Effetti 3D
  Public Enum chgEffect
    inset = 0
    raised = 1
  End Enum

' Definizione variabili colore
  Private COLOR_LEFT   As Long '= vbButtonShadow
  Private COLOR_TOP    As Long '= vbButtonShadow
  Private COLOR_RIGHT  As Long '= vb3DHighlight
  Private COLOR_BOTTON As Long '= vb3DHighlight

' Variabili property per l'effetto 3D : inset di default
  Private Const mDef_Effect As Integer = chgEffect.inset
  Private mProp_Effect As chgEffect
  Private tpEffect As Integer
Private Sub UserControl_Initialize()
' definisce l'impostazione utente
  Select Case mDef_Effect
    Case 0: tpEffect = 0
    Case 1: tpEffect = 1
  End Select
End Sub
Private Sub UserControl_InitProperties()
' Impostazione effetto 3D
  mProp_Effect = mDef_Effect
' Impostazione del colore dei bordi
  COLOR_LEFT = Line1.BorderColor
  COLOR_TOP = Line2.BorderColor
  COLOR_RIGHT = Line3.BorderColor
  COLOR_BOTTON = Line4.BorderColor
End Sub

Private Sub UserControl_Resize()
' Disposizione delle linee nel controllo (Sinistra)
  Line1.X1 = 0
  Line1.X2 = 0
  Line1.Y1 = 0
  Line1.Y2 = UserControl.Height
' Disposizione delle linee nel controllo (Superiore)
  Line2.X1 = 0
  Line2.X2 = UserControl.Width
  Line2.Y1 = 0
  Line2.Y2 = 0
' Disposizione delle linee nel controllo (Destra)
  Line3.X1 = UserControl.Width - 10
  Line3.X2 = UserControl.Width - 10
  Line3.Y1 = 0
  Line3.Y2 = UserControl.Height
' Disposizione delle linee nel controllo (Inferiore)
  Line4.X1 = 0
  Line4.X2 = UserControl.Width
  Line4.Y1 = UserControl.Height - 10
  Line4.Y2 = UserControl.Height - 10
End Sub
'*********************************************
' Colorazione delle linee da parte dell'utente
'*********************************************
Public Property Get line1Left() As OLE_COLOR
Attribute line1Left.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    line1Left = Line1.BorderColor
End Property
Public Property Let line1Left(ByVal New_line1Left As OLE_COLOR)
    Line1.BorderColor() = New_line1Left
    PropertyChanged "line1Left"
End Property
Public Property Get line1Top() As OLE_COLOR
Attribute line1Top.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    line1Top = Line2.BorderColor
End Property
Public Property Let line1Top(ByVal New_line1Top As OLE_COLOR)
    Line2.BorderColor() = New_line1Top
    PropertyChanged "line1Top"
End Property
Public Property Get line2Right() As OLE_COLOR
Attribute line2Right.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    line2Right = Line3.BorderColor
End Property
Public Property Let line2Right(ByVal New_line2Right As OLE_COLOR)
    Line3.BorderColor() = New_line2Right
    PropertyChanged "line2Right"
    COLOR_RIGHT = New_line2Right
End Property
Public Property Get line2Botton() As OLE_COLOR
Attribute line2Botton.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    line2Botton = Line4.BorderColor
End Property
Public Property Let line2Botton(ByVal New_line2Botton As OLE_COLOR)
    Line4.BorderColor() = New_line2Botton
    PropertyChanged "line2Botton"
End Property

'*********************************************
' Applicazione effetto 3D
'*********************************************
Public Property Get Effect() As chgEffect
  Effect = mProp_Effect
End Property
Public Property Let Effect(ByVal fbNewValue As chgEffect)
  mProp_Effect = fbNewValue
  tpEffect = fbNewValue
  applyEffect
  PropertyChanged "Effect"
End Property

' Routine di caricamento proprieta
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Line1.BorderColor = PropBag.ReadProperty("line1Left", vbButtonShadow)
    Line2.BorderColor = PropBag.ReadProperty("line1Top", vbButtonShadow)
    Line3.BorderColor = PropBag.ReadProperty("line2Right", vb3DHighlight)
    Line4.BorderColor = PropBag.ReadProperty("line2Botton", vb3DHighlight)
    mProp_Effect = PropBag.ReadProperty("Effect", mDef_Effect)
    applyProperty
End Sub

' Routine di assegnazione proprieta
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("line1Left", Line1.BorderColor, vbButtonShadow)
    Call PropBag.WriteProperty("line1Top", Line2.BorderColor, vbButtonShadow)
    Call PropBag.WriteProperty("line2Right", Line3.BorderColor, vb3DHighlight)
    Call PropBag.WriteProperty("line2Botton", Line4.BorderColor, vb3DHighlight)
    Call PropBag.WriteProperty("Effect", mProp_Effect, mDef_Effect)
End Sub
Public Sub app3D(vControl As Object)
  
  Dim PControl   As Form      'Il form che contiene il controllo
  Dim AControl   As Object    'variabile di loop
  Dim usrControl As Control   'il controllo utente nel form
  Dim xControl   As Control   'il controllo da gestire
  
  On Error Resume Next        'controlla l'errore di eliminazione
                              'delle cornici in caso di tipi deversi
  vControl.BorderStyle = 0    '(TextBox....)
  vControl.Appearance = 0     '(ListView....)

' Il form che contiene il controllo
  Set PControl = ParentControls.Item(1).Parent
  
  If TypeOf vControl Is TextBox Then
     vControl.BorderStyle = 0
  ElseIf TypeOf vControl Is ListBox Then
'     'MsgBox "e' una ListBox"
'      Dim xControl As ListBox
  ElseIf TypeOf vControl Is ComboBox Then
'     MsgBox "e' una combo"
      comboHDC = vControl.hWnd
      DrawCombo comboHDC
      applyEffect
  Else
'     MsgBox "altra roba"
  End If

' Cicla tutti i controllo cercando quello associato e quello utente
  For Each AControl In ParentControls
    If AControl.Name = vControl.Name Then
        Set xControl = AControl 'Controllo da gestire
    ElseIf AControl.Name = Ambient.DisplayName Then
        Set usrControl = AControl 'Controllo utente
    End If
  Next
  
' posiziona il controllo utente nelle coordinate
' del controllo associato
  usrControl.Left = xControl.Left
  usrControl.Top = xControl.Top
  usrControl.Width = xControl.Width
  usrControl.Height = xControl.Height

' colorazione delle linee
  applyProperty vControl

End Sub
Public Sub applyEffect()
' Inversione della colorazione dei bordi (effetto Inset/Raised)
  Line1.BorderColor = IIf((tpEffect) = 0, COLOR_LEFT, COLOR_RIGHT)
  Line2.BorderColor = IIf((tpEffect) = 0, COLOR_TOP, COLOR_BOTTON)
  Line3.BorderColor = IIf((tpEffect) = 0, COLOR_RIGHT, COLOR_LEFT)
  Line4.BorderColor = IIf((tpEffect) = 0, COLOR_BOTTON, COLOR_TOP)
  If tpEffect = 0 Then
     Line2.ZOrder
     Line1.ZOrder
  Else
     Line4.ZOrder
     Line3.ZOrder
  End If
  DrawCombo comboHDC
End Sub
Sub applyProperty(Optional xControl As Control)
' Assegnazione dei colori impostati dall'utente
  COLOR_LEFT = Line1.BorderColor
  COLOR_TOP = Line2.BorderColor
  COLOR_RIGHT = Line3.BorderColor
  COLOR_BOTTON = Line4.BorderColor
End Sub
'Disegna la nuova cornice per la combo
Private Sub DrawCombo(cbHDC As Long)
 Dim rct As RECT
 Dim cmbDC As Long
'passa l'handle del controllo da ridisegnare
 GetClientRect cbHDC, rct
 cmbDC = GetDC(cbHDC)
 InflateRect rct, -1, -1
 DrawRect cmbDC, rct, vbWhite, vbWhite
 DeleteDC cmbDC
End Sub
Private Function DrawRect(ByVal hdc As Long, ByRef rct As RECT, ByVal oTopLeftColor As OLE_COLOR, ByVal oBottomRightColor As OLE_COLOR)
 Dim hPen As Long
 Dim hPenOld As Long
 Dim tP As POINTAPI
'(Quadrato leftTop del pulsante)
 hPen = CreatePen(20, 0, vbWhite)
'e la memorizza
 hPenOld = SelectObject(hdc, hPen)
'aggiorna la posizione
 MoveToEx hdc, rct.Left, rct.Bottom, tP
'disegna le nuove linee
 LineTo hdc, rct.Left, rct.Top
 LineTo hdc, rct.Right, rct.Top
'riprende la penna
 SelectObject hdc, hPenOld
 DeleteObject hPen
'stesso discorso per la parte bottonRight
 If (rct.Left <> rct.Right) Then
    hPen = CreatePen(20, 0, vbWhite)
    hPenOld = SelectObject(hdc, hPen)
    LineTo hdc, rct.Right, rct.Bottom - 1
    LineTo hdc, rct.Left, rct.Bottom - 1
    SelectObject hdc, hPenOld
    DeleteObject hPen
 End If
End Function
