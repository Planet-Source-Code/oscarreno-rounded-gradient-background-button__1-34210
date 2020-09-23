VERSION 5.00
Begin VB.UserControl Boton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   ScaleHeight     =   855
   ScaleWidth      =   2955
   ToolboxBitmap   =   "Boton.ctx":0000
   Begin VB.PictureBox Pic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Botón"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   930
      TabIndex        =   0
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "Boton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim sh As Long ' ScaleHeight
Dim sw As Long ' ScaleWidth
Type Colores
    Red As Long
    Green As Long
    Blue As Long
End Type
Dim HasFocus As Boolean
Dim IsDown As Boolean

Dim MouseIn As Boolean

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Ocurre cuando el usuario presiona una tecla mientras un objeto tiene el enfoque."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Ocurre cuando el usuario presiona y libera una tecla ANSI."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Ocurre cuando el usuario libera una tecla mientras un objeto tiene el enfoque."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Ocurre cuando el usuario mueve el mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Ocurre cuando el usuario libera el botón del mouse mientras un objeto tiene el enfoque."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Ocurre cuando se muestra un formulario por primera vez o cuando cambia el tamaño de un objeto."
'Default Property Values:
Const m_def_Default = False
Const m_def_Cancel = False
Const m_def_RoundSize = 20
'Property Variables:
Dim m_Default As Boolean
Dim m_Cancel As Boolean
Dim m_RoundSize As Long





Private Sub Label_Click()
    'Call UserControl_Click
End Sub

Private Sub Label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub



Private Sub Pic_Click()
    'Call UserControl_Click
End Sub

Private Sub Pic_KeyDown(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub Pic_KeyPress(KeyAscii As Integer)
    Call UserControl_KeyPress(KeyAscii)
End Sub

Private Sub Pic_KeyUp(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyUp(KeyCode, Shift)
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_EnterFocus()
    HasFocus = True
    PaintGradient True, False
End Sub

Private Sub UserControl_ExitFocus()
    HasFocus = False
    PaintGradient
End Sub

Private Sub UserControl_GotFocus()
    HasFocus = True
End Sub

Private Sub UserControl_LostFocus()
    HasFocus = False
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    Dim hRgn&
    Dim w As Long
    Dim h As Long
    
    ' Create the round shape
    w = UserControl.Width \ 15
    h = UserControl.Height \ 15
    hRgn& = CreateRoundRectRgn(0, 0, w, h, RoundSize, RoundSize)
    SetWindowRgn UserControl.hWnd, hRgn, True
    
    ' Paint the Gradient
    PaintGradient
    
    ' Colocate controls
    Acomodate
    
End Sub

Private Sub Acomodate()
    Dim plus As Long
    ' If the button is down, then put the controls a bit more down
    plus = IIf(IsDown, 15, 0)
    'If plus = 0 Then Stop
    ' Acommodate the controls
    Label.Top = ((UserControl.Height - Label.Height) \ 2) + plus
    If Pic.Picture <> 0 Then
        Pic.Top = ((UserControl.Height - Pic.Height) \ 2) + plus
        Pic.Left = ((UserControl.Width - Pic.Width - Label.Width - 50) \ 2) + plus
        Label.Left = Pic.Left + Pic.Width + 50 + plus
    Else
        Label.Left = ((UserControl.Width - Label.Width) \ 2) + plus
    End If

End Sub

Private Sub UserControl_Show()
    PaintGradient
End Sub


Private Sub PaintGradient(Optional inFocus As Boolean, Optional inverse As Boolean)
    Const ColorIncrement As Integer = 25
    ' Three user-type variables to store separated RGB values of a color
    Dim ColorRGB1 As Colores
    Dim ColorRGB2 As Colores
    Dim ColorRGB3 As Colores
    ' The number of steps of the cycle
    Dim Pasos As Long
    
    ' working variables
    Dim Red As Double
    Dim Green As Double
    Dim Blue As Double
    Dim i As Long
    Dim start As Long
    Dim finish As Long
    Dim forSteps As Long
    
    ' increment variables
    Dim incR1 As Double
    Dim incG1 As Double
    Dim incB1 As Double
    Dim incR2 As Double
    Dim incG2 As Double
    Dim incB2 As Double
    
    Dim osh As Long ' Original ScaleHeight
    Dim osw As Long ' Original ScaleWidth
    
    
    'This started with a Joshua Foster's idea on PSC, but I changed it a lot!!
    
    'First assign the steps the cycle will do.
    ' I put 1/10th of the SH
    Pasos = UserControl.ScaleHeight / 10
    
    ' Store ocx's original scales and assign news
    osh = UserControl.ScaleHeight
    osw = UserControl.ScaleWidth
    UserControl.ScaleHeight = Pasos
    UserControl.ScaleWidth = 1
    
    ' Then we assign the separated RGB values to the middle color
    ColorRGB2 = Long2RGB(UserControl.BackColor)
    
    ' If the inFocus parameter is on, or the control has the focus, then
    ' let's light the control a bit more
    If inFocus Or HasFocus Then
        ColorRGB2.Red = ColorRGB2.Red + (ColorIncrement \ 2)
        ColorRGB2.Green = ColorRGB2.Green + (ColorIncrement \ 2)
        ColorRGB2.Blue = ColorRGB2.Blue + (ColorIncrement \ 2)
    End If
    
    ' In base on that colors, let's make the top color a bit more light
    ColorRGB1.Blue = ColorRGB2.Blue + ColorIncrement
    ColorRGB1.Green = ColorRGB2.Green + ColorIncrement
    ColorRGB1.Red = ColorRGB2.Red + ColorIncrement
    ' And we do the same with the bottom color, only darker
    ColorRGB3.Blue = ColorRGB2.Blue - ColorIncrement
    ColorRGB3.Green = ColorRGB2.Green - ColorIncrement
    ColorRGB3.Red = ColorRGB2.Red - ColorIncrement
    
    'Just prevents that no negative numbers exists in the bottom
    ColorRGB3.Blue = IIf(ColorRGB3.Blue < 0, 0, ColorRGB3.Blue)
    ColorRGB3.Green = IIf(ColorRGB3.Green < 0, 0, ColorRGB3.Green)
    ColorRGB3.Red = IIf(ColorRGB3.Red < 0, 0, ColorRGB3.Red)
    
    
    ' If the inverse parameter is on (when a button is pressed)
    If inverse Then
        start = Pasos
        finish = 0
        forSteps = -1
    Else
        start = 0
        finish = Pasos
        forSteps = 1
    End If
    
    'Calculate the increment factor of each RGB color
    incR1 = (ColorRGB2.Red - ColorRGB1.Red) / (Pasos \ 2)
    incG1 = (ColorRGB2.Green - ColorRGB1.Green) / (Pasos \ 2)
    incB1 = (ColorRGB2.Blue - ColorRGB1.Blue) / (Pasos \ 2)
    incR2 = (ColorRGB3.Red - ColorRGB2.Red) / (Pasos \ 2)
    incG2 = (ColorRGB3.Green - ColorRGB2.Green) / (Pasos \ 2)
    incB2 = (ColorRGB3.Blue - ColorRGB2.Blue) / (Pasos \ 2)
    
    ' Assign the first color to work variables
    Red = ColorRGB1.Red
    Green = ColorRGB1.Green
    Blue = ColorRGB1.Blue
    
    ' Let's PLAY!
    For i = start To finish Step forSteps
        ' draw colored lines
        UserControl.Line (0, i)-(1, i), RGB(Red, Green, Blue)
        ' calculate the increment
        If i < (Pasos \ 2) Then ' we are in the middle color
          Red = Red + incR1
          Green = Green + incG1
          Blue = Blue + incB1
        Else ' now we are at the bottom
          Red = Red + incR2
          Green = Green + incG2
          Blue = Blue + incB2
        End If
        ' No negatives!
        Red = IIf(Red < 0, 0, Red)
        Green = IIf(Green < 0, 0, Green)
        Blue = IIf(Blue < 0, 0, Blue)
    Next i
    ' return to original scale
    UserControl.ScaleHeight = osh
    UserControl.ScaleWidth = osw
End Sub

' Converts a Color(Long) to RGB values
Private Function Long2RGB(ByVal LongRGB As Long) As Colores
    Long2RGB.Red = LongRGB And 255
    Long2RGB.Green = (LongRGB \ 256) And 255
    Long2RGB.Blue = (LongRGB \ 65536) And 255
End Function


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label,Label,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Devuelve o establece el texto mostrado en la barra de título de un objeto o bajo el icono de un objeto."
    Caption = Label.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label.Caption() = New_Caption
    UserControl_Resize
    PropertyChanged "Caption"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label,Label,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = Label.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    HasFocus = True
    If (KeyCode = vbKeySpace) Or _
       (KeyCode = vbKeyEscape And Cancel = True) Or _
       (KeyCode = vbKeyReturn And Default = True) Then
            IsDown = True
            HasFocus = True
            PaintGradient False, True
            Acomodate
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeySpace) Or _
       (KeyCode = vbKeyEscape And Cancel = True) Or _
       (KeyCode = vbKeyReturn And Default = True) Then
            IsDown = False
            HasFocus = True
            PaintGradient
            Acomodate
            RaiseEvent Click
    End If
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HasFocus = True
    If Button = vbLeftButton Then
        IsDown = True
        Acomodate
        PaintGradient True, True
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Establece un icono personalizado para el mouse."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Thanks to Rick Bull for the lovely idea

    If Button = vbLeftButton Then
        If IsHot(UserControl.hWnd) Or IsHot(Pic.hWnd) Then
            HasFocus = True
            IsDown = True
        Else
            IsDown = False
            HasFocus = False
        End If
        Acomodate
        PaintGradient HasFocus, IsDown
    End If
    
    'PaintGradient True, False
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Devuelve o establece el tipo de puntero del mouse mostrado al pasar por encima de un objeto."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        PaintGradient
        IsDown = False
        Acomodate
        If IsHot(UserControl.hWnd) Or IsHot(Pic.hWnd) Then RaiseEvent Click
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Pic,Pic,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Devuelve o establece el gráfico que se mostrará en un control."
    Set Picture = Pic.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Pic.Picture = New_Picture
    If New_Picture Is Nothing Then
        Pic.Visible = False
    Else
        Pic.Visible = True
    End If
    UserControl_Resize
    PropertyChanged "Picture"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Obliga a volver a dibujar un objeto."
    UserControl.Refresh
End Sub
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=8,0,0,11842740
'Public Property Get Color() As Long
'    Color = m_Color
'End Property
'
'Public Property Let Color(ByVal New_Color As Long)
'    m_Color = New_Color
'    PropertyChanged "Color"
'End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
'    m_Color = m_def_Color
    m_RoundSize = m_def_RoundSize
    m_Default = m_def_Default
    m_Cancel = m_def_Cancel
    UserControl.BackColor = CorrectColor(UserControl.BackColor)
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label.Caption = PropBag.ReadProperty("Caption", "Botón")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", CorrectColor(&H80000009))
    Label.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    m_RoundSize = PropBag.ReadProperty("RoundSize", m_def_RoundSize)
    m_Default = PropBag.ReadProperty("Default", m_def_Default)
    m_Cancel = PropBag.ReadProperty("Cancel", m_def_Cancel)
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", Label.Caption, "Label")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Label.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000009)
    Call PropBag.WriteProperty("ForeColor", Label.ForeColor, &H80000008)
    Call PropBag.WriteProperty("RoundSize", m_RoundSize, m_def_RoundSize)
    Call PropBag.WriteProperty("Default", m_Default, m_def_Default)
    Call PropBag.WriteProperty("Cancel", m_Cancel, m_def_Cancel)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    New_BackColor = CorrectColor(New_BackColor)
    UserControl.BackColor() = New_BackColor
    Pic.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Private Function CorrectColor(color As Long) As Long
    If color < -1 Then  ' If Is a System color
         color = GetSysColor(color And &HFF&)
    End If
    CorrectColor = color

End Function

Private Function ConvertToSysColor(ByVal lColor As Long) As Long
' THIS ROUTINE WAS SUBMITED in PSC BY Rocky Clark's Color Coder 3.0
'Find a system color that matches lColor

Dim lIdx As Long
Dim sHex As String

    If lColor < 0 Then
        'Already a system color
        ConvertToSysColor = lColor
    Else
        For lIdx = 0 To 24
            If GetSysColor(lIdx) = lColor Then
                'Found a match
                sHex = Hex$(lIdx)
                If Len(sHex) < 2 Then
                    sHex = "0" & sHex
                End If
                ConvertToSysColor = Val("&H800000" & sHex)
                Exit For
            End If
        Next
        If lIdx > 24 Then
            'Didn't find a match
            ConvertToSysColor = -1
        End If
    End If
    
End Function



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label,Label,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
    ForeColor = Label.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,20
Public Property Get RoundSize() As Long
Attribute RoundSize.VB_Description = "Tamaño de la esquina redondeada"
    RoundSize = m_RoundSize
End Property

Public Property Let RoundSize(ByVal New_RoundSize As Long)
    m_RoundSize = New_RoundSize
    UserControl_Resize
    PropertyChanged "RoundSize"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,false
Public Property Get Default() As Boolean
    Default = m_Default
End Property

Public Property Let Default(ByVal New_Default As Boolean)
    m_Default = New_Default
    PropertyChanged "Default"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,false
Public Property Get Cancel() As Boolean
    Cancel = m_Cancel
End Property

Public Property Let Cancel(ByVal New_Cancel As Boolean)
    m_Cancel = New_Cancel
    PropertyChanged "Cancel"
End Property

