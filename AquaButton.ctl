VERSION 5.00
Begin VB.UserControl AquaButton 
   BackStyle       =   0  'Transparent
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   LockControls    =   -1  'True
   MaskColor       =   &H00FFFFFF&
   Picture         =   "AquaButton.ctx":0000
   ScaleHeight     =   421.875
   ScaleMode       =   0  'User
   ScaleWidth      =   1725
   ToolboxBitmap   =   "AquaButton.ctx":19AF
   Begin VB.Image Image7 
      Height          =   450
      Left            =   1440
      Picture         =   "AquaButton.ctx":1CC1
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "AquaButton"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   1770
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   0
      Picture         =   "AquaButton.ctx":310F
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image6 
      Height          =   450
      Left            =   1440
      Picture         =   "AquaButton.ctx":45F9
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   240
      Picture         =   "AquaButton.ctx":5AD1
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   240
      Picture         =   "AquaButton.ctx":7092
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image8 
      Height          =   450
      Left            =   0
      Picture         =   "AquaButton.ctx":8590
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "AquaButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_Enabled = 0
'Property Variables:
Dim m_Enabled As Boolean
'Event Declarations:
Event Click() 'MappingInfo=Label1,Label1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Label1,Label1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."







Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'========= Original Colour ==============
Image3.Visible = False
Image8.Visible = False
Image7.Visible = False

'========= Blue Colour ==================
Image4.Visible = True
Image5.Visible = True
Image6.Visible = True
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'========= Original Colour ==============
Image3.Visible = True
Image8.Visible = True
Image7.Visible = True

'========= Blue Colour ==================
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
End Sub

Private Sub UserControl_Resize()
Dim n As Integer
Dim m As Integer
n = Image3.Left
m = Image4.Left
Label1.Width = UserControl.ScaleWidth
Image3.Width = UserControl.ScaleWidth
Image4.Width = UserControl.ScaleWidth

'makes sure image7 is always at the right side
Image7.Left = UserControl.ScaleWidth - n
Image6.Left = UserControl.ScaleWidth - m

If UserControl.ScaleWidth = 0 Then
Exit Sub
End If

'won't allow user to adjust the height
UserControl.Height = 450
UserControl.ScaleHeight = 450

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

Private Sub Label1_Click()
    RaiseEvent Click
End Sub

Private Sub Label1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Label1.Caption = PropBag.ReadProperty("Caption", "AquaButton")
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
'    Label1.FontName = PropBag.ReadProperty("FontName", "")
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "AquaButton")
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
'    Call PropBag.WriteProperty("FontName", Label1.FontName, "")
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000008)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Label1,Label1,-1,FontName
'Public Property Get FontName() As String
'    FontName = Label1.FontName
'End Property
'
'Public Property Let FontName(ByVal New_FontName As String)
'    Label1.FontName() = New_FontName
'    PropertyChanged "FontName"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

