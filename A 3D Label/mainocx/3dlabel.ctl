VERSION 5.00
Begin VB.UserControl label3d 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3210
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   915
   ScaleWidth      =   3210
   Begin VB.Label Label1 
      Caption         =   "Pradeep Singh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   600
      Left            =   1215
      TabIndex        =   0
      Top             =   1710
      Visible         =   0   'False
      Width           =   2265
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "label3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'####################################################################
'#                                                                  #
'#                    3D LABEL ACTIVEX CONTROL                      #
'#                    CREADET BY PRADEEP SINGH                      #
'#               CREATED ON 13/JULY/2001 AT 3:26 PM                 #
'#            MAIL ME AT    pradeepsingh10@hotmail.com              #
'#                                                                  #
'####################################################################

Option Explicit
Event Click()
Public Sub CreatEffect(Shadow As Integer, xCor As Integer, yCor As Integer)
Cls 'Clear the Control text so 3D text don't stick to the Control when user click the control
Dim intWhiteX As Integer
Dim intWhiteY As Integer
Dim Shadowdepth  As Integer
Dim intTemp As Integer
Dim strText As String
strText = Label1.Caption
'Set the Form ScaleMode to pixels
UserControl.ScaleMode = 3
'Create black shadow effect
UserControl.ForeColor = vbBlack
For intTemp = 0 To Shadow
        UserControl.CurrentX = xCor - intTemp
        UserControl.CurrentY = yCor - intTemp
        UserControl.Print strText
Next
'Create yellow text
        UserControl.ForeColor = vbYellow
        UserControl.CurrentX = xCor
        UserControl.CurrentY = yCor
        UserControl.Print strText
End Sub
Private Sub Label1_Change()
    Call CreatEffect(4, 12, 15)
End Sub
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_Initialize()
    Call CreatEffect(4, 12, 15)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CreatEffect(2, 15, 18)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CreatEffect(5, 12, 15)
End Sub
Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property
Public Property Get Caption() As String
    Caption = Label1.Caption
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        Label1.Caption = PropBag.ReadProperty("Caption", "Pradeep")
        Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.UserMode)
        UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
      Call PropBag.WriteProperty("Caption", Label1.Caption, "Pradeep")
      Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.UserMode)
      Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


