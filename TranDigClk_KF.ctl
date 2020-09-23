VERSION 5.00
Begin VB.UserControl TraDigClk 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   BeginProperty Font 
      Name            =   "Roman"
      Size            =   9.75
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   MaskColor       =   &H00C0C0FF&
   ScaleHeight     =   900
   ScaleWidth      =   2820
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   495
      Top             =   2490
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "TraDigClk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'****************************************************
'*
'*         Project Name : Digital Date/Clock
'*        Version Number: 1.0.1
'*           Author Name: Ken Foster
'*                 Date : November 03, 2005
'*        Freeware - Use anyway you want.
'*
'****************************************************
'***************** Table of Procedures *************
'   Private Sub Usercontrol_ReadProperties
'   Private Sub UserControl_Resize
'   Private Sub Usercontrol_WriteProperties
'   Private Sub UserControl_Terminate
'   Private Sub Timer1_Timer
'   Private Sub Draw
'   Public Property Get BackColor
'   Public Property Let BackColor
'   Public Property Get Enabled
'   Public Property Let Enabled
'   Public Property Get ForeColor
'   Public Property Let ForeColor
'   Public Property Get Transparent
'   Public Property Let Transparent
'***************** End of Table ********************

   Private m_Enabled As Boolean
   Private m_Transparent As Boolean
   Private m_ForeColor As OLE_COLOR
   Private m_BackColor As OLE_COLOR
   Private m_Border As Boolean
   
   Const m_def_Border = True
   
Private Sub UserControl_Initialize()
   BackColor = vbBlack
   ForeColor = vbWhite
   Transparent = False
   Border = True
End Sub

Private Sub Usercontrol_ReadProperties(Propbag As PropertyBag)
   Enabled = Propbag.ReadProperty("Enabled", False)
   Transparent = Propbag.ReadProperty("Transparent", True)
   ForeColor = Propbag.ReadProperty("ForeColor", vbRed)
   BackColor = Propbag.ReadProperty("BackColor", vbBlack)
   Border = Propbag.ReadProperty("Border", m_def_Border)
End Sub

Private Sub UserControl_Resize()
   UserControl.Width = Shape1.Width
   UserControl.Height = Shape1.Height
   Draw
End Sub

Private Sub Usercontrol_WriteProperties(Propbag As PropertyBag)
   Call Propbag.WriteProperty("Enabled", m_Enabled, False)
   Call Propbag.WriteProperty("Transparent", m_Transparent, True)
   Call Propbag.WriteProperty("ForeColor", m_ForeColor, vbRed)
   Call Propbag.WriteProperty("BackColor", m_BackColor, vbBlack)
   Call Propbag.WriteProperty("Border", m_Border, m_def_Border)
End Sub

Private Sub UserControl_Terminate()
   Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
   Draw
End Sub

Private Sub Draw()
   
   UserControl.Picture = LoadPicture
   
   Dim Dstg As String
   Dim Dte As String
   Dim Tstg As String
   Dim LTstg As Integer
   Dim Tme As String
   Dim Wte As String
   Dim Mte As String
   
   'Format the time layout
   
   Tstg = Format(Now, "HH:MM:SS AMPM")
   LTstg = Len(Tstg)
   Tme = Left$(Tstg, LTstg - 3)      'don't show am or pm with time
   UserControl.Font = "Roman"        'Clock font
   UserControl.FontBold = True
   UserControl.CurrentX = 500
   UserControl.CurrentY = 20
   UserControl.FontSize = 22
   Print Tme
   
   UserControl.Font = "MS Sans Serif" 'Sets AM/PM and Date Font
   UserControl.CurrentX = 2100
   UserControl.CurrentY = 50
   UserControl.FontBold = False
   UserControl.FontSize = 8
   Print Right$(Tstg, 3)             'Display only AM or PM
   
   'Format the date layout
   
   'Day of Month
   Dstg = Format(Date, "MMM DDD DD")
   Dte = Right$(Dstg, 2)             ' get the right-most 2 chars
   If Dte <= 9 Then Dte = Right$(Dstg, 1)     'do not show the leading zero
   UserControl.CurrentX = 200
   UserControl.CurrentY = 50
   Print Dte
   
   'Month
   Mte = Left$(Dstg, 3)              ' get the left-most 3 chars
   UserControl.CurrentX = 50
   UserControl.CurrentY = 250
   Print UCase(Mte)                  ' make upper case
   
   'Day of Week
   Wte = Mid$(Dstg, 4, 5)            ' get middle 3 chars
   UserControl.CurrentX = 2025
   UserControl.CurrentY = 250
   Print UCase(Wte)                  ' make upper case
   
   UserControl.MaskPicture = UserControl.Image
End Sub

Public Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property

Public Property Let BackColor(NewBackColor As OLE_COLOR)
   m_BackColor = NewBackColor
   PropertyChanged "BackColor"
   UserControl.BackColor = m_BackColor
   UserControl.MaskColor = m_BackColor
   Draw
End Property

Public Property Get Border() As Boolean
   Border = m_Border
End Property

Public Property Let Border(ByVal NewBorder As Boolean)
   m_Border = NewBorder
   Shape1.Visible = m_Border
   PropertyChanged "Border"
End Property
Public Property Get Enabled() As Boolean
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
   m_Enabled = NewEnabled
   PropertyChanged "Enabled"
   Timer1.Enabled = m_Enabled
End Property

Public Property Get ForeColor() As OLE_COLOR
   ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(NewForeColor As OLE_COLOR)
   m_ForeColor = NewForeColor
   PropertyChanged "ForeColor"
   UserControl.ForeColor = m_ForeColor
   Shape1.BorderColor = m_ForeColor
   Draw
End Property

Public Property Get Transparent() As Boolean
   Transparent = m_Transparent
End Property

Public Property Let Transparent(NewTransparent As Boolean)
   m_Transparent = NewTransparent
   PropertyChanged "Transparent"
   If m_Transparent = True Then
      UserControl.BackStyle = 0
   Else
      UserControl.BackStyle = 1
   End If
   Draw
End Property
