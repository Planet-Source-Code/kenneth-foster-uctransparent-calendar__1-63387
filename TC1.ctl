VERSION 5.00
Begin VB.UserControl TC1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00FFC0FF&
   ScaleHeight     =   4260
   ScaleWidth      =   4440
   ToolboxBitmap   =   "TC1.ctx":0000
   Begin VB.CommandButton cmdYearDown 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2745
      TabIndex        =   13
      Top             =   105
      Width           =   255
   End
   Begin VB.CommandButton cmdYearUp 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3885
      TabIndex        =   12
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      ScaleHeight     =   285
      ScaleWidth      =   3990
      TabIndex        =   4
      Top             =   495
      Width           =   4020
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3510
         TabIndex        =   11
         Top             =   30
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3045
         TabIndex        =   10
         Top             =   30
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2400
         TabIndex        =   9
         Top             =   30
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1785
         TabIndex        =   8
         Top             =   30
         Width           =   525
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1140
         TabIndex        =   7
         Top             =   30
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   660
         TabIndex        =   6
         Top             =   30
         Width           =   465
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   60
         TabIndex        =   5
         Top             =   30
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCurrentDate 
      Caption         =   "Todays Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2610
      TabIndex        =   3
      Top             =   2910
      Width           =   1485
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2325
      TabIndex        =   2
      Top             =   105
      Width           =   255
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   3645
      Left            =   30
      Top             =   30
      Width           =   4275
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   435
      Left            =   1785
      Shape           =   3  'Circle
      Top             =   3465
      Width           =   435
   End
   Begin VB.Label lblDates 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   -420
      TabIndex        =   0
      Top             =   855
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "TC1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************
'*
'*         Project Name : ucTransparent Calendar
'*        Version Number: 1.0.0
'*           Author Name: Ken Foster
'*                 Date : November 26, 2005
'*        Freeware - Use anyway you want.
'*
'****************************************************

'##    ##  ########  ##    ##
'##   ##   ##        ###   ##
'##  ##    ##        ####  ##
'#####     ######    ## ## ##    ########   #######    ######   ########  ########  ########
'##  ##    ##        ##  ####    ##        ##     ##  ##    ##     ##     ##        ##     ##
'##   ##   ##        ##   ###    ##        ##     ##  ##           ##     ##        ##     ##
'##    ##  ########  ##    ##    ######    ##     ##   ######      ##     ######    ########
'                                ##        ##     ##        ##     ##     ##        ##   ##
'                                ##        ##     ##  ##    ##     ##     ##        ##    ##
'                                ##         #######    ######      ##     ########  ##     ##
'ascii art done with FigWin

'***************** Table of Procedures *************
'   Private Sub UserControl_Initialize
'   Private Sub UserControl_Resize
'   Private Sub cmdCurrentDate_Click
'   Private Sub cmdDown_Click
'   Private Sub cmdUp_Click
'   Private Sub cmdYearDown_Click
'   Private Sub cmdYearUp_Click
'   Private Sub DisplayDates
'   Private Sub ShowDates
'   Private Sub MonthName
'   Private Sub DisplayRefresh
'   Public Property Get FontColor
'   Public Property Let FontColor
'   Public Property Get MonthYearColor
'   Public Property Let MonthYearColor
'   Public Property Get TodayColor
'   Public Property Let TodayColor
'   Private Sub UserControl_ReadProperties
'   Private Sub UserControl_WriteProperties
'***************** End of Table ********************

Dim LastDay As Single
Dim CurMonth As String
Dim CurDay As Single
Dim CurYear As Single
Dim StoreDate As String
Dim sMon As String
Dim CurMonName As String
Dim LastIndex As Single

Const m_def_FontColor = vbBlack
Const m_def_TodayColor = vbRed
Const m_def_MonthYearColor = vbWhite
Const m_def_SundayColor = vbRed
Const m_def_TodayCircle = vbRed
Const m_def_Border = True
Const m_def_BorderColor = vbBlack

Private m_FontColor As OLE_COLOR
Private m_TodayColor As OLE_COLOR
Private m_MonthYearColor As OLE_COLOR
Private m_SundayColor As OLE_COLOR
Private m_TodayCircle As OLE_COLOR
Private m_Border As Boolean
Private m_BorderColor As OLE_COLOR

Private Sub UserControl_Initialize()
   CurMonth = Month(Now)
   MonthName                'go assign month number to month name
   CurYear = Year(Now)
   CurDay = Day(Now)
   DisplayDates
   ShowDates
   DisplayRefresh
End Sub

Private Sub UserControl_Resize()
   'positions first label so alignment is good . Do not change
   lblDates(0).Top = 855
   lblDates(0).Left = -320
   
   UserControl.Width = 4275
   UserControl.Height = 3645
   Shape2.Top = 8
   Shape2.Left = 8
   Shape2.Width = UserControl.Width - 8
   Shape2.Height = UserControl.Height - 8
   
End Sub

Private Sub cmdCurrentDate_Click()
   CurMonth = Month(Now)
   CurYear = Year(Now)
   CurDay = Day(Now)
   Shape1.Visible = True
   DisplayRefresh
End Sub

Private Sub cmdDown_Click()
   CurMonth = CurMonth - 1
   If CurMonth < 1 Then
      CurMonth = 12
      CurYear = CurYear - 1
   End If
   Shape1.Visible = False
   DisplayRefresh
End Sub

Private Sub cmdUp_Click()
   CurMonth = CurMonth + 1
   If CurMonth > 12 Then
      CurMonth = 1
      CurYear = CurYear + 1
   End If
   Shape1.Visible = False
   DisplayRefresh
End Sub

Private Sub cmdYearDown_Click()
   CurYear = CurYear - 1
   Shape1.Visible = False
   DisplayRefresh
End Sub

Private Sub cmdYearUp_Click()
   CurYear = CurYear + 1
   Shape1.Visible = False
   DisplayRefresh
End Sub

Private Sub DisplayDates()
   Dim iRow As Single
   Dim iColumn As Single
   Dim iDates As Single
   Dim CellTop As Single
   Dim CellLeft As Single
   Dim x As Integer
   
   'setup current month layout
   CellTop = lblDates(0).Top
   CellLeft = lblDates(0).Left
   
   For iRow = 1 To 6
      For iColumn = 1 To 7
         If iDates = 38 Then
            ShowDates
            Exit Sub
         End If
         On Error Resume Next
         iDates = iDates + 1
         Load lblDates(iDates)
           lblDates(iDates).Move CellLeft, CellTop
           CellLeft = CellLeft + lblDates(0).Width + 100
         Next
         CellTop = CellTop + lblDates(0).Height + 50
         CellLeft = lblDates(0).Left
      Next
         
End Sub

Private Sub ShowDates()
   Dim StartDay As Single
   Dim ctr As Single
   Dim CheckDates As String
   Dim DateCaption As Single
   
   'show current calendar
   On Error Resume Next
   StartDay = Weekday(CurMonth & "/1/" & CurYear)
   
   For ctr = StartDay To 38
      DateCaption = DateCaption + 1
      CheckDates = Format(CurMonth & "/" & DateCaption & "/" & CurYear, "Short Date")
      
      If Not IsDate(CheckDates) Then
         LastDay = lblDates(ctr - 1).Index
         Exit For
      End If
      
      'set color of Todays date label
      If Day(Now) = lblDates(ctr).Caption And Month(Now) = CurMonth And Year(Now) = CurYear Then
         UserControl.ForeColor = m_TodayColor
         Shape1.Visible = True
         Shape1.Top = lblDates(ctr).Top - 30
         Shape1.Left = lblDates(ctr).Left + 530
      Else
         UserControl.ForeColor = m_FontColor
      End If
      
      lblDates(ctr).Caption = DateCaption
      
      'center the numbers that are less than 10 so they look better
      If lblDates(ctr).Caption < 10 Then
         UserControl.CurrentX = lblDates(ctr).Left + lblDates(0).Width + 100
      Else
          UserControl.CurrentX = lblDates(ctr).Left + lblDates(0).Width
      End If
      
      Dim xDate As String
      'make Sunday labels a different color
      If ctr = 1 Or ctr = 8 Or ctr = 15 Or ctr = 22 Or ctr = 29 Or ctr = 36 Then
      
         'print week number to left of Sunday
         xDate = CurMonth & "/" & lblDates(ctr).Caption & "/" & CurYear
         UserControl.CurrentX = -30
         UserControl.CurrentY = lblDates(ctr).Top
         UserControl.FontSize = 8
         
         ' assures week number does not print in Today color
         If lblDates(ctr).Caption = CurDay And CurYear = Year(Now) And CurMonth = Month(Now) Then UserControl.ForeColor = m_FontColor
         UserControl.Print DatePart("ww", xDate, vbSunday, vbFirstFourDays)
         
         ' re-align date numbers and reset Font size and position
         If lblDates(ctr).Caption < 10 Then
            UserControl.CurrentX = lblDates(ctr).Left + lblDates(0).Width + 100
         Else
            UserControl.CurrentX = lblDates(ctr).Left + lblDates(0).Width
         End If
       
         UserControl.FontSize = 14
         UserControl.ForeColor = m_SundayColor
           ' in case Sunday is the current day print in Today color
           If lblDates(ctr).Caption = CurDay And CurYear = Year(Now) And CurMonth = Month(Now) Then UserControl.ForeColor = m_TodayColor
         lblDay.ForeColor = m_SundayColor
      End If
      
      UserControl.CurrentY = lblDates(ctr).Top
      UserControl.Print DateCaption
    Next
      
      'position and color of Month label
      UserControl.ForeColor = m_MonthYearColor
      UserControl.CurrentX = 600
      UserControl.CurrentY = 50
      UserControl.Print CurMonName
      
      'position of Year label
      UserControl.CurrentX = 3000
      UserControl.CurrentY = 50
      UserControl.Print CurYear
      
      'the following can be deleted to remove my signature
      UserControl.ForeColor = vbRed
      UserControl.FontSize = 8
      UserControl.CurrentX = 200
      UserControl.CurrentY = 3400
      UserControl.Print "by Ken Foster"   'change my name to yours if your going to claim it.
      UserControl.FontSize = 14
      'end here
      
      UserControl.MaskPicture = UserControl.Image
   End Sub

Private Sub MonthName()
   Select Case CurMonth
         Case 1: CurMonName = "January"
         Case 2: CurMonName = "Febuary"
         Case 3: CurMonName = "March"
         Case 4: CurMonName = "April"
         Case 5: CurMonName = "May"
         Case 6: CurMonName = "June"
         Case 7: CurMonName = "July"
         Case 8: CurMonName = "August"
         Case 9: CurMonName = "September"
         Case 10: CurMonName = "October"
         Case 11: CurMonName = "November"
         Case 12: CurMonName = "December"
   End Select
End Sub

Private Sub DisplayRefresh()
   MonthName
   UserControl.Cls
   DisplayDates
   ShowDates
End Sub

Public Property Get Border() As Boolean
   Border = m_Border
End Property

Public Property Let Border(NewBorder As Boolean)
   m_Border = NewBorder
   PropertyChanged "Border"
   Shape2.Visible = m_Border
   DisplayRefresh
End Property

Public Property Get BorderColor() As OLE_COLOR
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(NewBorderColor As OLE_COLOR)
   m_BorderColor = NewBorderColor
   PropertyChanged "BorderColor"
   Shape2.BorderColor = m_BorderColor
   DisplayRefresh
End Property

Public Property Get FontColor() As OLE_COLOR
   FontColor = m_FontColor
End Property

Public Property Let FontColor(NewFontColor As OLE_COLOR)
   m_FontColor = NewFontColor
   PropertyChanged "FontColor"
   DisplayRefresh
End Property

Public Property Get MonthYearColor() As OLE_COLOR
   MonthYearColor = m_MonthYearColor
End Property

Public Property Let MonthYearColor(NewMonthYearColor As OLE_COLOR)
   m_MonthYearColor = NewMonthYearColor
   PropertyChanged "MonthYearColor"
   DisplayRefresh
End Property

Public Property Get TodayColor() As OLE_COLOR
   TodayColor = m_TodayColor
End Property

Public Property Let TodayColor(NewTodayColor As OLE_COLOR)
   m_TodayColor = NewTodayColor
   PropertyChanged "TodayColor"
   DisplayRefresh
End Property

Public Property Get SundayColor() As OLE_COLOR
   SundayColor = m_SundayColor
End Property

Public Property Let SundayColor(NewSundayColor As OLE_COLOR)
   m_SundayColor = NewSundayColor
   PropertyChanged "SundayColor"
   DisplayRefresh
End Property

Public Property Get TodayCircle() As OLE_COLOR
   TodayCircle = m_TodayCircle
End Property

Public Property Let TodayCircle(NewTodayCircle As OLE_COLOR)
   m_TodayCircle = NewTodayCircle
   PropertyChanged "TodayCircle"
   Shape1.BorderColor = m_TodayCircle
   DisplayRefresh
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
   TodayColor = PropBag.ReadProperty("TodayColor", m_def_TodayColor)
   MonthYearColor = PropBag.ReadProperty("MonthYearColor", m_def_MonthYearColor)
   SundayColor = PropBag.ReadProperty("SundayColor", m_def_SundayColor)
   TodayCircle = PropBag.ReadProperty("TodayCircle", m_def_TodayCircle)
   Border = PropBag.ReadProperty("Border", m_def_Border)
   BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      Call .WriteProperty("FontColor", m_FontColor, m_def_FontColor)
      Call .WriteProperty("TodayColor", m_TodayColor, m_def_TodayColor)
      Call .WriteProperty("MonthYearColor", m_MonthYearColor, m_def_MonthYearColor)
      Call .WriteProperty("SundayColor", m_SundayColor, m_def_SundayColor)
      Call .WriteProperty("TodayCircle", m_TodayCircle, m_def_TodayCircle)
      Call .WriteProperty("Border", m_Border, m_def_Border)
     Call .WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)

   End With
End Sub
