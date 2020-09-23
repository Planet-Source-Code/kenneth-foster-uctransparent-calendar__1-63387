VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo ucTransparent Calendar"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin Project1.TraDigClk TraDigClk1 
      Height          =   495
      Left            =   885
      TabIndex        =   1
      Top             =   4545
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      ForeColor       =   8438015
   End
   Begin Project1.TC1 TC11 
      Height          =   3645
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   6429
      FontColor       =   8388736
      TodayColor      =   192
      MonthYearColor  =   12582912
      SundayColor     =   8421631
      BorderColor     =   33023
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
TraDigClk1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
TraDigClk1.Enabled = False
End Sub
