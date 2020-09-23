VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboOccurrence 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":0000
      Left            =   1260
      List            =   "Form1.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4020
      TabIndex        =   9
      Top             =   1380
      Width           =   1035
   End
   Begin VB.TextBox txtHoliday 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      TabIndex        =   8
      Top             =   1440
      Width           =   1515
   End
   Begin VB.ComboBox cboMonth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":004A
      Left            =   4740
      List            =   "Form1.frx":0075
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   1635
   End
   Begin VB.ComboBox cboDay 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":00DD
      Left            =   2760
      List            =   "Form1.frx":00F9
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6660
      TabIndex        =   1
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lblAnswer2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1320
      TabIndex        =   14
      Top             =   1860
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1020
      TabIndex        =   13
      Top             =   1500
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "a holiday?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      TabIndex        =   12
      Top             =   1500
      Width           =   930
   End
   Begin VB.Label lblAnswer 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1320
      TabIndex        =   11
      Top             =   660
      Width           =   75
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4500
      TabIndex        =   6
      Top             =   180
      Width           =   150
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "What is the"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7560
      TabIndex        =   4
      Top             =   180
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "of"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6420
      TabIndex        =   3
      Top             =   180
      Width           =   165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub Command2_Click()
Dim strAnswer As Variant
strAnswer = GetDateByOccurrence(txtYear, cboMonth.ListIndex, cboDay.ListIndex, cboOccurrence.ListIndex)
If strAnswer = False Then
   lblAnswer.Caption = "There is not a 5th occurrence of this day in the month and year."
Else
   lblAnswer.Caption = strAnswer
End If
End Sub

Private Sub Command3_Click()
If IsHoliday(CDate(txtHoliday)) Then
   lblAnswer2.Caption = "This is a holiday! Don't go to work!"
Else
   lblAnswer2.Caption = "This is NOT a holiday. You better call in sick!"
End If
End Sub

