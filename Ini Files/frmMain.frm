VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INI file manipulation DEMO"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Values read from the INI file"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4095
      Begin VB.CommandButton cmdRead 
         Caption         =   "&Read"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label lblLong 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label lblString 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Integer value :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "String value :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Values to Save"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   3855
      End
      Begin VB.HScrollBar hscrNumber 
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtString 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Text            =   "Test Value"
         Top             =   345
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Integer value to save :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "String value to save :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Name of the file to be used to save the various infomation
' It can be any file name. You are not restricted to the INI
' extension only.
Private Const INI_FILE As String = "Settings.ini"

' When the user clicks on Read
Private Sub cmdRead_Click()
    Dim FileName As String
    
    FileName = GetIniFileName
    
    lblString.Caption = ReadString("DATA", "StringData", FileName)
    lblLong.Caption = ReadNumber("DATA", "IntegerData", FileName)
End Sub

' When the user clicks Save
Private Sub cmdSave_Click()
    Dim FileName As String
    
    FileName = GetIniFileName
    
    WriteString "DATA", "StringData", txtString.Text, FileName
    WriteString "DATA", "IntegerData", hscrNumber.Value, FileName
End Sub

' Retrieves the full path of the INI file
' In this case we use the application path
' ---------------------------------------
' Note :
' For those who are not familiar with App.Path
' it returns the path where the application is
' stored. If the application is not in the root
' directory e.g. C:\, E:\ etc, it returns a string
' without a backslash ("\") at the end so we need
' to add one when it is not there.
Private Function GetIniFileName() As String
    ' If there is no "\" at the end of the App.Path
    If Right(App.Path, 1) <> "\" Then
        GetIniFileName = App.Path & "\" & INI_FILE
    ' If there is "\" at the end of the App.Path
    Else
        GetIniFileName = App.Path & INI_FILE
    End If
End Function
