VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock objWinsock 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Sub Form_Load()
    If objWinsock.State <> sckClosed Then objWinsock.Close
    objWinsock.RemoteHost = "www.boulder.nist.gov"
    objWinsock.RemotePort = 13
    objWinsock.LocalPort = 0
    objWinsock.Connect
End Sub
Private Sub objWinsock_Close()
    If objWinsock.State <> 0 Then
        objWinsock.Close
    End If
End Sub

Private Sub objWinsock_DataArrival(ByVal bytesTotal As Long)
    Dim temp As String
    Dim sTimeArray() As String
    Dim currentDate As Date
    Dim currentTime As Date
    Dim lReturn As Long
    Dim lpSystemTime As SYSTEMTIME
    
    temp = String(bytesTotal, " ")
    objWinsock.GetData temp, vbString, bytesTotal
    temp = Replace(temp, Chr(10), "")
    sTimeArray = Split(temp, " ")
    currentDate = CDate(Mid(sTimeArray(1), 4) & "-" & Mid(sTimeArray(1), 1, 2))
    currentTime = currentDate & " " & CDate(sTimeArray(2))
    
    With lpSystemTime
        .wYear = DatePart("yyyy", currentTime)
        .wMonth = DatePart("m", currentTime)
        .wDay = DatePart("d", currentTime)
        .wHour = DatePart("h", currentTime)
        .wMinute = DatePart("n", currentTime)
        .wSecond = DatePart("s", currentTime)
        .wMilliseconds = 0
    End With
    lReturn = SetSystemTime(lpSystemTime)
    Label1.Caption = currentTime & " GMT"
End Sub



