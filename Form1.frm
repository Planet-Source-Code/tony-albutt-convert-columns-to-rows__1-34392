VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Columns To Rows"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2460
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Convert Columns to Rows"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ConvertColumnsToRows()

Dim rCounter, cCounter, rCount, cCount, aPos, bPos As Integer
Dim MainString, TempString, RowString As String
Dim cells() As String
MainString = Clipboard.GetText()

aPos = InStr(1, MainString, vbCrLf)
RowString = Left(MainString, aPos)
bPos = 1
'Count Columns
Do While bPos <> 0
bPos = InStr(1, RowString, vbTab)
If bPos <> 0 Then
   RowString = Right(RowString, Len(RowString) - bPos)
   cCount = cCount + 1
End If
Loop
rCount = 0
'Read Data
Do While Len(MainString) <> 0
aPos = InStr(1, MainString, vbCrLf)
RowString = Left(MainString, aPos)
MainString = Right(MainString, Len(MainString) - aPos - 1)
ReDim Preserve cells(cCount, rCount)
   For cCounter = 0 To cCount
      bPos = InStr(1, RowString, vbTab)
      If bPos = 0 Then
         TempString = Left(RowString, Len(RowString) - 1)
      Else
         TempString = Left(RowString, bPos - 1)
      End If
      RowString = Right(RowString, Len(RowString) - bPos)
      cells(cCounter, rCount) = TempString
   Next cCounter
      rCount = rCount + 1
Loop
MainString = ""
   For cCounter = 0 To cCount
      For rCounter = 0 To rCount - 1
         If rCounter = 0 Then
            MainString = MainString & cells(cCounter, rCounter)
         Else
            MainString = MainString & vbTab & cells(cCounter, rCounter)
         End If
      Next rCounter
      MainString = MainString & vbCrLf
   Next cCounter
   ' Clear the contents of the Clipboard
   Clipboard.Clear
   ' Copy New string to Clipboard
   Clipboard.SetText MainString

End Sub

Private Sub Command1_Click()
   ConvertColumnsToRows
End Sub
