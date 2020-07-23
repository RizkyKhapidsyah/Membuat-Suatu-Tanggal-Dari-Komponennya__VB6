VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuat Suatu Tanggal dari Komponennya"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim tanggal As Date
  tanggal = DateTime.DateSerial(2020, 10, 22)
  MsgBox tanggal  'Menghasilkan tanggal 22/10/2020
End Sub

