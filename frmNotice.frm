VERSION 5.00
Begin VB.Form frmNotice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notice"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "frmNotice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   465
      Left            =   1890
      TabIndex        =   1
      Top             =   2175
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   $"frmNotice.frx":000C
      Height          =   840
      Left            =   75
      TabIndex        =   2
      Top             =   1125
      Width           =   4740
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmNotice.frx":00FB
      Height          =   915
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4740
   End
End
Attribute VB_Name = "frmNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    ' Close the notice window...
    Unload frmNotice
End Sub
