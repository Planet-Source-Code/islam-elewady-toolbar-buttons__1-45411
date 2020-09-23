VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "about"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Tag             =   "11001"
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   900
      TabIndex        =   0
      Tag             =   "11000"
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Please leave your feedback  at PlanetSourceCode.com"
      Height          =   435
      Left            =   180
      TabIndex        =   3
      Top             =   1200
      Width           =   2715
   End
   Begin VB.Label Label2 
      Caption         =   "email : islam@mshawki.com"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   795
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Author : islam elewady"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
LoadResStrings Me
End Sub
