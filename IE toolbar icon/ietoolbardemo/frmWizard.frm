VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7245
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "1002"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "More Options"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   3
      Left            =   -10000
      TabIndex        =   72
      Top             =   0
      Width           =   7335
      Begin VB.OptionButton optLoc 
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   84
         Tag             =   "7009"
         Top             =   3600
         Width           =   4035
      End
      Begin VB.OptionButton optLoc 
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   83
         Tag             =   "7008"
         Top             =   3240
         Width           =   3555
      End
      Begin VB.Frame fraMenuOptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1695
         Left            =   780
         TabIndex        =   78
         Top             =   1380
         Width           =   6135
         Begin VB.TextBox txtMenuStatusBarText 
            Height          =   285
            Left            =   2040
            TabIndex        =   82
            Top             =   600
            Width           =   3435
         End
         Begin VB.TextBox txtMenuText 
            Height          =   285
            Left            =   2040
            TabIndex        =   81
            Top             =   180
            Width           =   3435
         End
         Begin VB.OptionButton optMenu 
            Caption         =   "Option2"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   80
            Tag             =   "7006"
            Top             =   1380
            Width           =   3555
         End
         Begin VB.OptionButton optMenu 
            Caption         =   "Option1"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   79
            Tag             =   "7005"
            Top             =   1080
            Width           =   3315
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   255
            Left            =   240
            TabIndex        =   87
            Tag             =   "7004"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Tag             =   "7003"
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox txtToolTip 
         Height          =   285
         Left            =   2100
         TabIndex        =   77
         Top             =   3900
         Width           =   3315
      End
      Begin VB.CheckBox chkAddToMenu 
         Caption         =   "Check1"
         Height          =   315
         Left            =   360
         TabIndex        =   76
         Tag             =   "7002"
         Top             =   960
         Width           =   3915
      End
      Begin VB.PictureBox picHead 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   4
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   7155
         TabIndex        =   73
         Top             =   0
         Width           =   7155
         Begin VB.Label lblSubTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "lblStep"
            Height          =   390
            Index           =   4
            Left            =   960
            TabIndex        =   75
            Tag             =   "7001"
            Top             =   360
            Width           =   5160
         End
         Begin VB.Label lblHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   74
            Tag             =   "7000"
            Top             =   60
            Width           =   1575
         End
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   255
         Left            =   600
         TabIndex        =   85
         Tag             =   "7010"
         Top             =   3960
         Width           =   975
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Wizard use"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   1
      Left            =   -10000
      TabIndex        =   54
      Top             =   0
      Width           =   7335
      Begin VB.OptionButton optFunction 
         Caption         =   "Option1"
         Height          =   315
         Index           =   0
         Left            =   2040
         TabIndex        =   59
         Tag             =   "1003"
         Top             =   1800
         Width           =   4155
      End
      Begin VB.OptionButton optFunction 
         Caption         =   "Option2"
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   58
         Tag             =   "1004"
         Top             =   2640
         Width           =   4155
      End
      Begin VB.PictureBox picHead 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   0
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   7155
         TabIndex        =   55
         Top             =   0
         Width           =   7155
         Begin VB.Label lblSubTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "lblStep"
            Height          =   390
            Index           =   3
            Left            =   960
            TabIndex        =   57
            Tag             =   "6001"
            Top             =   360
            Width           =   5160
         End
         Begin VB.Label lblHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   56
            Tag             =   "6000"
            Top             =   60
            Width           =   1575
         End
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Button Uninstall"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   5
      Left            =   -10000
      TabIndex        =   13
      Top             =   0
      Width           =   7215
      Begin MSComctlLib.ListView lstButtons 
         Height          =   2775
         Left            =   120
         TabIndex        =   68
         Top             =   960
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.PictureBox picICON 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   4380
         ScaleHeight     =   795
         ScaleWidth      =   975
         TabIndex        =   66
         Top             =   1140
         Width           =   975
      End
      Begin VB.PictureBox picHead 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   3
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   7155
         TabIndex        =   17
         Top             =   0
         Width           =   7155
         Begin VB.Label lblSubTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "lblStep"
            Height          =   330
            Index           =   2
            Left            =   900
            TabIndex        =   23
            Tag             =   "4001"
            Top             =   300
            Width           =   6000
         End
         Begin VB.Label lblHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Tag             =   "4000"
            Top             =   60
            Width           =   1875
         End
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Left            =   3600
         TabIndex        =   71
         Tag             =   "11004"
         Top             =   3180
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Left            =   3600
         TabIndex        =   70
         Tag             =   "11003"
         Top             =   2700
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   3600
         TabIndex        =   69
         Tag             =   "11002"
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Label lblCount 
         Caption         =   "Label7"
         Height          =   255
         Left            =   360
         TabIndex        =   67
         Top             =   3900
         Width           =   2715
      End
      Begin VB.Label lblFunction 
         Caption         =   "Label9"
         Height          =   195
         Left            =   4980
         TabIndex        =   65
         Top             =   3180
         Width           =   2175
      End
      Begin VB.Label lblDefVisible 
         Caption         =   "Label8"
         Height          =   195
         Left            =   4980
         TabIndex        =   64
         Top             =   2760
         Width           =   2115
      End
      Begin VB.Label lblButtonText 
         Caption         =   "Label7"
         Height          =   195
         Left            =   4980
         TabIndex        =   63
         Top             =   2280
         Width           =   1575
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Finished"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   6
      Left            =   -10000
      TabIndex        =   11
      Top             =   0
      Width           =   7215
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4395
         Left            =   0
         Picture         =   "frmWizard.frx":0000
         ScaleHeight     =   4395
         ScaleWidth      =   7215
         TabIndex        =   25
         Top             =   0
         Width           =   7215
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   435
            Left            =   3420
            TabIndex        =   27
            Tag             =   "5001"
            Top             =   3120
            Width           =   3555
         End
         Begin VB.Label lblFinish 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   855
            Left            =   3300
            TabIndex        =   26
            Tag             =   "5000"
            Top             =   540
            Width           =   3615
         End
      End
      Begin VB.Label lblStep 
         Caption         =   "lblStep"
         Height          =   1470
         Index           =   3
         Left            =   2460
         TabIndex        =   12
         Top             =   960
         Width           =   3960
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Button functionality"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   4
      Left            =   -10000
      TabIndex        =   10
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmdBrowseExe 
         Caption         =   "---"
         Height          =   255
         Left            =   5820
         TabIndex        =   62
         Top             =   3600
         Width           =   315
      End
      Begin VB.CommandButton cmdBrowseSc 
         Caption         =   "---"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5820
         TabIndex        =   61
         Top             =   2940
         Width           =   315
      End
      Begin VB.TextBox txtVal 
         Height          =   315
         Index           =   3
         Left            =   2100
         TabIndex        =   52
         Top             =   3540
         Width           =   3675
      End
      Begin VB.TextBox txtVal 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2100
         TabIndex        =   50
         Top             =   2880
         Width           =   3675
      End
      Begin VB.TextBox txtVal 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2100
         TabIndex        =   48
         Top             =   2160
         Width           =   3675
      End
      Begin VB.TextBox txtVal 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2100
         TabIndex        =   46
         Top             =   1380
         Width           =   3675
      End
      Begin VB.OptionButton optFun 
         Caption         =   "Option4"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   44
         Tag             =   "3005"
         Top             =   3180
         Width           =   5535
      End
      Begin VB.OptionButton optFun 
         Caption         =   "Option3"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   43
         Tag             =   "3004"
         Top             =   2580
         Width           =   5295
      End
      Begin VB.OptionButton optFun 
         Caption         =   "Option2"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   42
         Tag             =   "3003"
         Top             =   1800
         Width           =   5295
      End
      Begin VB.OptionButton optFun 
         Caption         =   "Option1"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   41
         Tag             =   "3002"
         Top             =   960
         Width           =   5295
      End
      Begin VB.PictureBox picHead 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   7155
         TabIndex        =   16
         Top             =   0
         Width           =   7155
         Begin VB.Label lblSubTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "lblStep"
            ForeColor       =   &H80000008&
            Height          =   390
            Index           =   1
            Left            =   660
            TabIndex        =   21
            Tag             =   "3001"
            Top             =   300
            Width           =   5940
         End
         Begin VB.Label lblHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Tag             =   "3000"
            Top             =   60
            Width           =   2535
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   315
         Left            =   420
         TabIndex        =   51
         Tag             =   "3009"
         Top             =   3600
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   420
         TabIndex        =   49
         Tag             =   "3008"
         Top             =   2880
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   420
         TabIndex        =   47
         Tag             =   "3007"
         Top             =   2220
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   420
         TabIndex        =   45
         Tag             =   "3006"
         Top             =   1380
         Width           =   1635
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Button properties"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   2
      Left            =   -10000
      TabIndex        =   7
      Top             =   0
      Width           =   7230
      Begin MSComDlg.CommonDialog cd 
         Left            =   3120
         Top             =   3540
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdBrowseO 
         Caption         =   "---"
         Height          =   255
         Index           =   2
         Left            =   5580
         TabIndex        =   40
         Top             =   2940
         Width           =   315
      End
      Begin VB.CommandButton cmdBrowseO 
         Caption         =   "---"
         Height          =   255
         Index           =   1
         Left            =   5580
         TabIndex        =   39
         Top             =   2100
         Width           =   315
      End
      Begin VB.CheckBox chkDef 
         Caption         =   "Check1"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Tag             =   "2005"
         Top             =   3660
         Width           =   2655
      End
      Begin VB.TextBox txtProp 
         Height          =   315
         Index           =   2
         Left            =   1620
         TabIndex        =   33
         Top             =   2910
         Width           =   3855
      End
      Begin VB.TextBox txtProp 
         Height          =   315
         Index           =   1
         Left            =   1620
         TabIndex        =   31
         Top             =   2055
         Width           =   3855
      End
      Begin VB.TextBox txtProp 
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   29
         Top             =   1200
         Width           =   3855
      End
      Begin VB.PictureBox picHead 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   1
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   7155
         TabIndex        =   15
         Top             =   0
         Width           =   7155
         Begin VB.Label lblHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Tag             =   "2000"
            Top             =   60
            Width           =   1575
         End
         Begin VB.Label lblSubTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "lblStep"
            Height          =   390
            Index           =   0
            Left            =   960
            TabIndex        =   18
            Tag             =   "2001"
            Top             =   360
            Width           =   5160
         End
      End
      Begin VB.Label lblSep 
         Caption         =   "Label2"
         Height          =   375
         Index           =   6
         Left            =   1200
         TabIndex        =   38
         Tag             =   "2009"
         Top             =   3960
         Width           =   5895
      End
      Begin VB.Label lblSep 
         Caption         =   "Label2"
         Height          =   375
         Index           =   5
         Left            =   1800
         TabIndex        =   37
         Tag             =   "2008"
         Top             =   3300
         Width           =   5355
      End
      Begin VB.Label lblSep 
         Caption         =   "Label5"
         Height          =   375
         Index           =   4
         Left            =   1800
         TabIndex        =   36
         Tag             =   "2007"
         Top             =   2460
         Width           =   5355
      End
      Begin VB.Label lblSep 
         Caption         =   "Label5"
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   35
         Tag             =   "2006"
         Top             =   1620
         Width           =   5355
      End
      Begin VB.Label lblSep 
         Caption         =   "Label4"
         Height          =   255
         Index           =   2
         Left            =   250
         TabIndex        =   32
         Tag             =   "2004"
         Top             =   2970
         Width           =   1300
      End
      Begin VB.Label lblSep 
         Caption         =   "Label3"
         Height          =   315
         Index           =   1
         Left            =   250
         TabIndex        =   30
         Tag             =   "2003"
         Top             =   2115
         Width           =   1300
      End
      Begin VB.Label lblSep 
         Caption         =   "Label2"
         Height          =   315
         Index           =   0
         Left            =   250
         TabIndex        =   28
         Tag             =   "2002"
         Top             =   1260
         Width           =   1300
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Introduction"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   0
      Left            =   -10000
      TabIndex        =   6
      Top             =   0
      Width           =   7230
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4395
         Left            =   0
         Picture         =   "frmWizard.frx":651EA
         ScaleHeight     =   4395
         ScaleWidth      =   7215
         TabIndex        =   8
         Top             =   0
         Width           =   7215
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Label6"
            Height          =   315
            Left            =   2940
            TabIndex        =   60
            Tag             =   "1007"
            Top             =   3480
            Width           =   2835
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   495
            Left            =   300
            TabIndex        =   24
            Tag             =   "1000"
            Top             =   420
            Width           =   5955
         End
         Begin VB.Label lblNote 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   495
            Left            =   2940
            TabIndex        =   14
            Tag             =   "1006"
            Top             =   2700
            Width           =   4215
         End
         Begin VB.Label lblStep 
            BackStyle       =   0  'Transparent
            Caption         =   "lblStep"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   0
            Left            =   2940
            TabIndex        =   9
            Tag             =   "1001"
            Top             =   1080
            Width           =   3900
         End
      End
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   7245
      TabIndex        =   0
      Top             =   4440
      Width           =   7245
      Begin VB.CommandButton cmdNav 
         Caption         =   "About"
         Height          =   312
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "100"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   312
         Index           =   1
         Left            =   2250
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "101"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Height          =   312
         Index           =   2
         Left            =   3435
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "102"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   312
         Index           =   3
         Left            =   4545
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Tag             =   "103"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Height          =   312
         Index           =   4
         Left            =   5910
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   108
         X2              =   7012
         Y1              =   24
         Y2              =   24
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   108
         X2              =   7012
         Y1              =   0
         Y2              =   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' The wizard is the simplest way i found
'to make use of my library
' the object model is like this
'Toolbar ' a Global Multiuse
'      ----Toolbarbuttons
'      -------ToolBarButton

'author: islam elewady
'date: mar 8 2003
'email: islam@mshawki.com
'vote at www.planetsourcecode.com

'Please have a look at the description of each
'method ,property in the Object Browser

Const NUM_STEPS = 7

Const BTN_ABOUT = 0
Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4


Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_4 = 4
Const STEP_5 = 5
Const STEP_FINISH = 6




Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2


Dim mbFinishOK      As Boolean
Dim mnCurStep       As Integer

Private Sub chkAddToMenu_Click()
If chkAddToMenu.Value = vbChecked Then
   fraMenuOptions.Enabled = True
   txtMenuText.BackColor = vbWindowBackground
   txtMenuStatusBarText.BackColor = vbWindowBackground
   optMenu(0).Enabled = True
   optMenu(1).Enabled = True
Else
   fraMenuOptions.Enabled = False
  txtMenuText.BackColor = vbButtonFace
   txtMenuStatusBarText.BackColor = vbButtonFace
   optMenu(0).Enabled = False
   optMenu(1).Enabled = False
End If
End Sub

Private Sub cmdBrowseExe_Click()
With cd
   .FileName = ""
   .DialogTitle = LoadResString(11005)
   .Filter = "Exe files (*.exe)|*.exe"
   .InitDir = App.Path
   .ShowOpen
   If Len(.FileName) <> 0 Then
      txtVal(3).Text = .FileName
   End If
End With
End Sub
Private Sub cmdBrowseO_Click(Index As Integer)
With cd
   .DialogTitle = LoadResString(11006)
   .FileName = ""
   .Filter = "Icon files (*.ico)|*.ico|All Resource files (*.*)|*.*"
   .DefaultExt = "*.*"
   .InitDir = App.Path
   .ShowOpen
   If Len(.FileName) <> 0 Then
      txtProp(Index).Text = .FileName
   End If
End With
End Sub
Private Sub cmdBrowseSc_Click()
With cd
   .DialogTitle = LoadResString(11007)
   .DefaultExt = "*.*"
   .Filter = "Script files (*.txt)|*.txt|Visual Basic Scripts (*.vbs)|*.vbs|Windows Scripts (*.wsf)|*.wsf|All Resource Files (*.*)|*.*"
   .InitDir = App.Path
   .ShowOpen
   If Len(.FileName) <> 0 Then
      txtVal(2).Text = .FileName
   End If
   
End With
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Dim nAltStep As Integer
    Select Case Index
        Case BTN_ABOUT
            frmAbout.Show 1
        Case BTN_CANCEL
            Unload Me
          
        Case BTN_BACK
            
            nAltStep = GetAltStep(DIR_BACK)
            SetStep nAltStep, DIR_BACK
          
        Case BTN_NEXT
   
            nAltStep = GetAltStep(DIR_NEXT)
            SetStep nAltStep, DIR_NEXT
          
        Case BTN_FINISH
   
            FinishWizard
            Unload Me
            
         
            
        
    End Select
End Sub
Private Function GetAltStep(ByVal nDirection As Integer) As Integer

Dim nAltStep As Integer
Select Case nDirection
   Case DIR_NEXT
      If mnCurStep = STEP_1 _
         And optFunction(1).Value = True Then
            nAltStep = STEP_5
      ElseIf mnCurStep = STEP_4 Then
            nAltStep = STEP_FINISH
      Else
            nAltStep = mnCurStep + 1
      End If
   Case DIR_BACK
      If mnCurStep = STEP_5 _
         And optFunction(1).Value = True Then
            nAltStep = STEP_1
      ElseIf mnCurStep = STEP_FINISH _
         And optFunction(0).Value = True Then
            nAltStep = STEP_4
      Else
         nAltStep = mnCurStep - 1
      End If
End Select
GetAltStep = nAltStep
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        cmdNav_Click BTN_ABOUT
    End If
End Sub

Private Sub SetStep(nStep As Integer, nDirection As Integer)
   mbFinishOK = False
    Select Case nStep
        Case STEP_INTRO
      
        Case STEP_1
            
        Case STEP_2
          
        Case STEP_3
           
        Case STEP_4
        
        Case STEP_5
           
        Case STEP_FINISH
            mbFinishOK = True
        
    End Select
    
    'move to new step
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).Left = 0
    If nStep <> mnCurStep Then
        fraStep(mnCurStep).Left = -10000
    End If
    fraStep(nStep).Enabled = True
  
    
    SetNavBtns nStep
    
End Sub
Private Sub LoadButtonsList()
ietoolbarbuttons.Refresh
lstButtons.ListItems.Clear
lblCount.Caption = Replace$(LoadResString(11008), "'%'", ietoolbarbuttons.Count)
Dim aToolBarButton As CIEToolbarButton
Dim oLstItem As ListItem
For Each aToolBarButton In ietoolbarbuttons

   Set oLstItem = lstButtons.ListItems.Add
   oLstItem.Text = aToolBarButton.Text
   oLstItem.Tag = aToolBarButton.Guid
   
Next
lstButtons_Click

End Sub
Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    
    If mnCurStep = 0 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub
Private Sub FinishWizard()


Dim i As Integer
Dim oButton As CIEToolbarButton
Set oButton = New CIEToolbarButton

If optFunction(0).Value = True Then

       With oButton
         .Text = txtProp(0).Text
         .HotIcon = txtProp(1).Text
         .Icon = txtProp(2).Text
         .DefaultVisible = chkDef.Value
         If chkAddToMenu.Value = vbChecked Then
            .MenuText = txtMenuText
            .MenuStatusBarText = txtMenuStatusBarText
         End If
         If optMenu(0).Value = True Then
            .AddToMenu = ToolsMenu
         Else
            .AddToMenu = HelpMenu
         End If
         If optLoc(0).Value = True Then
            .CurrentUserButton = False
         Else
            .CurrentUserButton = True
         End If
         If Len(Trim$(txtToolTip.Text)) > 0 Then
            .ToolTip = txtToolTip.Text
         End If
            
         For i = 0 To optFun.Count - 1
               If optFun(i).Value = True Then
                  .Functionlity = i + 1
                  .Value = txtVal(i).Text
                  Exit For
               End If
         Next i
         ietoolbarbuttons.InstallButton oButton
         
      End With

Else

   For i = 1 To lstButtons.ListItems.Count
      If lstButtons.ListItems(i).Checked Then
            oButton.Value = lstButtons.ListItems(i).Tag
            ietoolbarbuttons.UninstallButton oButton
      End If
   Next i

End If



End Sub
Private Sub Form_Load()
 Dim i As Integer
    'init all vars
    mbFinishOK = False
    For i = 0 To NUM_STEPS - 1
      fraStep(i).Left = -10000
    Next
    
    'Load All string info for Form
    LoadResStrings Me
    'Load wizard fonts and layout
    LoadWizardDesign
    'Load Defaults
    LoadDefaults
    'Determine 1st Step:
    SetStep 0, DIR_NONE
End Sub
Private Sub LoadDefaults()
  optFunction(0).Value = True
  txtProp(0).Text = "Run My Calculator"
  txtProp(1).Text = App.Path & "\calc.ico"
  txtProp(2).Text = App.Path & "\calc.ico"
  chkDef.Value = vbChecked
  chkAddToMenu.Value = vbUnchecked
  Call chkAddToMenu_Click
  optFun(3).Value = True
  txtVal(3).Text = "calc.exe"
  Call optFun_Click(3)
  optMenu(0).Value = True
  optLoc(0).Value = True
  
  LoadButtonsList
  
End Sub
Sub LoadWizardDesign()
'On Error Resume Next
'i allow the design to be loaded from
'resource so that font names are not hard coded
Dim ctl As Control
Const resTitle = 10000
Const resHeader = 10001
Const resSubTitle = 10002
Const resText = 10003

Dim sTextFont As String, sTitleFont As String, _
sHeaderFont As String, sSubTitleFont As String

sTextFont = LoadResString(resText)
sTitleFont = LoadResString(resTitle)
sHeaderFont = LoadResString(resHeader)
sSubTitleFont = LoadResString(resSubTitle)
   
   For Each ctl In Me.Controls
      If TypeOf ctl Is Label Or _
         TypeOf ctl Is OptionButton Or _
         TypeOf ctl Is Frame Or _
         TypeOf ctl Is PictureBox Or _
         TypeOf ctl Is CheckBox Then
         If ctl.Name = "lblTitle" Then
            ctl.FontName = sTitleFont
            ctl.FontBold = True
            ctl.FontSize = 12
         ElseIf ctl.Name = "lblHeader" Then
            ctl.FontName = sHeaderFont
            ctl.FontBold = True
            ctl.FontSize = 9
            ctl.Move 250, 100
         ElseIf ctl.Name = "lblSubTitle" Then
            ctl.FontName = sSubTitleFont
            ctl.FontBold = False
            ctl.FontSize = 9
            ctl.Move 400, 400
         ElseIf ctl.Name = "lblFinish" Then
            ctl.FontName = sTitleFont
            ctl.FontBold = True
            ctl.FontSize = 12
         ElseIf ctl.Name = "picHead" Then
            ctl.Move 0, 0, fraStep(ctl.Index).Width, 850
         Else
            ctl.FontName = sTextFont
            ctl.FontBold = False
            ctl.FontSize = 9
         End If
      End If
   Next
   
   
   
End Sub


Private Function FunctionToText(ByVal fIn As enumIEButtonFunctions) As String
Dim sFunString As String
Select Case fIn
   Case ComObject
      sFunString = LoadResString(3002)
   Case ExecutableFile
      sFunString = LoadResString(3005)
   Case ExplorerBar
      sFunString = LoadResString(3003)
   Case ScriptFile
      sFunString = LoadResString(3004)
End Select
   FunctionToText = sFunString
   
   
End Function
Private Sub lstButtons_Click()
'On Error Resume Next
Dim aButton As CIEToolbarButton
Dim picIco As StdPicture
If lstButtons.ListItems.Count = 0 Then
   Exit Sub
End If
Set aButton = ietoolbarbuttons(lstButtons.SelectedItem.Tag)

lblButtonText.Caption = aButton.Text
lblFunction.Caption = FunctionToText(aButton.Functionlity)
lblDefVisible.Caption = IIf(aButton.DefaultVisible, "Yes", "No")
Set picICON.Picture = LoadPicture()
Set picICON.Picture = aButton.GetHotIconPic
'picICON.AutoRedraw = True

End Sub

Private Sub optFun_Click(Index As Integer)
Dim i As Integer
For i = 0 To optFun.Count - 1
   If Index <> i Then
      txtVal(i).Enabled = Not optFun(Index).Value
      txtVal(i).BackColor = vbButtonFace
   Else
      txtVal(i).Enabled = optFun(Index).Value
      txtVal(i).BackColor = vbWindowBackground
   End If
Next i
cmdBrowseSc.Enabled = False
cmdBrowseExe.Enabled = False
'txtVal(Index).SetFocus
txtVal(Index).SelStart = 0
txtVal(Index).SelLength = Len(txtVal(Index).Text)
If Index = 2 Then
   cmdBrowseSc.Enabled = True
End If
If Index = 3 Then
   cmdBrowseExe.Enabled = True
End If
End Sub
Private Sub txtProp_GotFocus(Index As Integer)
txtProp(Index).SelStart = 0
txtProp(Index).SelLength = Len(txtProp(Index).Text)
End Sub
Private Sub txtVal_GotFocus(Index As Integer)
txtVal(Index).SelStart = 0
txtVal(Index).SelLength = Len(txtVal(Index).Text)

End Sub
