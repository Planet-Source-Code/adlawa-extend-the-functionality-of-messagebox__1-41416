VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMsgBox 
   Caption         =   "Custom Message Box Designer"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   Icon            =   "frmmsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   9210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMsgBox 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   120
      TabIndex        =   46
      Top             =   5880
      Width           =   9015
      Begin VB.CommandButton cmdCode 
         Caption         =   "&Genrate Code So i Can Used it in My Projects too immediately"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   54
         Top             =   720
         Width           =   3060
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   53
         Top             =   120
         Width           =   1500
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View MsgBox"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   52
         Top             =   120
         Width           =   1500
      End
      Begin VB.Frame fraMsgBox 
         Caption         =   "Miscellaneous"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   7
         Left            =   4200
         TabIndex        =   47
         Top             =   120
         Width           =   4695
         Begin VB.CheckBox chkMsgBox 
            Caption         =   "vbMsgBoxHelpButton"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkMsgBox 
            Caption         =   "vbMsgBoxRight"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   50
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox chkMsgBox 
            Caption         =   "vbMsgBoxRtlReading"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   49
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox chkMsgBox 
            Caption         =   "vbMsgBoxSetForeground"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   48
            Top             =   600
            Width           =   2175
         End
      End
   End
   Begin VB.Frame fraMsgBox 
      Caption         =   "Button Style and Button Default"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   120
      TabIndex        =   39
      Top             =   5040
      Width           =   9015
      Begin VB.ComboBox cmbButtons 
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2655
      End
      Begin VB.PictureBox picButtonFrame 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   3
         Left            =   4200
         ScaleHeight     =   615
         ScaleWidth      =   4635
         TabIndex        =   40
         Top             =   240
         Width           =   4635
         Begin VB.PictureBox picButtonFrame 
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   2
            Left            =   3000
            ScaleHeight     =   615
            ScaleWidth      =   1350
            TabIndex        =   43
            Top             =   -10
            Width           =   1350
            Begin VB.CommandButton MsgBoxButtons 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   2
               Left            =   50
               TabIndex        =   12
               Top             =   50
               Width           =   1215
            End
         End
         Begin VB.PictureBox picButtonFrame 
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   1
            Left            =   1560
            ScaleHeight     =   615
            ScaleWidth      =   1350
            TabIndex        =   42
            Top             =   -10
            Width           =   1350
            Begin VB.CommandButton MsgBoxButtons 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   1
               Left            =   50
               TabIndex        =   11
               Top             =   50
               Width           =   1215
            End
         End
         Begin VB.PictureBox picButtonFrame 
            Height          =   615
            Index           =   0
            Left            =   -10
            ScaleHeight     =   555
            ScaleWidth      =   1290
            TabIndex        =   41
            Top             =   -10
            Width           =   1350
            Begin VB.CommandButton MsgBoxButtons 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   0
               Left            =   50
               TabIndex        =   10
               Top             =   50
               Width           =   1215
            End
         End
      End
   End
   Begin VB.Frame fraMsgBox 
      Caption         =   "Custom Button Captions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   5
      Left            =   4320
      TabIndex        =   37
      Top             =   1560
      Width           =   4335
      Begin VB.TextBox txtButtonCaption 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtButtonCaption 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtButtonCaption 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   18
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox chkButtonCaption 
         Caption         =   "Button 1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkButtonCaption 
         Caption         =   "Button 2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkButtonCaption 
         Caption         =   "Button 3"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame fraMsgBox 
      Caption         =   "Special Display Features"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   6
      Left            =   4320
      TabIndex        =   32
      Top             =   3480
      Width           =   4815
      Begin VB.TextBox txtCheckBoxText 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   44
         Text            =   "CheckBox Text"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chkSpecial 
         Caption         =   "Positioned"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   350
         Width           =   1455
      End
      Begin VB.CheckBox chkSpecial 
         Caption         =   "Self Closing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   725
         Width           =   1455
      End
      Begin VB.CheckBox chkSpecial 
         Caption         =   "Add Checkbox"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   1100
         Width           =   1455
      End
      Begin VB.TextBox txtPosition 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   20
         Top             =   335
         Width           =   615
      End
      Begin VB.TextBox txtPosition 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   21
         Top             =   335
         Width           =   615
      End
      Begin VB.TextBox txtTimeOut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   23
         Top             =   725
         Width           =   615
      End
      Begin VB.Label lblPositionID 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   36
         Top             =   375
         Width           =   105
      End
      Begin VB.Label lblPositionID 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   34
         Top             =   375
         Width           =   105
      End
      Begin VB.Label lblTimeOutID 
         AutoSize        =   -1  'True
         Caption         =   "Timeout  in Seconds"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   33
         Top             =   750
         Width           =   1425
      End
   End
   Begin VB.Frame fraMsgBox 
      Caption         =   "Icon Styles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   31
      Top             =   3480
      Width           =   4095
      Begin VB.PictureBox picIconFrame 
         BorderStyle     =   0  'None
         Height          =   745
         Index           =   5
         Left            =   120
         ScaleHeight     =   750
         ScaleWidth      =   3765
         TabIndex        =   38
         Top             =   240
         Width           =   3770
         Begin VB.PictureBox picIconFrame 
            BorderStyle     =   0  'None
            Height          =   735
            Index           =   4
            Left            =   2985
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   28
            Top             =   -10
            Width           =   735
            Begin VB.CommandButton cmdIcons 
               Height          =   600
               Index           =   4
               Left            =   50
               Picture         =   "frmmsg.frx":000C
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   50
               Width           =   600
            End
         End
         Begin VB.PictureBox picIconFrame 
            BorderStyle     =   0  'None
            Height          =   735
            Index           =   3
            Left            =   2235
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   27
            Top             =   -10
            Width           =   735
            Begin VB.CommandButton cmdIcons 
               Height          =   600
               Index           =   3
               Left            =   50
               Picture         =   "frmmsg.frx":044E
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   50
               Width           =   600
            End
         End
         Begin VB.PictureBox picIconFrame 
            BorderStyle     =   0  'None
            Height          =   735
            Index           =   2
            Left            =   1485
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   26
            Top             =   -10
            Width           =   735
            Begin VB.CommandButton cmdIcons 
               Height          =   600
               Index           =   2
               Left            =   50
               Picture         =   "frmmsg.frx":0890
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   50
               Width           =   600
            End
         End
         Begin VB.PictureBox picIconFrame 
            BorderStyle     =   0  'None
            Height          =   735
            Index           =   1
            Left            =   750
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   25
            Top             =   -10
            Width           =   735
            Begin VB.CommandButton cmdIcons 
               Height          =   600
               Index           =   1
               Left            =   50
               Picture         =   "frmmsg.frx":0CD2
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   50
               Width           =   600
            End
         End
         Begin VB.PictureBox picIconFrame 
            Height          =   735
            Index           =   0
            Left            =   -10
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   35
            Top             =   -10
            Width           =   735
            Begin VB.CommandButton cmdIcons 
               Height          =   600
               Index           =   0
               Left            =   50
               Style           =   1  'Graphical
               TabIndex        =   2
               Top             =   50
               Width           =   600
            End
         End
      End
   End
   Begin VB.Frame fraMsgBox 
      Caption         =   "Display Styles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   4320
      TabIndex        =   30
      Top             =   360
      Width           =   4335
      Begin VB.OptionButton optMsgBox 
         Caption         =   "Application Modal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optMsgBox 
         Caption         =   "System Modal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame fraMsgBox 
      Caption         =   "Message Box Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   9015
      Begin VB.TextBox txtMsgBoxText 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Text            =   "frmmsg.frx":1114
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtMsgBoxTitle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Text            =   "Message Box Title"
         Top             =   360
         Width           =   3495
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   45
      Top             =   7170
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Button Clicked: "
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  frmMsgBox
'  form used to allow visual creation of msgbox
'=========================================================================================
'  Created By: Amer
'  Published Date: 02/23/2001
'  Legal Copyright: Amer © 02/23/2001
'=========================================================================================
Option Explicit
Dim MsgBoxIcon As Integer
Dim MsgBoxModal As Integer
Dim MsgBoxButtonStyle As Integer
Dim MsgBoxDefaultButton As Integer
Dim MsgBoxMiscellaneous As Long
Dim MsgBoxButtonCount As String
Dim MsgBoxOriginalButtonCaptions As String
Dim TimeOut As Integer
Dim OldCaption As String
Dim X As Long
Dim Y As Long
Dim CheckBoxValue As Boolean
Dim CheckBoxText As String
'=========================================================================================
Private Sub chkButtonCaption_Click(Index As Integer)
  On Error Resume Next
  txtButtonCaption(Index).Enabled = Not txtButtonCaption(Index).Enabled
  If txtButtonCaption(Index).Enabled = True Then
    txtButtonCaption(Index).SetFocus
    txtButtonCaption(Index).SelLength = Len(txtButtonCaption(Index).Text)
    OldCaption = txtButtonCaption(Index).Text
  End If
End Sub 'chkButtonCaption_Click(Index As Integer)
'=========================================================================================
Private Sub chkMsgBox_Click(Index As Integer)
  ' routine to add miscellaneous effects to msgbox
  On Error Resume Next
  Dim Counter As Integer
  MsgBoxMiscellaneous = 0
  For Counter = 0 To 3
    If chkMsgBox(Counter).Value = vbChecked Then
      Select Case Counter
        Case 0 ' vbMsgBoxHelpButton
          MsgBoxMiscellaneous = MsgBoxMiscellaneous + 16384
        Case 1 ' vbMsgBoxRight
          MsgBoxMiscellaneous = MsgBoxMiscellaneous + 524288
        Case 2 ' vbMsgBoxRtlReading
          MsgBoxMiscellaneous = MsgBoxMiscellaneous + 1048576
        Case 3 ' vbMsgBoxSetForeground
          MsgBoxMiscellaneous = MsgBoxMiscellaneous + 65536
      End Select
    End If
  Next Counter
End Sub 'chkMsgBox_Click(Index As Integer)
'=========================================================================================
Private Sub chkSpecial_Click(Index As Integer)
' routine to enable/disable controls
  On Error Resume Next
  txtPosition(0).Enabled = chkSpecial(0).Value
  txtPosition(1).Enabled = chkSpecial(0).Value
  txtTimeOut.Enabled = chkSpecial(1).Value
  txtCheckBoxText.Enabled = chkSpecial(2).Value
  If txtPosition(Index).Enabled = True Then
    txtPosition(Index).SetFocus
    txtPosition(Index).SelLength = Len(txtPosition(Index).Text)
  End If
  If txtTimeOut.Enabled = True Then
    txtTimeOut.SetFocus
    txtTimeOut.SelLength = Len(txtTimeOut.Text)
  End If
  If txtCheckBoxText.Enabled = True Then
    txtCheckBoxText.SetFocus
    txtCheckBoxText.SelLength = Len(txtCheckBoxText.Text)
  End If
End Sub 'chkSpecial_Click(Index As Integer)
'=========================================================================================
'=========================================================================================
Private Sub cmdCode_Click()
' routine to generate the code and show the code form
    
  MsgBox "Donot Forget to Copy CustomMsg Module to your Projects" & vbCrLf _
        & "Before Using Syntex Here is Syntex For Custom Message Box" & vbCrLf _
        & "Module With Your selected Options and Settings" & vbCrLf _
        & "Have A Good day And Vote if You Satisfied" & vbCrLf _
        & "Thanks ", vbInformation, "Msgbox"
  On Error Resume Next
  Load frmCode
  If NeedSpecialCode = True Then
    frmCode.txtCode.Text = IIf(chkSpecial(2).Value = vbChecked, "Dim CheckBoxValue As Boolean" & vbCrLf, "") & "CreateSpecialMsgBox " & Chr(34) & GetMsgBoxText(True) & Chr(34) & "," & GetMsgBoxOptions & "," & Chr(34) & GetMsgBoxTitle & Chr(34) & "," & TimeOut & "," & X & "," & Y & "," & Chr(34) & CheckBoxText & Chr(34) & IIf(chkSpecial(2).Value = vbChecked, ", CheckBoxValue,", ",,") & MsgBoxButtonCount & "," & Chr(34) & GetMsgBoxButtonCaptions & Chr(34) & "," & Chr(34) & MsgBoxOriginalButtonCaptions & Chr(34)
       ' MsgBox "Added modCustomMsgBox to your project", vbInformation, "Added Required Module"
     
  Else
    frmCode.txtCode.Text = "MsgBox " & Chr(34) & GetMsgBoxText(True) & Chr(34) & "," & GetMsgBoxOptions & "," & Chr(34) & GetMsgBoxTitle & Chr(34)
  End If
  frmCode.Show vbModal, Me
End Sub 'cmdCode_Click()
'=========================================================================================
Private Sub cmdExit_Click()
  ' quitting time
  On Error Resume Next

  Unload frmMsgBox
  Set frmMsgBox = Nothing
End Sub 'cmdExit_Click()
'=========================================================================================
Private Sub cmdView_Click()
  ' routine to show created msgbox
  On Error Resume Next
  Dim ButtonClicked As Integer
  sbStatus.SimpleText = "Button Clicked: "
  If NeedSpecialCode = True Then
    ButtonClicked = CreateSpecialMsgbox(GetMsgBoxText(False), GetMsgBoxOptions, GetMsgBoxTitle, TimeOut, X, Y, CheckBoxText, CheckBoxValue, MsgBoxButtonCount, GetMsgBoxButtonCaptions, MsgBoxOriginalButtonCaptions)
  Else
    ButtonClicked = MsgBox(GetMsgBoxText(False), GetMsgBoxOptions, GetMsgBoxTitle)
  End If
  sbStatus.SimpleText = "Button Clicked: " & Chr(34) & GetButtonClicked(ButtonClicked) & Chr(34) & IIf(NeedSpecialCode = True, " CheckBox Value: " & Chr(34) & CheckBoxValue & Chr(34), "")
End Sub 'cmdView_Click()
'=========================================================================================
Private Sub Form_Load()
  ' routine used to load combobox and select starting entry
  On Error Resume Next
  With cmbButtons
    .AddItem "vbAbortRetryIgnore"
    .AddItem "vbOKCancel"
    .AddItem "vbOKOnly"
    .AddItem "vbRetryCancel"
    .AddItem "vbYesNo"
    .AddItem "vbYesNoCancel"
  End With
  cmbButtons.ListIndex = 2
 ' Me.Width = 7560
  Me.Refresh
End Sub 'Form_Load()
'=========================================================================================
Private Sub cmdIcons_Click(Index As Integer)
  ' routine to show whats been selected
  On Error Resume Next
  Dim Counter As Integer
  For Counter = 0 To 4
    If Counter <> Index Then
      picIconFrame(Counter).BorderStyle = 0
    Else
      picIconFrame(Index).BorderStyle = 1
    End If
  Next Counter
  ' set the value for the icon wanted
  Select Case Index
    Case 0 ' no image
      MsgBoxIcon = 0
    Case 1 ' critical
      MsgBoxIcon = 16
    Case 2 ' exclamation
      MsgBoxIcon = 48
    Case 3 ' information
      MsgBoxIcon = 64
    Case 4 ' question
      MsgBoxIcon = 32
  End Select
End Sub 'cmdIcons_Click(Index As Integer)
'=========================================================================================
Private Sub Form_Unload(Cancel As Integer)
' quitting time
  cmdExit_Click
End Sub 'Form_Unload(Cancel As Integer)
'=========================================================================================
Private Sub MsgBoxButtons_Click(Index As Integer)
  ' routine to show whats been selected
  On Error Resume Next
  Dim Counter As Integer
  For Counter = 0 To 2
    If Counter <> Index Then
      picButtonFrame(Counter).BorderStyle = 0
    Else
      picButtonFrame(Index).BorderStyle = 1
    End If
  Next Counter
  ' set the value for the default button wanted
  Select Case Index
    Case 0 ' first button
      MsgBoxDefaultButton = 0
    Case 1 ' second button
      MsgBoxDefaultButton = 256
    Case 2 ' third button
      MsgBoxDefaultButton = 512
  End Select
End Sub 'MsgBoxButtons_Click(Index As Integer)
'=========================================================================================
Private Sub cmbButtons_Click()
  ' routine to find out what buttons needed and enable for selection of default
  On Error Resume Next
  Dim Counter As Integer
  Dim ButtonString() As String
  Dim TempString As String

  For Counter = 0 To 2
    MsgBoxButtons(Counter).Enabled = False
    chkButtonCaption(Counter).Enabled = False
  Next Counter

  Select Case cmbButtons.ListIndex
    Case 0 'vbAbortRetryIgnore
      TempString = "&Abort,&Retry,&Ignore"
      MsgBoxOriginalButtonCaptions = TempString
      MsgBoxButtonStyle = 2
      MsgBoxButtonCount = 3
    Case 1 'vbOKCancel
      TempString = "OK,Cancel,"
      MsgBoxOriginalButtonCaptions = "OK,Cancel"
      MsgBoxButtonStyle = 1
      MsgBoxButtonCount = 2
    Case 2 'vbOKOnly
      TempString = "OK,,"
      MsgBoxOriginalButtonCaptions = "OK"
      MsgBoxButtonStyle = 0
      MsgBoxButtonCount = 1
    Case 3 'vbRetryCancel
      TempString = "&Retry,Cancel,"
      MsgBoxOriginalButtonCaptions = "&Retry,Cancel"
      MsgBoxButtonStyle = 5
      MsgBoxButtonCount = 2
    Case 4 'vbYesNo
      TempString = "&Yes,&No,"
      MsgBoxOriginalButtonCaptions = "&Yes,&No"
      MsgBoxButtonStyle = 4
      MsgBoxButtonCount = 2
    Case 5 'vbYesNoCancel
      TempString = "&Yes,&No,Cancel"
      MsgBoxOriginalButtonCaptions = TempString
      MsgBoxButtonStyle = 3
      MsgBoxButtonCount = 3
  End Select
  ButtonString = Split(TempString, ",")
  MsgBoxButtons(0).Caption = ButtonString(0)
  MsgBoxButtons(1).Caption = ButtonString(1)
  MsgBoxButtons(2).Caption = ButtonString(2)
  txtButtonCaption(0).Text = ButtonString(0)
  txtButtonCaption(1).Text = ButtonString(1)
  txtButtonCaption(2).Text = ButtonString(2)
  If MsgBoxButtonCount > 1 Then
    For Counter = 0 To MsgBoxButtonCount - 1
      MsgBoxButtons(Counter).Enabled = True
    Next Counter
  End If
  For Counter = 0 To MsgBoxButtonCount - 1
    chkButtonCaption(Counter).Enabled = True
  Next Counter
  MsgBoxButtons_Click 0
End Sub 'cmbButtons_Click()
'=========================================================================================
Private Function GetButtonClicked(ButtonClicked As Integer) As String
' routine to determine what button has been clicked
  On Error Resume Next
  Select Case ButtonClicked
    Case 1
      GetButtonClicked = "vbOK"
    Case 2
      GetButtonClicked = "vbCancel"
    Case 3
      GetButtonClicked = "vbAbort"
    Case 4
      GetButtonClicked = "vbRetry"
    Case 5
      GetButtonClicked = "vbIgnore"
    Case 6
      GetButtonClicked = "vbYes"
    Case 7
      GetButtonClicked = "vbNo"
  End Select
End Function 'GetButtonClicked(ButtonClicked As Integer) As String
'=========================================================================================
Private Function GetMsgBoxOptions() As Long
' routine to generate a number that is based on user selection for msgbox display styles
  On Error Resume Next
  GetMsgBoxOptions = MsgBoxIcon + MsgBoxModal + MsgBoxButtonStyle + MsgBoxDefaultButton + MsgBoxMiscellaneous
End Function 'GetMsgBoxOptions() As Long
'=========================================================================================
Private Function GetMsgBoxText(ViewCode As Boolean) As String
  ' routine to trim out extra lines and spaces and return message text
  On Error Resume Next
  Dim Counter As Integer
  Dim DeleteLineCount As Integer
  Dim EndLineCount As Integer
  Dim MessageText As String

  MessageText = Trim$(txtMsgBoxText.Text)
  Do While Asc(Mid$(MessageText, Len(MessageText), 1)) < 32
    MessageText = Left$(MessageText, Len(MessageText) - 1)
  Loop
  MessageText = Replace(MessageText, Chr(34), Chr(34) & " & chr(34) & " & Chr(34))
  If ViewCode = True Then MessageText = Replace(MessageText, Chr(13) + Chr(10), Chr(34) & " & vbCrLf & " & Chr(34)) '& vbCrLf)
  MessageText = Replace(MessageText, Chr(38) & " " & Chr(34) & Chr(34) & " " & Chr(38), Chr(38)) '& "" &
  GetMsgBoxText = MessageText
End Function 'GetMsgBoxText(ViewCode As Boolean) As String
'=========================================================================================
Private Function GetMsgBoxTitle() As String
' routine to get the title caption of msgbox
  On Error Resume Next
  GetMsgBoxTitle = Trim$(txtMsgBoxTitle.Text)
End Function 'GetMsgBoxTitle() As String
'=========================================================================================
Private Sub optMsgBox_Click(Index As Integer)
' routine to determine if system or application modal
  On Error Resume Next
  MsgBoxModal = IIf(Index = 0, 0, 4096)
End Sub 'optMsgBox_Click(Index As Integer)
'=========================================================================================
Private Function GetMsgBoxButtonCaptions() As String
' routine to create the string for custom button captions
  On Error Resume Next
  GetMsgBoxButtonCaptions = Trim$(txtButtonCaption(0).Text)
  If Trim$(txtButtonCaption(1).Text) <> "" Then GetMsgBoxButtonCaptions = GetMsgBoxButtonCaptions & "," & txtButtonCaption(1).Text
  If Trim$(txtButtonCaption(2).Text) <> "" Then GetMsgBoxButtonCaptions = GetMsgBoxButtonCaptions & "," & txtButtonCaption(2).Text
End Function 'GetMsgBoxButtonCaptions() As String
'=========================================================================================
Private Sub txtButtonCaption_Change(Index As Integer)
' routine to change the button captions to match with custom captions
  On Error Resume Next
  MsgBoxButtons(Index).Caption = txtButtonCaption(Index).Text
End Sub 'txtButtonCaption_Change(Index As Integer)
'=========================================================================================
Private Function NeedSpecialCode() As Boolean
' routine to determine if special features are being added and to set reqired values
  On Error Resume Next
  Dim Counter As Integer
  Dim SpecialFlag As Boolean

  TimeOut = IIf(Trim$(txtTimeOut.Text) = "", -1, txtTimeOut.Text)
  X = IIf(Trim$(txtPosition(0).Text) = "", -1, txtPosition(0).Text)
  Y = IIf(Trim$(txtPosition(1).Text) = "", -1, txtPosition(1).Text)
  CheckBoxText = IIf(chkSpecial(2).Value = vbChecked, Trim$(txtCheckBoxText.Text), "")
  ' if no text then no checkbox
  If CheckBoxText = "" And chkSpecial(2).Value = vbChecked Then chkSpecial(2).Value = vbUnchecked
  ' set to matching values if one is not filled in
  If X = -1 Then X = Y
  If Y = -1 Then Y = X
  ' no position specified then don't do position
  If X = -1 And Y = -1 And chkSpecial(0).Value = vbChecked Then chkSpecial(0).Value = vbUnchecked
  ' no timeout specified then don't do timeout
  If TimeOut = -1 And chkSpecial(1).Value = vbChecked Then chkSpecial(1).Value = vbUnchecked
  
  ' reset return value and update value if needed
  NeedSpecialCode = False
  For Counter = 0 To 2
    If chkButtonCaption(Counter).Value = vbChecked Then NeedSpecialCode = True
  Next Counter
  For Counter = 0 To 2
    If chkSpecial(Counter).Value = vbChecked Then NeedSpecialCode = True
  Next Counter
End Function 'NeedSpecialCode() As Boolean
'=========================================================================================
Private Sub txtButtonCaption_Validate(Index As Integer, Cancel As Boolean)
' routine to validate that a caption is entered
  On Error Resume Next
  If Trim$(txtButtonCaption(Index).Text) = "" Then
    MsgBox "You need to have a caption for this button", vbInformation, "Need Caption"
    txtButtonCaption(Index).Text = OldCaption
    txtButtonCaption(Index).SelLength = Len(txtButtonCaption(Index).Text)
    Cancel = True
  End If
End Sub 'txtButtonCaption_Validate(Index As Integer, Cancel As Boolean)
'=========================================================================================

