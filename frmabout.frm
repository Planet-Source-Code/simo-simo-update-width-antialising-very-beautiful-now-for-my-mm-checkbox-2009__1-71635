VERSION 5.00
Begin VB.Form dlgAbout 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "A Propos de ..."
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Fermer"
      Height          =   405
      Left            =   1740
      TabIndex        =   3
      Top             =   2685
      Width           =   1365
   End
   Begin VB.Label lblCopy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "http://mmvb2008.unblog.fr/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   975
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1860
      Width           =   2850
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "[ MAROC ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   1860
      TabIndex        =   4
      Top             =   2265
      Width           =   1125
   End
   Begin VB.Label lblFeatures 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Thanks for using MM_CheckBox 2009 v2.1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   2
      Top             =   495
      Width           =   4395
   End
   Begin VB.Label lblCopy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "M_simohamed@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   975
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1500
      Width           =   2850
   End
   Begin VB.Label lblCopy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copyright © 2008-2009 by M.Simohamed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   1035
      Width           =   4395
   End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_LostFocus()
Unload Me
End Sub
