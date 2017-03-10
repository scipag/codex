VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2190
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkay 
      Cancel          =   -1  'True
      Caption         =   "&Okay"
      Height          =   372
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Height          =   1332
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.Label lblProjectWebsite 
         AutoSize        =   -1  'True
         Caption         =   "http://www.computec.ch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1200
         MouseIcon       =   "frmAbout.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "Visit web site"
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lblDeveloperName 
         Caption         =   "(c) 2006-2009 by Marc Ruef"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblApplicationName 
         Caption         =   "codEX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   612
         Left            =   240
         MouseIcon       =   "frmAbout.frx":0614
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":091E
         Stretch         =   -1  'True
         ToolTipText     =   "Mainly code defines security."
         Top             =   360
         Width           =   612
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOkay_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & APP_NAME
    Me.lblApplicationName.Caption = APP_NAME
    Me.lblProjectWebsite.Caption = APP_WEBSITE_URL
End Sub

Private Sub imgLogo_Click()
    Call OpenProjectWebsite
End Sub

Private Sub lblProjectWebsite_Click()
    Call OpenProjectWebsite
End Sub
