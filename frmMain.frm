VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atwinda Shortcut Arrow"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAble 
      Caption         =   "Disable"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Image imgStatus 
      Height          =   495
      Left            =   1800
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin VB.Image imgArrow 
      Height          =   480
      Left            =   1080
      Picture         =   "frmMain.frx":0CCA
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image imgNone 
      Height          =   480
      Left            =   360
      Picture         =   "frmMain.frx":1994
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":265E
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Andy Stagg - Atwinda Software copyrighted 2000

Private Sub cmdAble_Click() ' this checks to see what the status is of what to di.
If cmdAble.Caption = "Disable" Then ' if it needs to disable it
    Call DeleteValue(HKEY_CLASSES_ROOT, "lnkfile", "IsShortcut") 'delete the value from the key
    cmdAble.Caption = "Enable" 'Then set all your stats to enable
    imgStatus.Picture = imgNone.Picture
    lblStatus.Caption = "Disabled"
    MsgBox "You must restart for settings to take effects.", vbSystemModal, "All Done!" ' tell the user they need to restart
ElseIf cmdAble.Caption = "Enable" Then 'if it needs to enabled then
    Call savestring(HKEY_CLASSES_ROOT, "lnkfile", "IsShortcut", "") 'write the value
    cmdAble.Caption = "Disable" ' and change the enables to disables...
    imgStatus.Picture = imgArrow.Picture
    lblStatus.Caption = "Enabled"
    MsgBox "You must restart for settings to take effects.", vbSystemModal, "All Done!" ' tell the user they need to restart
End If
End Sub

Private Sub Form_Load() ' this check the status of IsShortcut when the program loads!
If CheckKey(HKEY_CLASSES_ROOT, "lnkfile", "IsShortcut") = "Yes" Then ' if it is, then change all the stats to say so
    cmdAble.Caption = "Enable"
    imgStatus.Picture = imgNone.Picture
    lblStatus.Caption = "Disabled"
ElseIf CheckKey(HKEY_CLASSES_ROOT, "lnkfile", "IsShortcut") = "No" Then ' if it isn't then change te stats to say so
    cmdAble.Caption = "Disable"
    imgStatus.Picture = imgArrow.Picture
    lblStatus.Caption = "Enabled"
End If
End Sub


'all the fun is in a mod
