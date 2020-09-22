VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download File"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "Progress"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6615
      Begin ComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.CommandButton Command 
      Caption         =   "Download"
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin Project1.FileDownloader FileDownloader 
      Left            =   3240
      Top             =   2640
      _ExtentX        =   1799
      _ExtentY        =   1667
   End
   Begin VB.Label Label 
      Caption         =   "File URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   160
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2004 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Private Sub Command_Click()
    If Command.Caption = "Download" Then
        If Trim(Text.Text) <> "" Then
            Label.Enabled = False
            Text.Enabled = False
            
            Command.Caption = "Cancel"
            
            ProgressBar.Value = 0
            
            FileDownloader.DownloadFile Text.Text, App.Path & "\" & GetFileName(Text.Text)
        End If
    Else
        FileDownloader.Cancel
        
        Label.Enabled = True
        Text.Enabled = True
        
        ProgressBar.Value = 0
        
        Command.Caption = "Download"
    End If
End Sub

Private Sub FileDownloader_DowloadComplete()
    MsgBox "Download complete", vbOKOnly + vbInformation, "Success"
    
    Label.Enabled = True
    Text.Enabled = True
    
    Command.Caption = "Download"
    
    ProgressBar.Value = 0
End Sub

Private Sub FileDownloader_DownloadErrors(strError As String)
    MsgBox strError, vbOKOnly + vbCritical, "Error"
    
    Label.Enabled = True
    Text.Enabled = True
    
    Command.Caption = "Download"
    
    ProgressBar.Value = 0
    
    FileDownloader.Cancel
End Sub

Private Sub FileDownloader_DownloadProgress(intPercent As String)
    ProgressBar.Value = intPercent
End Sub

