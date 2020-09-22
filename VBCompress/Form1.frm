VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "VBCompress"
   ClientHeight    =   5496
   ClientLeft      =   996
   ClientTop       =   1428
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5496
   ScaleWidth      =   7440
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   5772
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7080
      TabIndex        =   8
      Top             =   2040
      Width           =   372
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7080
      TabIndex        =   7
      Top             =   1320
      Width           =   372
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   5772
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   2652
      Begin VB.OptionButton Option3 
         Caption         =   "expand with API"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2292
      End
      Begin VB.OptionButton Option2 
         Caption         =   "expand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1332
      End
      Begin VB.OptionButton Option1 
         Caption         =   "compress"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1332
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3840
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do it"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1572
   End
   Begin VB.Label Label6 
      Caption         =   "Click on the button ""Do it"""
      Height          =   372
      Left            =   1200
      TabIndex        =   14
      Top             =   4680
      Width           =   5412
   End
   Begin VB.Label Label5 
      Caption         =   "Give in the path you would save the compressed file."
      Height          =   372
      Left            =   1200
      TabIndex        =   13
      Top             =   4080
      Width           =   5892
   End
   Begin VB.Label Label4 
      Caption         =   "Give in the path to file to open."
      Height          =   372
      Left            =   1200
      TabIndex        =   12
      Top             =   3360
      Width           =   5772
   End
   Begin VB.Label Label3 
      Caption         =   "Select the option you want (compress or expand - expand with api"
      Height          =   372
      Left            =   1200
      TabIndex        =   11
      Top             =   2880
      Width           =   5772
   End
   Begin VB.Label Label2 
      Caption         =   "Filename to Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "File to open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As New clsCompress
Private Sub Command1_Click()
Command1.Enabled = False
If Text1.Text = "" Then
MsgBox "Please select the File to compress or expand (File to open)!!"
Command1.Enabled = True
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "Please select the Savefilename (Filename to save)!!"
Command1.Enabled = True
Exit Sub
End If
If Option1.Value = True Then
a.Compress Text1.Text, Text2.Text
End If
If Option2.Value = True Then
a.Expand Text1.Text, Text2.Text
End If
If Option3.Value = True Then
a.ExpandWithAPI Text1.Text, Text2.Text
End If
MsgBox "ready"
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
If Text1.Text = "" Then Exit Sub
If Option1.Value = True Then
Text2.Text = Left(Text1.Text, Len(Text1.Text) - 1) & "_"
End If
If Option2.Value = True Or Option3.Value = True Then
Text2.Text = Left(Text1.Text, Len(Text1.Text) - 1) & "&"
End If

End Sub

Private Sub Command3_Click()
CommonDialog1.FileName = Text2.Text
CommonDialog1.ShowSave
Text2.Text = CommonDialog1.FileName

End Sub

Private Sub Option1_Click()
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub Option2_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Option3_Click()
Text1.Text = ""
Text2.Text = ""
End Sub
