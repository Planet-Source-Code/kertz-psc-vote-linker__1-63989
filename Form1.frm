VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PSC VOTE LINKER"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin Project1.PSCVOTE PSCVOTE1 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   4680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   6735
   End
   Begin VB.CommandButton cmdmail 
      Caption         =   "Mail Me..."
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdweb 
      Caption         =   "My Website..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "The above button will take you to my voting page. Try it in  your project..."
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   5280
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdmail_Click()
    ShellExecute 0, "open", "mailto:streetcoder@gmail.com", "", "", 4
End Sub

Private Sub cmdweb_Click()
    ShellExecute 0, "open", "http:\\www.wackycoder.com", "", "", 4
End Sub

Private Sub Label1_Click()

End Sub
