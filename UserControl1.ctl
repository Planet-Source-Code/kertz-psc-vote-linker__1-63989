VERSION 5.00
Begin VB.UserControl PSCVOTE 
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ScaleHeight     =   615
   ScaleWidth      =   1935
   Begin VB.CommandButton cmdvote 
      Caption         =   "VOTE NOW!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "PSCVOTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Dim sFileName As String 'stores the PSC Read Me File Name
Dim URL As String ' the URL to the voting page
'Event Declarations:
Event Click() 'MappingInfo=cmdvote,cmdvote,-1,Click
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=cmdvote,cmdvote,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=cmdvote,cmdvote,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=cmdvote,cmdvote,-1,MouseUp

Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Function FindFile(path As String, SearchStr As String)
    Dim FileName As String ' Walking filename variable...
    Dim FileCount As Long 'File Counter
    Dim hSearch As Long ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    If Right(path, 1) <> "\" Then path = path & "\"
    hSearch = FindFirstFile(path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            FileName = StripNulls(WFD.cFileName)
           If (FileName <> ".") And (FileName <> "..") Then
               FindFile = FindFile
               FileCount = FileCount + 1
               sFileName = FileName
           End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
End Function

Private Sub GetURL()
    Dim SearchPath As String, FindStr As String, TopicID As String
    'find the readme file
    SearchPath = App.path
    FindStr = "@PSC_ReadMe_*.txt"
    'find file
    FindFile SearchPath, FindStr
    DoEvents
    'Creates URL
    sFileName = Replace(sFileName, "@PSC_ReadMe_", "")
    TopicID = Replace(Replace(sFileName, ".txt", ""), "_1", "")
    URL = "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=" & TopicID & "&lngWId=1"
End Sub

Private Sub cmdvote_Click()
    RaiseEvent Click
    'open the URL...
    ShellExecute 0, "open", URL, "", "", 4
End Sub

Private Sub UserControl_Initialize()
        GetURL ' get the URL of Voting Page...
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdvote,cmdvote,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = cmdvote.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    cmdvote.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdvote,cmdvote,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = cmdvote.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    cmdvote.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdvote,cmdvote,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = cmdvote.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set cmdvote.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdvote,cmdvote,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = cmdvote.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    cmdvote.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdvote,cmdvote,-1,DisabledPicture
Public Property Get DisabledPicture() As Picture
Attribute DisabledPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is disabled, if Style is set to 1."
    Set DisabledPicture = cmdvote.DisabledPicture
End Property

Public Property Set DisabledPicture(ByVal New_DisabledPicture As Picture)
    Set cmdvote.DisabledPicture = New_DisabledPicture
    PropertyChanged "DisabledPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdvote,cmdvote,-1,DownPicture
Public Property Get DownPicture() As Picture
Attribute DownPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is in the down position, if Style is set to 1."
    Set DownPicture = cmdvote.DownPicture
End Property

Public Property Set DownPicture(ByVal New_DownPicture As Picture)
    Set cmdvote.DownPicture = New_DownPicture
    PropertyChanged "DownPicture"
End Property

Private Sub cmdvote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdvote,cmdvote,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = cmdvote.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set cmdvote.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub cmdvote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cmdvote_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdvote,cmdvote,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a CommandButton, OptionButton or CheckBox control, if Style is set to 1."
    Set Picture = cmdvote.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set cmdvote.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdvote,cmdvote,-1,Style
Public Property Get Style() As Integer
Attribute Style.VB_Description = "Returns/sets the appearance of the control, whether standard (standard Windows style) or graphical (with a custom picture)."
    Style = cmdvote.Style
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    cmdvote.Caption = PropBag.ReadProperty("Caption", "VOTE NOW!")
    cmdvote.Enabled = PropBag.ReadProperty("Enabled", True)
    Set cmdvote.Font = PropBag.ReadProperty("Font", Ambient.Font)
    cmdvote.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set DisabledPicture = PropBag.ReadProperty("DisabledPicture", Nothing)
    Set DownPicture = PropBag.ReadProperty("DownPicture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_Resize()
    cmdvote.Height = Height
    cmdvote.Width = Width
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", cmdvote.Caption, "VOTE NOW!")
    Call PropBag.WriteProperty("Enabled", cmdvote.Enabled, True)
    Call PropBag.WriteProperty("Font", cmdvote.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", cmdvote.BackColor, &H8000000F)
    Call PropBag.WriteProperty("DisabledPicture", DisabledPicture, Nothing)
    Call PropBag.WriteProperty("DownPicture", DownPicture, Nothing)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

