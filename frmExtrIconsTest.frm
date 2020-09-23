VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExtrIconsTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirTree"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7620
   Icon            =   "frmExtrIconsTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   2295
   End
   Begin MSComctlLib.Toolbar tbrDrives 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   714
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlDrives 
      Left            =   3120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExtrIconsTest.frx":0442
            Key             =   "floppy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExtrIconsTest.frx":0D1E
            Key             =   "hdisk"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExtrIconsTest.frx":15FA
            Key             =   "cd"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1920
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   3780
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic32 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5400
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   4680
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   3840
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3975
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   7011
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "File"
         Text            =   "File"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Path"
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblView 
      Caption         =   "Report"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "View :"
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
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuStyle 
         Caption         =   "&Icon"
         Index           =   0
      End
      Begin VB.Menu mnuStyle 
         Caption         =   "&SmallIcon"
         Index           =   1
      End
      Begin VB.Menu mnuStyle 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu mnuStyle 
         Caption         =   "&Report"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmExtrIconsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O
Dim sPath As String
'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260
Dim FSO As New FileSystemObject
Dim Drv As Drive
Dim Bttn As Button
Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type


'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal x&, ByVal y&, ByVal flags&) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO

Sub FillLvwWithFiles(ByVal path As String)
'-------------------------------------------
'Scan the selected folder for files
'and add then to the listview
'-------------------------------------------
Dim Item As ListItem
Dim s As String, sFile As File

path = CheckPath(path)    'Add '\' to end if not present
s = Dir(path, vbNormal)
Do While s <> ""
  Set Item = lvw.ListItems.Add()
  Item.Key = path & s
  'Item.SmallIcon = "Folder"
  Item.Text = s
  Item.SubItems(1) = path
  Set sFile = FSO.GetFile(path & s)
  Item.SubItems(2) = FormatSize(sFile.Size)
  Item.SubItems(3) = sFile.Type
  Item.SubItems(4) = sFile.DateLastModified
  s = Dir
Loop
LvAutoSize lvw
End Sub
Private Function CheckPath(ByVal path As String) As String
'--------------------------------------------------
'Checks if path ends with "\". If not, add it.
'--------------------------------------------------
If Right(path, 1) <> "\" Then
  CheckPath = path & "\"
Else
  CheckPath = path
End If

End Function

Private Sub ShowFileWithIcons(path As String)
'-------------------------------------------
'Load the files into the listview
'-------------------------------------------
    Caption = "DirTree - " & path
    Initialise
    FillLvwWithFiles path
    GetAllIcons
    ShowIcons
End Sub

Private Sub Initialise()
'-----------------------------------------------
'Initialise the controls
'-----------------------------------------------
On Local Error Resume Next

'Break the link to iml lists
lvw.ListItems.Clear
lvw.Icons = Nothing
lvw.SmallIcons = Nothing

'Clear the image lists
iml32.ListImages.Clear
iml16.ListImages.Clear

End Sub

Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

On Local Error Resume Next
For Each Item In lvw.ListItems
  FileName = Item.SubItems(1) & Item.Text
  GetIcon FileName, Item.Index
Next

End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, pic32.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function
Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With lvw
  '.ListItems.Clear
  .Icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    Item.Icon = Item.Index
    Item.SmallIcon = Item.Index
  Next
End With

End Sub

Private Sub Dir1_Change()
    sPath = Dir1.path
    sPath = CheckPath(sPath)
    ShowFileWithIcons (sPath)
End Sub

Private Sub Form_Load()
    pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
    pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
    pic16.BackColor = vbWhite
    pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
    pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY
    pic32.BackColor = vbWhite
    'cmbStyle.ListIndex = lvw.View
    tbrDrives.ImageList = imlDrives
    ShowFileWithIcons ("C:\")
    For Each Drv In FSO.Drives
        Set Bttn = tbrDrives.Buttons.Add
        Bttn.Caption = LCase(Drv.DriveLetter)
        If Drv.IsReady Then
            Bttn.ToolTipText = Drv.VolumeName
        Else
            Bttn.ToolTipText = LCase(Drv.DriveLetter) & ":\"
        End If
        Select Case Drv.DriveType
            Case 1
                Bttn.Image = "floppy"
            Case 2
                Bttn.Image = "hdisk"
            Case 4
                Bttn.Image = "cd"
        End Select
        Bttn.Tag = LCase$(Drv.DriveLetter) & ":\"
    Next
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim sOption As String
    sOption = Item.SubItems(1) & Item.Text
    Item.Selected = True
    OpenFile (sOption)
End Sub

Private Sub mnuStyle_Click(Index As Integer)
    lvw.View = Index
    Select Case Index
        Case 0
            lblView = "Icon"
        Case 1
            lblView = "SmallIcon"
        Case 2
            lblView = "List"
        Case 3
            lblView = "Report"
    End Select
End Sub

Private Function OpenFile(sFile As String)
Dim lRetval As Long
On Error GoTo ErrRoutine
        lRetval = ShellExecute(Me.hwnd, "open", sFile, vbNullString, vbNullString, 1)
        If lRetval > 32 Then
            Exit Function
        End If
ErrRoutine:
        MsgBox "Can't open " & sFile
        Exit Function
End Function

Private Sub tbrDrives_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrRoutine
        Dir1.path = Button.Tag
        Exit Sub
ErrRoutine:
    MsgBox Err.Description
    Exit Sub
End Sub
