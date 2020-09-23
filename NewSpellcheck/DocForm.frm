VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Checker"
   ClientHeight    =   8565
   ClientLeft      =   2175
   ClientTop       =   765
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   10860
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   8160
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9060
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   3120
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7410
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "DocForm.frx":0000
      Top             =   720
      Width           =   10725
   End
   Begin VB.Menu Dile 
      Caption         =   "&File"
      Begin VB.Menu Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu Spell 
         Caption         =   "&Spell Check"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'####################################################################################
'# Created by - Pradeep Singh                                                       #
'# Date:- 23/Jan/2001 at 11:41 PM                                                   #
'# Note:-This example only shows wrong word and it doesn't replace that word Because#
'# i don't know how to do that if you have any idea feel free to mail me            #
'#                    pradeepsingh10@hotmail.com                                    #
'####################################################################################

Public Sub check()
Dim DRange As Range

    StatusBar1.Panels(2).Text = "Checking for alternatives please wait  ..."
On Error Resume Next
    Set AppWord = GetObject(, "Word.Application")
    If AppWord Is Nothing Then
        Set AppWord = CreateObject("Word.Application")
        If AppWord Is Nothing Then
            MsgBox "Could not start Word. Application will end"
            End
        End If
    End If
On Error GoTo ErrorHandler
    AppWord.Documents.Add
    Set DRange = AppWord.ActiveDocument.Range
    DRange.InsertAfter Text1.Text
    Set SpellCollectionSpellCollection = DRange.SpellingErrors
    If SpellCollectionSpellCollection.Count > 0 Then
        SuggestionsForm.List1.Clear
        SuggestionsForm.List2.Clear
        For iWord = 1 To SpellCollectionSpellCollection.Count
            SuggestionsForm!List1.AddItem SpellCollectionSpellCollection.Item(iWord)
        Next
    End If
    StatusBar1.Panels(2).Text = "Successfully Done"
    SuggestionsForm.Show
    Exit Sub

ErrorHandler:
    MsgBox "An error occured during the document's spelling" & vbCrLf & Err.Description
End Sub
Private Sub Openn_Click()
Dim strFilename As String
Com.CancelError = True
On Error GoTo errhandler
Com.Filter = _
"Text Files|*.txt*|HTML Files|*.htm*"
Com.ShowOpen
strFilename = Com.FileName
Open strFilename For Input As #1
Text1.Text = Input(LOF(1), 1)
Close #1
Exit Sub
errhandler:
End Sub
Private Sub Command1_Click()
Call check
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub About_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
Dim tlbr_btn As Button
Dim img As ListImage
Dim mypanel As Panel
Set mypanel = StatusBar1.Panels.Add()
mypanel.AutoSize = sbrSpring
mypanel.MinWidth = 1
Set mypanel = StatusBar1.Panels.Add(, , , sbrCaps)
Set mypanel = StatusBar1.Panels.Add(, , , sbrNum)
mypanel.Bevel = sbrInset
mypanel.Alignment = sbrRight
StatusBar1.Panels(1).Text = "Status:" 'StatusBar Caption
StatusBar1.Panels(2).Text = ""
StatusBar1.Panels(1).AutoSize = sbrContents

'Make sure you have installed Graphics in VB6 otherwise you may get error.
On Error GoTo india:

Set img = ImageList1.ListImages.Add(1, "SpellCheck", LoadPicture("C:\Program Files\Microsoft Visual Studio\Common\Graphics\Bitmaps\OffCtlBr\Large\Color\SPELL.BMP"))
Set img = ImageList1.ListImages.Add(2, "Open", LoadPicture("C:\Program Files\Microsoft Visual Studio\Common\Graphics\Bitmaps\OffCtlBr\Large\Color\OPEN.BMP"))
Set img = ImageList1.ListImages.Add(3, "About", LoadPicture("C:\Program Files\Microsoft Visual Studio\Common\Graphics\Icons\Computer\W95MBX04.ICO"))
Toolbar1.ImageList = ImageList1
Set tlbr_btn = Toolbar1.Buttons.Add(1, , "SpellCheck", tbrDefault, "SpellCheck")
tlbr_btn.ToolTipText = "Check Spell" 'ToolTipText
Set tlbr_btn = Toolbar1.Buttons.Add(2, , , tbrSeparator, "Open")
Set tlbr_btn = Toolbar1.Buttons.Add(3, , "Open", tbrDefault, "Open")
tlbr_btn.ToolTipText = "Open File" 'ToolTipText
Set tlbr_btn = Toolbar1.Buttons.Add(4, , , tbrSeparator, "About")
Set tlbr_btn = Toolbar1.Buttons.Add(5, , "About", tlbDefault, "About")
tlbr_btn.ToolTipText = "About" 'ToolTipText
Exit Sub
india:
MsgBox "Make sure you have installed Graphics in VB6 otherwise you get error messages like this" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub Open_Click()
Call Openn_Click
End Sub

Private Sub Spell_Click()
Call check
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
Call check
ElseIf Button.Index = 3 Then
Call Openn_Click
ElseIf Button.Index = 5 Then
frmAbout.Show vbModal, Me
End If
End Sub
