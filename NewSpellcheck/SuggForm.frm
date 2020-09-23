VERSION 5.00
Begin VB.Form SuggestionsForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spell Checker"
   ClientHeight    =   5400
   ClientLeft      =   5070
   ClientTop       =   2325
   ClientWidth     =   4860
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1380
      TabIndex        =   2
      Top             =   4920
      Width           =   1845
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      ItemData        =   "SuggForm.frx":0000
      Left            =   180
      List            =   "SuggForm.frx":0002
      TabIndex        =   1
      Top             =   2700
      Width           =   4365
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "SuggForm.frx":0004
      Left            =   180
      List            =   "SuggForm.frx":0006
      TabIndex        =   0
      ToolTipText     =   "Select for suggestions"
      Top             =   420
      Width           =   4365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Words:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   4635
   End
   Begin VB.Frame Frame2 
      Caption         =   "Suggestions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   60
      TabIndex        =   4
      Top             =   2400
      Width           =   4695
   End
End
Attribute VB_Name = "SuggestionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
AppWord.Application.Quit False 'close MS-Word
Set AppWord = Nothing 'kill object
Unload SuggestionsForm
End Sub


Private Sub List1_Click()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Set CorrectionsCollectionCollection = _
        AppWord.GetSpellingSuggestions(SpellCollectionSpellCollection.Item _
        (List1.ListIndex + 1))
    List2.Clear
    For iSuggWord = 1 To CorrectionsCollectionCollection.Count
        List2.AddItem CorrectionsCollectionCollection.Item(iSuggWord)
    Next
    Screen.MousePointer = vbDefault

End Sub
