VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "HTML Pad 1.0.0"
   ClientHeight    =   3180
   ClientLeft      =   3225
   ClientTop       =   3540
   ClientWidth     =   5700
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnchor 
      Height          =   375
      Left            =   4320
      Picture         =   "frmMain.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Insert Image!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdInsertHyperlink 
      Height          =   375
      Left            =   3960
      Picture         =   "frmMain.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Insert Hyperlink!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdPreview 
      Height          =   375
      Left            =   5160
      Picture         =   "frmMain.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Preview Document!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdInsertImage 
      Height          =   375
      Left            =   3600
      Picture         =   "frmMain.frx":0748
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Insert Image!"
      Top             =   0
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgImage 
      Left            =   1560
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Jpeg |*.jpg|Gif|*.gif"
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   4800
      Picture         =   "frmMain.frx":084A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Find Word or Phrase!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdAlignRight 
      Height          =   375
      Left            =   3120
      Picture         =   "frmMain.frx":094C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Insert Right Align Tag!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdAlignCenter 
      Height          =   375
      Left            =   2760
      Picture         =   "frmMain.frx":0A4E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Insert Center Tag!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdLeftAlign 
      Height          =   375
      Left            =   2400
      Picture         =   "frmMain.frx":0B50
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Insert Left Align Tag!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdUnderline 
      Height          =   375
      Left            =   1920
      Picture         =   "frmMain.frx":0C52
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Insert Underline Tag!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdItalic 
      Height          =   375
      Left            =   1560
      Picture         =   "frmMain.frx":0D54
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Insert Italic Tag!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdBold 
      Height          =   375
      Left            =   1200
      Picture         =   "frmMain.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Insert Bold Tag!"
      Top             =   0
      Width           =   375
   End
   Begin RichTextLib.RichTextBox txtMain 
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "This is your actual HTML code page. Edit your code here!"
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2143
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      MousePointer    =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0F58
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2760
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open HTML File"
      Filter          =   "HTML | *.html |HTM | *.htm"
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   375
      Left            =   720
      Picture         =   "frmMain.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Open Document!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   360
      Picture         =   "frmMain.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save Document!"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdNew 
      Height          =   375
      Left            =   0
      Picture         =   "frmMain.frx":120A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "New Document!"
      Top             =   0
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save HTML File"
      Filter          =   "HTML | *.html |HTM | *.htm"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Begin VB.Menu mnuNewBlank 
            Caption         =   "&Blank Document"
         End
         Begin VB.Menu mnuNewStanTem 
            Caption         =   "&Standard Template"
         End
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuColours 
         Caption         =   "&Colours"
         Begin VB.Menu mnuBlue 
            Caption         =   "&Blue"
         End
         Begin VB.Menu mnuLime 
            Caption         =   "&Lime"
         End
         Begin VB.Menu mnuNavy 
            Caption         =   "&Navy"
         End
         Begin VB.Menu mnuRed 
            Caption         =   "&Red"
         End
         Begin VB.Menu mnuYellow 
            Caption         =   "&Yellow"
         End
      End
      Begin VB.Menu mnuHeadings 
         Caption         =   "&Headings"
         Begin VB.Menu mnuInsertH1 
            Caption         =   "&H1"
            Shortcut        =   +{F1}
         End
         Begin VB.Menu mnuInsertH2 
            Caption         =   "&H2"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu mnuInsertH3 
            Caption         =   "&H3"
            Shortcut        =   +{F3}
         End
         Begin VB.Menu mnuInsertH4 
            Caption         =   "&H4"
            Shortcut        =   +{F4}
         End
         Begin VB.Menu mnuInsertH5 
            Caption         =   "&H5"
            Shortcut        =   +{F5}
         End
      End
      Begin VB.Menu mnuForm 
         Caption         =   "&Form"
         Begin VB.Menu mnuTextBox 
            Caption         =   "&TextBox"
         End
         Begin VB.Menu mnuScroTextBox 
            Caption         =   "&Scrolling Text Box"
         End
      End
      Begin VB.Menu mnuOther 
         Caption         =   "&Other Stuff"
         Begin VB.Menu mnuAnchor 
            Caption         =   "&Anchor"
         End
         Begin VB.Menu mnuBreak 
            Caption         =   "&Break"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuInsertComment 
            Caption         =   "&Comment"
         End
         Begin VB.Menu mnuInsertImage 
            Caption         =   "&Image"
         End
         Begin VB.Menu mnuHorizontal 
            Caption         =   "&Horizontal Rule"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuHyperlink 
            Caption         =   "&Hyperlink"
         End
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "&Text"
      Begin VB.Menu mnuInsertFont 
         Caption         =   "&FontName"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuCenter 
         Caption         =   "&Center"
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "&Left Align"
      End
      Begin VB.Menu mnuRight 
         Caption         =   "&Right Align"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuItalic 
         Caption         =   "&Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuUnderline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'HTML Pad 1.0.3 (Open Source)
'Free source code at www.planet-source-code.com/vb
'By Mahangu Weerasinghe (vbdude777@email.com)
'Website - http://mahangu.homepage.com
'You may use this code freely in your applications as
'long as you do not sell anything you make with this
'code.

Dim strState As String
Dim strSaved


Private Sub cmdAlignCenter_Click()
Call AlignCenter

End Sub

Private Sub cmdAlignRight_Click()
Call AlignRight
End Sub

Private Sub cmdAnchor_Click()
Call InsertAnchor
End Sub

Private Sub cmdBold_Click()
Call Bold
End Sub

Private Sub cmdFind_Click()
Call Find

End Sub

Private Sub cmdInsertHyperlink_Click()
Call InsertHyperlink
End Sub

Private Sub cmdInsertImage_Click()
Call InsertImage
End Sub

Private Sub cmdItalic_Click()
Call Italic
End Sub

Private Sub cmdLeftAlign_Click()
Call AlignLeft
End Sub

Private Sub cmdNew_Click()
Call NewBlank
End Sub

Private Sub cmdOpen_Click()
Call OpenDoc
End Sub

Private Sub cmdPreview_Click()
Call Preview
End Sub

Private Sub cmdSave_Click()
Call SaveDoc

End Sub

Private Sub cmdUnderline_Click()
Call Underline
End Sub


Private Sub Form_Load()
Me.Caption = App.Title



txtMain.Height = ScaleHeight
txtMain.Width = ScaleWidth



End Sub

Private Sub Form_Resize()
txtMain.Height = ScaleHeight
txtMain.Width = ScaleWidth

End Sub


Private Sub Form_Terminate()
Call ExitNow

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ExitNow
End Sub

Private Sub mnuAbout_Click()
MsgBox "HTML Pad Version 1.0.3(Open Source). Developed by Mahangu Weerasinghe vbdude777@email.com)", vbInformation, "About HTML Pad"



End Sub

Private Sub mnuAnchor_Click()
Call InsertAnchor
End Sub

Private Sub mnuBlue_Click()
Call Colour_Blue
End Sub

Private Sub mnuBold_Click()
Call Bold
End Sub

Private Sub mnuBreak_Click()
Call InsertBreak
End Sub

Private Sub mnuCenter_Click()
Call AlignCenter
End Sub

Private Sub mnuCopy_Click()
Clipboard.SetText txtMain.SelText


End Sub

Private Sub mnuHorizontal_Click()
Call Insert_Horizontal_Rule

End Sub

Private Sub mnuHyperlink_Click()
Call InsertHyperlink
End Sub

Private Sub mnuInsertComment_Click()
Call InsertComment

End Sub

Private Sub mnuInsertFont_Click()
Call InsertFont

End Sub

Private Sub mnuInsertH1_Click()
InsertHeader 1
End Sub

Private Sub mnuInsertH2_Click()
InsertHeader 2
End Sub

Private Sub mnuInsertH3_Click()
InsertHeader 3
End Sub

Private Sub mnuInsertH4_Click()
InsertHeader 4
End Sub

Private Sub mnuInsertH5_Click()
InsertHeader 5
End Sub

Private Sub mnuInsertImage_Click()
Call InsertImage
End Sub

Private Sub mnuItalic_Click()
Call Italic
End Sub

Private Sub mnuLeft_Click()
Call AlignLeft
End Sub

Private Sub mnuLime_Click()
Call Colour_Lime

End Sub

Private Sub mnuNavy_Click()
Call Colour_Navy

End Sub

Private Sub mnuNewBlank_Click()
Call NewBlank
End Sub

Private Sub mnuNewStanTem_Click()
Call NewStandardDoc
End Sub

Private Sub mnuOpen_Click()
Call OpenDoc
End Sub

Private Sub mnuPaste_Click()
txtMain.SelText = Clipboard.GetText

End Sub

Private Sub mnuPreview_Click()
Call Preview
End Sub

Private Sub mnuRed_Click()
Call Colour_Red
End Sub

Private Sub mnuRight_Click()
Call AlignRight
End Sub

Private Sub mnuSave_Click()
Call SaveDoc

End Sub

Private Sub mnuScroTextBox_Click()
Call InsertScrollTextBox

End Sub

Private Sub mnuTextBox_Click()
InsertTextBox
End Sub

Private Sub mnuUnderline_Click()
Call Underline

End Sub

Private Sub mnuYellow_Click()
Call Colour_Yellow

End Sub

Private Sub txtMain_Change()
strState = "Dirty"
strSaved = "False"

End Sub


Function StandardTemplate()
Me.txtMain.Text = ""
Me.txtMain.SelText = Chr(13) + Chr(10) + Me.txtMain.SelText + "<HTML>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
Me.txtMain.SelText = Chr(13) + Chr(10) + Me.txtMain.SelText + "<HEAD>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10) + "<TITLE>" + " Your Title Here </TITLE>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
Me.txtMain.SelText = Chr(13) + Chr(10) + Me.txtMain.SelText + "</HEAD>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)


Me.txtMain.SelText = Chr(13) + Chr(10) + Me.txtMain.SelText + "<BODY>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10) + "<!--- This is the Body of your HTML document. Add all the tags you want here!--->"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
Me.txtMain.SelText = Chr(13) + Chr(10) + Me.txtMain.SelText + "</BODY>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
Me.txtMain.SelText = Chr(13) + Chr(10) + Me.txtMain.SelText + "</HTML>"

End Function

Function OpenDoc()
Dim result
If strSaved = "False" Then result = MsgBox("This file is not saved. Do you want to save it before you open a new file?", vbYesNoCancel, "Hey!")

If result = vbYes Then GoTo 20
If result = vbNo Then GoTo 30

If result = vbCancel Then GoTo 10

30
dlgOpen.FileName = ""
20

dlgOpen.ShowOpen
If dlgOpen.FileName <> "" Then
    Open dlgOpen.FileName For Input As #1
    Me.txtMain.Text = StrConv(InputB(LOF(1), 1), vbUnicode)

10
    dlgOpen.FileName = ""
Close #1

End If

End Function

Function NewBlank()
If txtMain.Text <> "" Then GoTo 1 Else GoTo 10
1


Dim result
If strSaved = "False" Then result = MsgBox("This file is not saved. Do you want to save it before you create a new file?", vbYesNoCancel, "Hey!")

If result = vbYes Then Call mnuSave_Click
If result = vbNo Then txtMain.Text = ""

If result = vbCancel Then GoTo 10
10
End Function

Function NewStandardDoc()
If txtMain.Text = "" Then GoTo 20 Else GoTo 10




10
Dim result
If strSaved = "False" Then result = MsgBox("This file is not saved. Do you want to save it before you create a new file?", vbYesNoCancel, "Hey!")

If result = vbYes Then Call mnuSave_Click
If result = vbNo Then Call StandardTemplate

If result = vbCancel Then GoTo 30

20

Call StandardTemplate
30


End Function
Function SaveDoc()

dlgSave.ShowSave
If dlgSave.FileName <> "" Then

Open dlgSave.FileName For Output As #1
Print #1, txtMain.Text

Close 1

End If
If dlgSave.FileName <> "" Then strSaved = "True"

End Function

Function InsertHeader(Size As String)
Dim S
S = Size

Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10) + "<H" + S + ">" + " Your Heading Here</H" + S + ">"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
End Function

Function InsertBreak()
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10) + "<BR>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)


End Function

Function Insert_Horizontal_Rule()
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10) + "<BR>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)

End Function

Function AlignCenter()
Me.txtMain.SelText = "<CENTER>" + "Your Text Here</CENTER>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
End Function

Function AlignRight()
Me.txtMain.SelText = "<P ALIGN='right'>" + "Your Text Here</P>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
End Function

Function AlignLeft()
Me.txtMain.SelText = "<P ALIGN='left'>" + "Your Text Here</P>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
End Function

Function Bold()
Me.txtMain.SelText = "<B>" + "Your Text Here</B>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)

End Function

Function Italic()
Me.txtMain.SelText = "<I>" + "Your Text Here</I>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)

End Function

Function Underline()
Me.txtMain.SelText = "<U>" + "Your Text Here</U>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)

End Function

Function Colour_Blue()
Me.txtMain.SelText = "<FONT COLOR=#0000FF>" + Me.txtMain.SelText + " Your Text Here</FONT>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)

End Function

Function Colour_Red()
Me.txtMain.SelText = "<FONT COLOR=#FF0000>" + "Your Text Here</FONT>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)

End Function
Function Colour_Navy()
Me.txtMain.SelText = "<FONT COLOR=#000080>" + "Your Text Here</FONT>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)


End Function

Function Colour_Lime()
Me.txtMain.SelText = "<FONT COLOR=#00FF00>" + "Your Text Here</FONT>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)


End Function

Function Colour_Yellow()
Me.txtMain.SelText = "<FONT COLOR=#FFFF00>" + "Your Text Here</FONT>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)


End Function
Function Find()
Dim f
f = InputBox("Type what I should look for.", "What should I find?")
Me.txtMain.Find f

End Function

Function InsertImage()
Dim i
dlgImage.FileName = ""
dlgImage.ShowOpen
i = dlgImage.FileName
If i <> "" Then
    Me.txtMain.SelText = "<img src='file:///" + i + "'>"
    Else: Exit Function
End If
End Function

Function Preview()
Open App.Path + "Htmldude.html" For Output As #1
Print #1, txtMain.Text

Close 1

frmPrev.Show

End Function

Function ExitNow()
If txtMain.Text <> "" Then GoTo 1 Else GoTo 20
1
Dim result
If strSaved = "False" Then result = MsgBox("This file is not saved. Do you want to save it before you quit?", vbYesNoCancel, "Hey!") Else GoTo 20

If result = vbYes Then Call mnuSave_Click
If result = vbNo Then Unload Me
End


If result = vbCancel Then GoTo 10
10

GoTo 30
20

End
30

End Function

Function InsertHyperlink()
h = InputBox("Type the location of URL you want to link to.", "What is the URL you want to link to?")
If h <> "" Then
    Me.txtMain.SelText = "<A HREF =' " + h + "'> Your Link Desciption Here </A>"
    Else: Exit Function
End If
End Function

Function InsertAnchor()
a = InputBox("Type a short name for your anchor.", "What will your Anchor be called?")

Me.txtMain.SelText = "<A NAME =' " + a + "'></A>"

End Function

Function InsertTextBox()
Me.txtMain.SelText = "<input type='text' value = 'Your Text Here'>"
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10)

End Function

Function InsertScrollTextBox()

Me.txtMain.SelText = "<textarea rows='5' cols='20'> Your Text Here</textarea>"
End Function

Function InsertComment()
Me.txtMain.SelText = Chr(13) + Chr(13) + Chr(10) + "<!--- Your Comment Here --->"
End Function

Function InsertFont()
dlgImage.Flags = cdlCFScreenFonts
dlgImage.ShowFont
If dlgImage.FontName <> "" Then
    txtMain.SelText = "<font face='" + dlgImage.FontName + "'> Your Text Here</font>"
    txtMain.SelText = Chr(13) + Chr(13) + Chr(10)
    Else: Exit Function
End If

End Function
