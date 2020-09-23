VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Call and Read from a file..."
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5775
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saving and Opening"
      BeginProperty Font 
         Name            =   "Myriad Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.ListBox lstAlignment 
         Height          =   645
         ItemData        =   "frmMain.frx":0742
         Left            =   2400
         List            =   "frmMain.frx":074F
         TabIndex        =   12
         ToolTipText     =   "Can't Select Multiple"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.ListBox lstFont 
         Height          =   645
         ItemData        =   "frmMain.frx":077A
         Left            =   120
         List            =   "frmMain.frx":0787
         MultiSelect     =   2  'Extended
         TabIndex        =   10
         ToolTipText     =   "Hold CTRL To Select More Than 1"
         Top             =   2880
         Width           =   2175
      End
      Begin RichTextLib.RichTextBox rtbText 
         Height          =   2295
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   4048
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":07A4
      End
      Begin VB.TextBox txtDestroy 
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Text            =   "C:\Windows\Desktop\Test.txt"
         Top             =   4440
         Width           =   2655
      End
      Begin VB.CommandButton cmdDestroy 
         Caption         =   "Destroy this file..."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   4440
         Width           =   1815
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   5160
         Width           =   4575
      End
      Begin VB.CommandButton cmdOpen2 
         Caption         =   "Open from saved path..."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   4800
         Width           =   4575
      End
      Begin VB.TextBox txtOpen 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Text            =   "C:\Windows\Win.ini"
         Top             =   4080
         Width           =   2655
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open from this path..."
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox txtSave 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Text            =   "C:\Windows\Desktop\Test.txt"
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save to this path..."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Edit Alignment"
         BeginProperty Font 
            Name            =   "Myriad Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Edit Font"
         BeginProperty Font 
            Name            =   "Myriad Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   2640
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This code is not mine.. I found it on PSC.. dont know who made it
'It basically deletes all text in a file, then deletes the file
Private Sub cmdDestroy_Click()
    'This routine makes sure you want to delete the chosen file
    Dim YesNo As Integer
    YesNo = MsgBox("This will delete the selected file. &_ Are you sure you wish to continue?", vbYesNo + vbCritical, "Warning")
    If YesNo = vbYes Then
    'If you wanna destroy it, Then:
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, offset As Long
    Dim sFileName
    Const BLOCKSIZE = 4096
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
    hFileHandle = FreeFile
    'Identifies the file to be destroyed
    sFileName = txtDestroy.Text
    'opens the file to ne destroyed
    Open sFileName For Binary As hFileHandle
    Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1
    For iLoop = 1 To Blocks
        offset = Seek(hFileHandle)
        Put hFileHandle, , Block1
        Put hFileHandle, offset, Block2
    Next iLoop
    Close hFileHandle
    'destroys the file
    Kill sFileName
    MsgBox "File has been deleted: " & txtDestroy.Text, vbInformation, "Destroyed"
    StatusBar1.Panels(1).Text = "File deleted: " & txtDestroy.Text
    End If
    If YesNo = vbNo Then
    'If you dont wanna destroy it, Then:
    MsgBox "File will not be destroyed!", vbInformation, "Information"
    Exit Sub
    End If
End Sub

Private Sub cmdExit_Click()
'close the form and stop running
Unload Me
End
End Sub

Private Sub cmdOpen_Click()
'loads a completey new and specifically chosen file
'opens it directly to rtbText
rtbText.LoadFile txtOpen.Text
'confirming your open
MsgBox "File was opened from " & txtOpen.Text, vbInformation, "Opened"
StatusBar1.Panels(1).Text = "File opened from: " & txtOpen.Text
End Sub

Private Sub cmdOpen2_Click()
'opens the file which you saved
'Assumes: you didnt alter your save path
rtbText.LoadFile txtSave.Text
MsgBox "The currently saved file was just opened.", vbInformation, "Opened"
StatusBar1.Panels(1).Text = "Saved file has been opened."
End Sub

Private Sub cmdSave_Click()
'i think this is the simplest way to save simple text
Open txtSave.Text For Append As 1
'prints the text on the first line
'Warning: this will not overwrite old text..
Print #1, rtbText.Text
'close the opened file
Close 1
MsgBox "File was saved to " & txtSave.Text, vbInformation, "Saved"
StatusBar1.Panels(1).Text = "File saved to: " & txtSave.Text
End Sub



Private Sub lstAlignment_Click()
'aligns the text in rtbText
Select Case lstAlignment
Case "Align Left"
rtbText.SelAlignment = rtfLeft
Case "Align Center"
rtbText.SelAlignment = rtfCenter
Case "Align Right"
rtbText.SelAlignment = rtfRight
End Select
End Sub

Private Sub lstFont_Click()
'this is messed up.. you gotta click it again to unbold or whatever
'sorry about that
Select Case lstFont
Case "Bold"
rtbText.SelBold = Not rtbText.SelBold
Case "Underline"
rtbText.SelUnderline = Not rtbText.SelUnderline
Case "Italic"
rtbText.SelItalic = Not rtbText.SelItalic
End Select
End Sub
