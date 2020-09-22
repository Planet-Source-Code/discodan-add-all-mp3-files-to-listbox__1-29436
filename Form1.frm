VERSION 5.00
Begin VB.Form formMP3List 
   Caption         =   "Add MP3's to List"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox driveMain 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox listFiles 
      Height          =   2400
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton btnGet 
      Caption         =   "Get Files"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox listDirs 
      Height          =   2400
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.DirListBox dirTemp 
      Height          =   540
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.DirListBox dirMain 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label labelResult 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "formMP3List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGet_Click()
'Clear Listboxes
listDirs.Clear
listFiles.Clear
'Add Main Path to List
listDirs.AddItem dirMain.Path
'Add Directories to listDirs
Y = 0
Do Until Y = listDirs.ListCount
    dirTemp.Path = listDirs.List(Y)
    If dirTemp.ListCount > 0 Then
        For X = 0 To dirTemp.ListCount - 1
            listDirs.AddItem dirTemp.List(X)
        Next X
    End If
    Y = Y + 1
Loop
'Go through directory list and add files to listFiles
For X = 0 To listDirs.ListCount - 1
    If listDirs.List(X) Like "*\" Then
    txtFile = Dir(listDirs.List(X) & "*.mp3")
    Else
    txtFile = Dir(listDirs.List(X) & "\*.mp3")
    End If
    If Not txtFile = "" Then
        Do
            tFileName = Path & txtFile
            listFiles.AddItem tFileName
            txtFile = Dir$
        Loop Until txtFile = ""
    End If
Next X
'Display result
labelResult.Caption = listDirs.ListCount & "Folders & " & listFiles.ListCount & " Files."
End Sub
Private Sub driveMain_Change()
'Change drive
dirMain.Path = driveMain.Drive
End Sub
