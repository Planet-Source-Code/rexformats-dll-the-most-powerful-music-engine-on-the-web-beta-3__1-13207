VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REX Format - Beta 2"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopyFiles 
      Caption         =   "Copy Files"
      Height          =   375
      Left            =   9240
      TabIndex        =   18
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdFileExists 
      Caption         =   "File Exists"
      Height          =   375
      Left            =   9240
      TabIndex        =   17
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplayFindFiles 
      Caption         =   "FindFiles dialog"
      Height          =   375
      Left            =   9240
      TabIndex        =   14
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register &Association"
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdReadAll 
      Caption         =   "Read all records"
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdWriteMp3ToRex 
      Caption         =   "Write MP3 to Rex"
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdMP3 
      Caption         =   "Read a MP3 file"
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdOpenRexFile 
      Caption         =   "Open Rex File"
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   600
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Result"
         Object.Width           =   9701
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "m3u Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdM3U 
      Caption         =   "&Read a m3u file"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find Data"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdWriteSampleData 
      Caption         =   "&Write Sample Data"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   5400
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCreateRex 
      Caption         =   "&Create Rex File"
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "RexFormats.dll has huge number of useful commands. For example, try :"
      Height          =   855
      Left            =   9240
      TabIndex        =   16
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "rexFormats.DLL can display various dialogs."
      Height          =   855
      Left            =   9240
      TabIndex        =   15
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTestRex.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   7095
   End
   Begin VB.Label Label2 
      Caption         =   "RexFormats.DLL supports m3u files, try it."
      Height          =   495
      Left            =   7320
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "RexFormats Driver 1.0 - Written by Sveinn R. Sigurdss (MrHippo) (C) 2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private r As New SoundFormats ' First, let's declare the RexFormat DLL
Private rexFile As String ' A static variable for the rex file location












' Create a rex PlayList file
Private Sub cmdCreateRex_Click()
        ' Now, let's create the RexFile
        cd.Filter = "Rex File Format (*.rex)|*.rex"
        cd.DefaultExt = "Untitled.rex|*.rex"
        cd.ShowSave
        r.FileName = cd.FileName
        rexFile = cd.FileName
        r.Rex.CreateRexFile
        r.Rex.Initialize
        
        cmdWriteSampleData.Enabled = True
        cmdFind.Enabled = True
End Sub



' If you would like to prevent your playlist to have duplicate
' files, you need to determine if file is already on your disk f.ex.
' The FileExist routine will return a boolean value
Private Sub cmdFileExists_Click()
    MsgBox (r.Other.FileExists(App.Path & "\" & App.EXEName))
End Sub




' Now, find data from the file we've previously created
Private Sub cmdFind_Click()
    Dim i As Long

    ' This clears all possible search criteries
    r.Rex.NewSearch
    ' Simply set desired criteria
    'r.Rex.Find.Album = "Great"
    'r.Rex.Find.Artist = "Var"
    r.Rex.Find.Genre = "Tec"

    ' Find the result
    r.Rex.Search
    
    If r.Rex.RecordCount <> -1 Then
        ListView1.ListItems.Clear
        For i = 0 To lResults
            ' This will read all properties properties from the
            ' selected row into the record object
            r.Rex.ReadRecord (i)
            Call DIsplayRow
        Next i
    End If
End Sub


' Read a Winamp m3u file
Private Sub cmdM3U_Click()
    Dim Count As Long
    Dim x As Long
        
    cd.Filter = "Winamp m3u File (*.m3u)|*.m3u"
    cd.DefaultExt = "*.m3u"
    cd.ShowOpen
    r.FileName = cd.FileName

    r.m3u.Readm3u (r.FileName)
    Count = r.m3u.Count
   
    ' Dim FilePath$, tmpString$, i%, FindComma%
    ListView1.ListItems.Clear
    For x = 0 To r.m3u.Count - 1
        ListView1.ListItems.Add , , r.m3u.FileTitle(x)
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , r.m3u.FileSeconds(x) & " seconds"
    Next x
End Sub



' Read in a mp3 file
Private Sub cmdMP3_Click()
    cd.DialogTitle = "Open Existing mp3 File"
    cd.Filter = "MP3 Music format (*.mp3)|*.mp3"
    cd.DefaultExt = "*.mp3"
    cd.ShowOpen
    If cd.FileName <> "" Then
        r.FileName = cd.FileName
        ' Now let's read in all the tags
        r.mp3.ReadMP3Info
        r.mp3.ReadHeader
        
        Call ReadMP3Tags
        
        cmdWriteSampleData.Enabled = True
        cmdFind.Enabled = True
    End If
End Sub



' Open an existing rex format file
Private Sub cmdOpenRexFile_Click()
    cd.DialogTitle = "Open Existing Rex File"
    cd.Filter = "Rex File Format (*.rex)|*.rex"
    cd.DefaultExt = "*.rex"
    cd.ShowOpen
    
    If cd.FileName <> "" Then
        r.FileName = cd.FileName
        r.Rex.Initialize
        rexFile = r.FileName ' I'm storing the location of the file as static
                             ' this comes in handy when we are writing other
                             ' file formats to rex.
        cmdWriteSampleData.Enabled = True
        cmdFind.Enabled = True
    End If
End Sub



Private Sub cmdReadAll_Click()
    Dim i As Long
    Call r.Rex.LoadAllRecords
    
    With r.Rex
        If .RecordCount > 0 Then
            ListView1.ListItems.Clear
            For i = 0 To .RecordCount - 1
                r.Rex.ReadRecord (i)
                ListView1.ListItems.Add , , "Title := " & r.Rex.Record.Title
            Next i
        End If
    End With
            
End Sub

' REX MusicFormat Beta #3
' Programmed by Sveinn R. Sigurdsson
' e-mail : sveinns@talhf.is
' website : svenni.com
''''''''''''''''''''''''''''''''''''



' Write a sample Data to the file
Private Sub cmdWriteSampleData_Click()

        ' You can also set Global preferences in the file
        ' by using the Preferences attributes
        ' Note :
        '       - EnablePreviews enables the dll to cut
        '         a 5 - 15 second clip from the file
        '         and store it in the rexFile.
        '       - AllowDuplicates True/False (Boolean)
        '         If you do not want the same song to
        '         be written twice to the database
        '         set the property to false
        r.Rex.Preferences.Security_PasswordProtected = False
        r.Rex.Preferences.Security_Password = "anna79"
        r.Rex.Preferences.Security_Username = "svenni76"
        r.Rex.Preferences.EnablePreviews = True
        r.Rex.Preferences.AllowDuplicates = False
        r.Rex.Preferences.Author_Address = "Sunnuflöt 42"
        r.Rex.Preferences.Author_Name = "Sveinn R. Sigurðsson"
        r.Rex.Preferences.Author_Gender = Mr
        r.Rex.Preferences.Author_Country = ICELAND

        ' Now, let's write a single record to the new file
        ' Notice that the filename & path fields must be separated
        r.Rex.Record.Album = "Johny Be Good"
        r.Rex.Record.Title = "Hello"
        r.Rex.Record.FileName = "TestMusic.mp3"
        r.Rex.Record.Path = "C:\"
        r.Rex.Record.ArtistWebsite = "www.svenni.com"
        r.Rex.Record.Genre = Pop
        r.Rex.Record.Rate = r75
        r.Rex.Record.Artist = "Sveinn R. Sigurðsson"
        ' Save the song to the file
        r.Rex.Add
        
        ' Write another song to the file
        ' Notice that now we are writing to another field
        r.Rex.Record.Album = "Great Expectations"
        r.Rex.Record.Title = "Mono - Live in Mono"
        r.Rex.Record.FileName = "mono - live in mono.mp3"
        r.Rex.Record.Path = "C:\music"
        r.Rex.Record.Comments = "No Comments at all"
        r.Rex.Record.Genre = Pop
        r.Rex.Record.Producer = "Sveinn R. Sigurdsson"
        r.Rex.Record.Artist = "Various"
        r.Rex.Record.Arrangement = "Alan Silvestri"
        r.Rex.Record.ArtistWebsite = "www.greatexpectations.com"
        r.Rex.Record.AttributesLastAccessed = "24.12.1999"
        r.Rex.Record.AttributesLastModified = "23.12.1999"
        r.Rex.Record.AttributesReadOnly = "False"
        r.Rex.Record.AudioSiteUrl = "www.amazone.com"
        r.Rex.Record.CDLabel = "Great Expectations - the CD"
        ' Save the song to the file
        r.Rex.Add
        
End Sub



Public Sub DIsplayRow()
    With ListView1
        .ListItems.Add , , "Album := " & r.Rex.Record.Album
        .ListItems.Add , , "Arrangement := " & r.Rex.Record.Arrangement
        .ListItems.Add , , "Artist := " & r.Rex.Record.Artist
        .ListItems.Add , , "ArtistWebsite := " & r.Rex.Record.ArtistWebsite
        .ListItems.Add , , "LastAccessed := " & r.Rex.Record.AttributesLastAccessed
        .ListItems.Add , , "LastModified := " & r.Rex.Record.AttributesLastModified
        .ListItems.Add , , "ReadOnly := " & r.Rex.Record.AttributesReadOnly
        .ListItems.Add , , "AudioSiteURL := " & r.Rex.Record.AudioSiteUrl
        .ListItems.Add , , "ArtistBiography := " & r.Rex.Record.Biography
        .ListItems.Add , , "BuyCDUrl := " & r.Rex.Record.BuyCDUrl
        .ListItems.Add , , "CD Friendly Name := " & r.Rex.Record.CDLabel
        .ListItems.Add , , "Comments := " & r.Rex.Record.Comments
        .ListItems.Add , , "Conductor := " & r.Rex.Record.Conductor
        .ListItems.Add , , "AlbumImage := " & r.Rex.Record.CoverImage
        .ListItems.Add , , "Distribution := " & r.Rex.Record.Distribution
        .ListItems.Add , , "Engineer := " & r.Rex.Record.Engineer
        .ListItems.Add , , "FanSiteUrl := " & r.Rex.Record.FanSite
        .ListItems.Add , , "Filename := " & r.Rex.Record.FileName
        .ListItems.Add , , "FileSize := " & r.Rex.Record.FileSize
        .ListItems.Add , , "Genre := " & r.Rex.Record.Genre
        .ListItems.Add , , "LastPlayed := " & r.Rex.Record.LastPlayedDate
        .ListItems.Add , , "Length := " & r.Rex.Record.Length
        .ListItems.Add , , "Lyrics := " & r.Rex.Record.Lyrics
        .ListItems.Add , , "Mood := " & r.Rex.Record.Mood
        .ListItems.Add , , "Notes provided by := " & r.Rex.Record.MusicNotesProvidedBy
        .ListItems.Add , , "Notes := " & r.Rex.Record.Notes
        .ListItems.Add , , "Original := " & r.Rex.Record.Original
        .ListItems.Add , , "Path := " & r.Rex.Record.Path
        .ListItems.Add , , "Preferences := " & r.Rex.Record.Preference
        .ListItems.Add , , "Preview := " & r.Rex.Record.Preview
        .ListItems.Add , , "Producer := " & r.Rex.Record.Producer
        .ListItems.Add , , "Rate := " & r.Rex.Record.Rate
        .ListItems.Add , , "RecordedAt := " & r.Rex.Record.RecordedAt
        .ListItems.Add , , "Situation := " & r.Rex.Record.Situation
        .ListItems.Add , , "Bitrate := " & r.Rex.Record.SongBitrate
        .ListItems.Add , , "Copyright := " & r.Rex.Record.SongCopyright
        .ListItems.Add , , "CRC := " & r.Rex.Record.SongCRC
        .ListItems.Add , , "Duration := " & r.Rex.Record.SongDuration
        .ListItems.Add , , "Frequency := " & r.Rex.Record.SongFrequency
        .ListItems.Add , , "Layer := " & r.Rex.Record.SongLayer
        .ListItems.Add , , "Mode := " & r.Rex.Record.SongMode
        .ListItems.Add , , "Padding := " & r.Rex.Record.SongPadding
        .ListItems.Add , , "Private := " & r.Rex.Record.SongPrivate
        .ListItems.Add , , "Version := " & r.Rex.Record.SongVersion
        .ListItems.Add , , "Volume := " & r.Rex.Record.SongVolume
        .ListItems.Add , , "Studio := " & r.Rex.Record.Studio
        .ListItems.Add , , "Tempo := " & r.Rex.Record.Tempo
        .ListItems.Add , , "Title := " & r.Rex.Record.Title
        .ListItems.Add , , "TrackNumber := " & r.Rex.Record.TrackNumber
        .ListItems.Add , , "year := " & r.Rex.Record.Year
   End With
End Sub



Private Sub ReadMP3Tags()
    With ListView1
        .ListItems.Clear
        .ListItems.Add , , "Album := " & r.mp3.Album
        .ListItems.Add , , "Artist := " & r.mp3.Artist
        .ListItems.Add , , "Bitrate := " & r.mp3.BitRate
        .ListItems.Add , , "ChannelMode := " & r.mp3.ChannelMode
        .ListItems.Add , , "Comment := " & r.mp3.Comment
        .ListItems.Add , , "Copyright := " & r.mp3.Copyright
        .ListItems.Add , , "CRC Present :=" & r.mp3.CRCPresent
        .ListItems.Add , , "Emphasis := " & r.mp3.Emphasis
        .ListItems.Add , , "File Attributes := " & r.mp3.FileAttributes
        .ListItems.Add , , "FileName :=" & r.mp3.FileName
        .ListItems.Add , , "Framelength := " & r.mp3.FrameLength
        .ListItems.Add , , "FullName := " & r.mp3.FullName
        .ListItems.Add , , "Genre := " & r.mp3.Genre
        .ListItems.Add , , "Layer := " & r.mp3.Layer
        .ListItems.Add , , "Mode Extension := " & r.mp3.ModeExtension
        .ListItems.Add , , "MPEG Version := " & r.mp3.MPEGVersion
        .ListItems.Add , , "Original := " & r.mp3.Original
        .ListItems.Add , , "Padding := " & r.mp3.Padding
        .ListItems.Add , , "Path := " & r.mp3.Path
        .ListItems.Add , , "PlayTime := " & r.mp3.PlayTime
        .ListItems.Add , , "PrivateBits := " & r.mp3.PrivateBit
        .ListItems.Add , , "SampleRate :=" & r.mp3.SampleRate
        .ListItems.Add , , "TagPresent := " & r.mp3.TagPresent
        .ListItems.Add , , "Title := " & r.mp3.Title
        .ListItems.Add , , "TotalFrames := " & r.mp3.TotalFrames
        .ListItems.Add , , "ValidHeader := " & r.mp3.ValidHeader
        .ListItems.Add , , "Year := " & r.mp3.Year
    End With
End Sub



' This demonstrates how to use the RexFormats.DLL to validate information
' and write them to another format using only the DLL
Private Sub cmdWriteMp3ToRex_Click()
    
    ' Let's start by opening a mp3 file
    cd.DialogTitle = "Open Existing MP3 File"
    cd.Filter = "MP3 sound Format (*.mp3)|*.mp3"
    cd.DefaultExt = "*.mp3"
    cd.ShowOpen
    
    If cd.FileName <> "" Then
        r.FileName = cd.FileName
        ' Initalize the mp3 engine
        r.mp3.ReadHeader
        r.mp3.ReadMP3Info
        ' Now I retrieve the stored value of the rexFilename
        r.FileName = rexFile
        ' Initialize the rex engine
        r.Rex.Initialize
        ' Now let's read in the information
    Else
        Exit Sub
    End If
    
    With r.Rex
        ListView1.ListItems.Clear
        
        .Record.ClearFields ' Clear all data fram memory
        .Record.Album = r.mp3.Album
        .Record.Artist = r.mp3.Artist
        .Record.SongBitrate = r.mp3.BitRate
        .Record.Comments = r.mp3.Comment
        .Record.SongCopyright = r.mp3.Copyright
        .Record.SongCRC = r.mp3.CRCPresent
        .Record.SongEmpasis = r.mp3.Emphasis
        .Record.AttributesReadOnly = r.mp3.FileAttributes
        .Record.FileName = r.mp3.FileName
        .Record.Genre = r.mp3.Genre
        .Record.SongLayer = r.mp3.Layer
        .Record.SongMode = r.mp3.ModeExtension
        .Record.Original = r.mp3.Original
        .Record.SongPadding = r.mp3.Padding
        .Record.Path = r.mp3.Path
        .Record.SongDuration = r.mp3.PlayTime
        .Record.SongPrivate = r.mp3.PrivateBit
        .Record.Title = r.mp3.Title
        .Record.Year = r.mp3.Year
        ' Of course you can also add tags to other fields that the mp3 file
        ' format does not store. F.ex. the user enters the url for the
        ' artists website, then you could add this also by simply writing
        .Record.ArtistWebsite = "http:\\www.johnlennon.com"
        .Add ' Save this data to rex
    End With
End Sub



' Let's create an association to the rex file format
' When a user clicks on a rex file from windows explorer
' Windows will launch the associated program.. that it,
' your program.
'
' Note : You can test this function by opening the
' Windows explorer, select a detailed view and click
' on a rex file that you have created.
Private Sub cmdRegister_Click()
    With r.Registry
        .ApplicationExecutable = App.EXEName
        .ApplicationName = "DemoApp"
        .ApplicationPath = App.Path
        .DocumentName = "Document"
        .DocumentDescription = "Rex Music Library"
        .CreateAssociation
    End With
End Sub



' Display the FindFiles Windows dialog
' You can open the window with a preselected
' path, let's open the window, displaying
' "C:\" in the "Look In" field.
Private Sub cmdDisplayFindFiles_Click()
    r.Dialogs.DisplayFindFiles ("C:")
End Sub



' You may want to enable the user to copy files
' from one location to another, f.ex. a mp3 file
' from a cd to a local drive.
' RexFormats.DLL helps you to do that, simply :
Private Sub cmdCopyFiles_Click()
   Dim sSourcePath  As String   ' Where are the file(s) to be copied.
   Dim sDestination As String   ' To Where should we copy the files
   Dim sFiles       As String   ' You can use the standard DOS
                                ' attributes when selecting files.
                                ' F.ex. the command *.txt means,
                                ' that you will copy all files
                                ' with the *.txt extension.
   sSourcePath = "c:\winnt\"    ' You may need to change the value
   sDestination = "c:\temptest\"
   sFiles = "*.txt"
   
   ' Now, let's copy the files.
   ' Note that the CopyFiles function returns a long value
   ' that contains the number of files, copied
   MsgBox ("Files Copied : " & r.Other.CopyFiles(sSourcePath, _
                                                sDestination, _
                                                sFiles))
End Sub







