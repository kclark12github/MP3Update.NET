Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Security.Policy
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports TagLib
Imports TagLib.Jpeg
Public Class clsMP3Update
    Public Sub New()
        MyBase.New()
        Dim CurrentAssembly As [Assembly] = System.Reflection.Assembly.GetEntryAssembly()
        mAppName = CurrentAssembly.GetName.Name.ToString()
        mEventLog = New EventLog("Application")
        mEventLog.EnableRaisingEvents = True
        'Create Source manually using PowerShell: New-EventLog -LogName Application -Source "MP3Update.NET"
        If Not EventLog.SourceExists(mAppName) Then EventLog.CreateEventSource(mAppName, "Application")
        mEventLog.Source = mAppName
    End Sub
    Public Event ListMessage(ByVal Message As String)
    Public Event ListTitle(ByVal Title As String)
    Private UIEnabled As Boolean = True
    Private DBEnabled As Boolean = True
    Private mEventLog As EventLog
    Private mAppName As String
    Private mCancel As Boolean = False
    Private mErrorCount As Long = 0
    Private mProgressCount As Long = 0
    Private mTotalFiles As Long = 0
    Private mSQLConnection As SqlClient.SqlConnection
#Region "Properties"
    Public ReadOnly Property AppName() As String
        Get
            Return mAppName
        End Get
    End Property
    Public Property Cancel() As Boolean
        Get
            Return mCancel
        End Get
        Set(ByVal Value As Boolean)
            mCancel = Value
        End Set
    End Property
    Public ReadOnly Property ProgressCount() As Long
        Get
            Return mProgressCount
        End Get
    End Property
    Public ReadOnly Property TotalFiles() As Long
        Get
            Return mTotalFiles
        End Get
    End Property
#End Region
#Region "Methods"
    Private Sub AddSong(ByVal Song As myTag, ByVal GeneralizedPath As String)
        Try
            Dim SQLSource As String = $"Insert Into Songs " &
                    "(Genre,Artist,Performer,ArtistID,Year,Album,AlbumID,Disc,Track,Title,Duration,Composer,Comment,Path,Lyrics,Publisher) " &
                    "Values (@Genre,@Artist,@Performer,@ArtistID,@Year,@Album,@AlbumID,@Disc,@Track,@Title,@Duration,@Composer,@Comment,@Path,@Lyrics,@Publisher)"
            Dim da As New SqlDataAdapter
            da.InsertCommand = New SqlCommand : da.InsertCommand.Connection = mSQLConnection
            With da.InsertCommand
                .CommandText = SQLSource
                .CommandType = CommandType.Text
                'Now for each column we intend to update, create and initialize its parameter...
                .Parameters.Clear()
                .Parameters.AddWithValue("@Genre", Song.Genre)
                .Parameters.AddWithValue("@Artist", Song.Artist)
                .Parameters.AddWithValue("@Performer", Song.Performer)
                .Parameters.AddWithValue("@ArtistID", GetArtist(Song))
                .Parameters.AddWithValue("@Year", Song.Year)
                .Parameters.AddWithValue("@Album", Song.Album)
                .Parameters.AddWithValue("@AlbumID", GetAlbum(Song))
                .Parameters.AddWithValue("@Disc", Song.Disc)
                .Parameters.AddWithValue("@Track", Song.Track)
                .Parameters.AddWithValue("@Title", Song.Title)
                .Parameters.AddWithValue("@Duration", Song.Duration)
                .Parameters.AddWithValue("@Composer", Song.Composer)
                .Parameters.AddWithValue("@Comment", Song.Comment)
                .Parameters.AddWithValue("@Path", GeneralizedPath)
                .Parameters.AddWithValue("@Lyrics", Song.Lyrics)
                .Parameters.AddWithValue("@Publisher", Song.Publisher)
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            Song.Dump()
            Throw
        End Try
    End Sub
    Public Sub CheckFiles(ByVal BaseDir As DirectoryInfo)
        Try
            Dim diList As DirectoryInfo() = BaseDir.GetDirectories()
            For Each di As DirectoryInfo In diList
                Select Case di.Name
                    Case "System Volume Information"
                    Case "Temporary Internet Files"
                    Case Else
                        Try
                            If UIEnabled Then RaiseEvent ListMessage(di.FullName)
                            Application.DoEvents() : If mCancel Then Exit Try
                        Catch ex As Exception
                            Dim Message As String = String.Format("Error processing {0}; ", di.Name) & vbCrLf
                            Message &= vbCrLf
                            Message &= String.Format("Exception: {0}", ex.ToString)
                            mEventLog.WriteEntry(Message)
                        End Try
                        Application.DoEvents() : If mCancel Then Exit Try
                        CheckFiles(di)
                End Select
            Next
            Dim fiList As FileInfo() = BaseDir.GetFiles()
            For Each fi As FileInfo In fiList
                Try
                    mProgressCount += 1
                    If UIEnabled Then RaiseEvent ListMessage(fi.FullName)
                    Application.DoEvents() : If mCancel Then Exit Try

                    Dim objMP3Info As New MP3.MP3Info(fi.FullName)
                    If (objMP3Info.ID3v1Tag.TagAvailable) Then
                        With objMP3Info.ID3v1Tag
                            Dim PreferredFileName As String = String.Format("{0:00} - {1} - {2}.mp3", .Track, .Artist, GetFileNameFromTitle(.Title))
                            If fi.Name <> PreferredFileName Then
                                If MessageBox.Show(String.Format("""{0}"" does not match preferred name ""{1}""", fi.Name, PreferredFileName), "Rename?", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                                    fi.MoveTo(String.Format("{0}\{1}", fi.DirectoryName, PreferredFileName))
                                End If
                            End If
                        End With
                    End If
                Catch ex As Exception
                    Dim Message As String = String.Format("Error processing {0}; ", fi.Name) & vbCrLf
                    Message &= vbCrLf
                    Message &= String.Format("Exception: {0}", ex.ToString)
                    mEventLog.WriteEntry(Message)
                End Try
                Application.DoEvents() : If mCancel Then Exit Try
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString, mAppName & ".CheckFiles", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
        End Try
    End Sub
    Public Sub CleanupFiles(ByVal BaseDir As DirectoryInfo)
        Try
            Dim diList As DirectoryInfo() = BaseDir.GetDirectories()
            For Each di As DirectoryInfo In diList
                Select Case di.Name
                    Case "System Volume Information"
                    Case "Temporary Internet Files"
                    Case "Christmas", "Classical", "Comedy", "Country", "Jazz", "New Age", "Rock", "Soundtracks"
                        If UIEnabled Then RaiseEvent ListMessage(di.FullName)
                        Application.DoEvents() : If mCancel Then Exit Try
                        ProcessFiles(di)
                End Select
            Next
            Dim allowedExtensions As String() = {".mp3", ".flac"}
            Dim audioFiles As Generic.IEnumerable(Of FileInfo) = BaseDir.EnumerateFiles("*.*", SearchOption.AllDirectories) _
            .Where(Function(f) allowedExtensions.Contains(f.Extension.ToLower()))
            For Each fi As FileInfo In audioFiles
                Try
                    If UIEnabled Then RaiseEvent ListMessage(fi.FullName)
                    Application.DoEvents() : If mCancel Then Exit Try

                    If Not fi.Exists Then
                        If DBEnabled Then DoCommand($"Delete From Songs Where Path='{fi.FullName.Replace("'", "''").Replace("[", "[[]")}'")
                    End If
                    mProgressCount += 1
                Catch ex As Exception
                    Dim Message As String = $"Error processing {fi.Name}; {vbCrLf}{vbCrLf}Exception: {ex.ToString}"
                    mEventLog.WriteEntry(Message)
                End Try
                Application.DoEvents() : If mCancel Then Exit Try
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString, mAppName & ".CleanupFiles", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
        End Try
    End Sub
    Public Sub CountFiles(ByVal BaseDir As DirectoryInfo, ByRef cntFolders As Long, ByRef cntFiles As Long)
        Try
            Dim diList As DirectoryInfo() = BaseDir.GetDirectories()
            cntFolders = 0
            cntFiles = 0
            For Each di As DirectoryInfo In diList
                Select Case di.Name
                    Case "System Volume Information"
                    Case "Temporary Internet Files"
                    Case "Christmas", "Classical", "Comedy", "Country", "Jazz", "New Age", "Rock", "Soundtracks"
                        cntFolders += 1
                End Select
            Next
            Dim allowedExtensions As String() = {".mp3", ".flac"}
            Dim audioFiles As Generic.IEnumerable(Of FileInfo) = BaseDir.EnumerateFiles("*.*", SearchOption.AllDirectories) _
            .Where(Function(f) allowedExtensions.Contains(f.Extension.ToLower()))
            For Each fi As FileInfo In audioFiles
                cntFiles += 1
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString, mAppName & ".CountFiles", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
        End Try
    End Sub
    Public Sub DoCommand(ByVal SQLSource As String)
        Dim SQLCommand As New SqlClient.SqlCommand
        Try
            With SQLCommand
                .CommandText = SQLSource
                .CommandType = CommandType.Text
                .Connection = mSQLConnection
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
            Throw ex
        Finally
            SQLCommand = Nothing
        End Try
    End Sub
    Public Function DoScalar(ByVal SQLSource As String) As Object
        Dim SQLCommand As New SqlClient.SqlCommand
        Try
            With SQLCommand
                .CommandText = SQLSource
                .CommandType = CommandType.Text
                .Connection = mSQLConnection
                Return .ExecuteScalar
            End With
        Catch ex As Exception
            Throw ex
        Finally
            SQLCommand = Nothing
        End Try
    End Function
    Private Function ElapsedTime(ByRef StartTime As Date, ByRef FinishTime As Date, Optional ByRef tFormat As Short = 0) As String
        Dim MM, HH, SS As Integer
        Dim strTime As String

        strTime = vbNullString
        SS = DateDiff(Microsoft.VisualBasic.DateInterval.Second, StartTime, FinishTime)
        HH = SS \ 3600
        SS = SS - (HH * 3600)
        MM = SS \ 60
        SS = SS - (MM * 60)
        If HH > 0 Then strTime = HH & " Hours, "
        If MM > 0 Then strTime = strTime & MM & " Minutes, "
        strTime = strTime & SS & " Seconds"

        Select Case tFormat
            Case 0
                ElapsedTime = $"{HH:00}:{MM:00}:{SS:00}"
            Case Else
                ElapsedTime = strTime
        End Select
    End Function
    Private Function GetAlphaSort(ByVal Artist As String) As String
        Select Case Artist
            Case ".38 Special",
                "Adam Ant",
                "Alkaline Trio",
                "American Hi-Fi",
                "Bachman-Turner Overdrive",
                "Bachman–Turner Overdrive",
                "Backstreet Boys",
                "Bad Company",
                "Bad English",
                "Band Aid",
                "Bee Gees",
                "Black Sabbath",
                "Blind Faith",
                "Blue Swede",
                "Blues Image",
                "Buffalo Springfield",
                "Cheap Trick",
                "Coldstream Guards",
                "Collective Soul",
                "Counting Crows",
                "Cowboy Junkies",
                "Crowded House",
                "Cutting Crew",
                "Deep Purple",
                "Def Leppard",
                "Dire Straits",
                "Doctor Who",
                "Duran Duran",
                "Elephant's Memory",
                "Fleetwood Mac",
                "Gin Blossoms",
                "Glass Tiger",
                "Golden Earring",
                "Human League",
                "Iron Butterfly",
                "Jefferson Airplane",
                "Jefferson Starship",
                "Jethro Tull",
                "King Crimson",
                "King Harvest",
                "Lady Gaga",
                "Laidback Luke",
                "Led Zeppelin",
                "Lifescapes Music",
                "Little Feat",
                "Los Muppets",
                "Lynyrd Skynyrd",
                "Mannheim Steamroller",
                "Maroon 5",
                "Matchbox Twenty",
                "Meat Loaf",
                "Midnight Oil",
                "Miss Piggy",
                "Molly Hatchet",
                "Monty Python",
                "Naked Eyes",
                "New Edition",
                "Night Ranger",
                "OK Go",
                "Pearl Jam",
                "Pink Floyd",
                "Quadriga Consort",
                "Rascal Flats",
                "Red Rider",
                "REO Speedwagon",
                "Royal Philharmonic",
                "Rusted Root",
                "Simple Minds",
                "Soft Cell",
                "Spandau Ballet",
                "Status Quo",
                "Sugar Beats",
                "T.Rex",
                "Talking Heads",
                "Thin Lizzy",
                "'Til Tuesday",
                "Trans-Siberian Orchestra",
                "Traveling Wilburys",
                "TV Themes",
                "Van Halen",
                "Various Artists",
                "Village People",
                "Wang Chung",
                "Whirling Dervishes"
                Artist = Artist.ToUpper
            Case Else
                Artist = Artist.ToUpper
                If Artist.StartsWith("THE ") Then
                    Artist = Artist.Substring("THE ".Length) & ", THE"
                ElseIf Artist.StartsWith("A ") Then
                    Artist = Artist.Substring("A ".Length) & ", A"
                ElseIf Artist.StartsWith("AN ") Then
                    Artist = Artist.Substring("AN ".Length) & ", AN"
                Else
                    Dim words() As String = Artist.Split(" ")
                    'Let's only deal with two-word patterns for now...
                    If words.Length = 2 Then Artist = $"{words(1)}, {words(0)}"
                End If
        End Select
        Return Artist
    End Function
    Private Function GetFileNameFromTitle(ByVal Title As String) As String
        If InStr(Title, "/") > 0 Then Title = Replace(Title, "/", "-")
        If InStr(Title, "\") > 0 Then Title = Replace(Title, "\", "-")
        If InStr(Title, ":") > 0 Then Title = Replace(Title, ":", "-")
        If InStr(Title, "*") > 0 Then Title = Replace(Title, "*", "-")
        If InStr(Title, """") > 0 Then Title = Replace(Title, """", "'")
        If InStr(Title, "?") > 0 Then Title = Replace(Title, "?", "_")
        If InStr(Title, "<") > 0 Then Title = Replace(Title, "<", "_")
        If InStr(Title, ">") > 0 Then Title = Replace(Title, ">", "_")
        If InStr(Title, "|") > 0 Then Title = Replace(Title, "|", "-")
        Return Title
    End Function
    Private Function GetAlbum(ByVal Song As myTag) As Integer
        Try
            GetAlbum = DoScalar($"Select ID From Albums Where Artist='{Song.Artist.Replace("'", "''").Replace("[", "[[]")}' And Title='{Song.Album.Replace("'", "''").Replace("[", "[[]")}' And Year={Song.Year}")
            If IsNothing(GetAlbum) OrElse IsDBNull(GetAlbum) OrElse GetAlbum = 0 Then
                Dim SQLSource As String = $"Insert Into Albums (Artist,Title,Year) Values (@Artist,@Title,@Year);Select SCOPE_IDENTITY();"
                Dim da As New SqlDataAdapter
                da.InsertCommand = New SqlCommand : da.InsertCommand.Connection = mSQLConnection
                With da.InsertCommand
                    .CommandText = SQLSource
                    .CommandType = CommandType.Text
                    'Now for each column we intend to update, create and initialize its parameter...
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@Artist", Song.Artist)
                    .Parameters.AddWithValue("@Title", Song.Album)
                    .Parameters.AddWithValue("@Year", Song.Year)
                    Return .ExecuteScalar
                End With
            End If
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            Song.Dump()
            Throw
        End Try
    End Function
    Private Function GetArtist(ByVal Song As myTag) As Integer
        Try
            GetArtist = DoScalar($"Select ID From Artists Where Name='{Song.Artist.Replace("'", "''").Replace("[", "[[]")}'")
            If IsNothing(GetArtist) OrElse IsDBNull(GetArtist) OrElse GetArtist = 0 Then
                Dim SQLSource As String = $"Insert Into Artists (Name,AlphaSort) Values (@Name,@AlphaSort);Select SCOPE_IDENTITY();"
                Dim da As New SqlDataAdapter
                da.InsertCommand = New SqlCommand : da.InsertCommand.Connection = mSQLConnection
                With da.InsertCommand
                    .CommandText = SQLSource
                    .CommandType = CommandType.Text
                    'Now for each column we intend to update, create and initialize its parameter...
                    .Parameters.Clear()
                    Dim Artist As String = Song.Artist
                    .Parameters.AddWithValue("@Name", Artist)
                    .Parameters.AddWithValue("@AlphaSort", GetAlphaSort(Artist))
                    Return .ExecuteScalar
                End With
            End If
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            Song.Dump()
            Throw
        End Try
    End Function
    Public Sub ProcessFiles(ByVal BaseDir As DirectoryInfo)
        Try
            Dim diList As DirectoryInfo() = BaseDir.GetDirectories()
            For Each di As DirectoryInfo In diList
                Select Case di.Name
                    Case "System Volume Information"
                    Case "Temporary Internet Files"
                    Case "Christmas", "Classical", "Comedy", "Country", "Jazz", "New Age", "Radio Shows", "Rock", "Soundtracks"
                        If UIEnabled Then RaiseEvent ListMessage(di.FullName)
                        Application.DoEvents() : If mCancel Then Exit Try
                        ProcessFiles(di)
                End Select
            Next
            Dim allowedExtensions As String() = {".mp3", ".flac"}
            Dim audioFiles As Generic.IEnumerable(Of FileInfo) = BaseDir.EnumerateFiles("*.*", SearchOption.AllDirectories) _
            .Where(Function(f) allowedExtensions.Contains(f.Extension.ToLower()))
            For Each fi As FileInfo In audioFiles
                Try
                    Dim Path As String = fi.FullName
                    If UIEnabled Then RaiseEvent ListMessage(Path)
                    Application.DoEvents() : If mCancel Then Exit Try

                    Dim Song As myTag
                    Try
                        Dim taglibFile As TagLib.File = TagLib.File.Create(Path)
                        Song = New myTag(taglibFile, Path)
                    Catch ex As Exception
                        'Try to fail over to the old architecture...
                        Dim mp3Info As New MP3.MP3Info(Path)
                        If (mp3Info.ID3v1Tag.TagAvailable) Then
                            Song = New myTag(mp3Info.ID3v1Tag, Path)
                        Else
                            Throw
                        End If
                    End Try

                    'Debug.WriteLine($"Processing Song: {Path}")
                    Dim ID As Integer = DoScalar($"Select ID From Songs Where Path Like '%{Path.Replace("'", "''").Replace("[", "[[]")}'")
                        If DBEnabled Then
                            If IsNothing(ID) OrElse IsDBNull(ID) OrElse ID = 0 Then
                                AddSong(Song, Path)
                            Else
                                UpdateSong(Song, ID)
                            End If
                        End If
                        mProgressCount += 1
                    Catch ex As Exception
                        mErrorCount += 1
                    Dim Message As String = $"Error processing {fi.FullName}; {vbCrLf}{vbCrLf}Exception: {ex.ToString}"
                    mEventLog.WriteEntry(Message)
                    Debug.WriteLine(Message)
                End Try
                Application.DoEvents() : If mCancel Then Exit Try
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString, mAppName & ".ProcessFiles", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
        End Try
    End Sub
    'Private Sub SaveHistory(ByVal dataRow As DataRow, ByVal dataMember As String)
    '    Dim HistoryCommand As SqlCommand = Nothing
    '    Dim oldVersion As DataRowVersion
    '    Dim newVersion As DataRowVersion
    '    Try
    '        'We wouldn't have gotten here if something wasn't changed... We just need to rediscover the actual change...
    '        'If there's no Original version, the row must be new, so consider it a change...
    '        If Not dataRow.HasVersion(DataRowVersion.Original) Then
    '            'Record Added
    '        End If
    '        oldVersion = DataRowVersion.Original

    '        newVersion = DataRowVersion.Default
    '        'According to the documentation, Proposed should only exist when BeginEdit/EndEdit/CancelEdit are being used
    '        'but I'm not seeing this as being the case. So until I get it straight, use it if it exists...
    '        'It may be that DataRowVersion.Proposed is only intended on being used in RowChanging event handlers...
    '        If dataRow.HasVersion(DataRowVersion.Proposed) Then newVersion = DataRowVersion.Proposed

    '        Dim SQL As String = "Insert Into History([Who],[DateChanged],[TableName],[RecordID],[Column],[OriginalValue],[Value]) Values (@Who,CURRENT_TIMESTAMP,@TableName,@RecordID,@Column,@OriginalValue,@Value)"
    '        Dim OriginalValue As String = String.Empty
    '        Dim Value As String = String.Empty

    '        Select Case mMode
    '            Case ActionModeEnum.modeAdd, ActionModeEnum.modeCopy
    '                OriginalValue = "Record Added"
    '                Value = GetDataRowValue(dataRow, dataMember, newVersion)
    '            Case ActionModeEnum.modeDelete
    '                OriginalValue = GetDataRowValue(dataRow, dataMember, oldVersion)
    '                Value = "Record Deleted"
    '            Case Else
    '                OriginalValue = GetDataRowValue(dataRow, dataMember, oldVersion)
    '                Value = GetDataRowValue(dataRow, dataMember, newVersion)
    '        End Select
    '        HistoryCommand = New SqlCommand(SQL, tcConnection, tcTransaction)
    '        With HistoryCommand
    '            .Parameters.Add(New System.Data.SqlClient.SqlParameter("@Who", mUserID))
    '            .Parameters.Add(New System.Data.SqlClient.SqlParameter("@TableName", dataRow.Table.TableName))
    '            .Parameters.Add(New System.Data.SqlClient.SqlParameter("@RecordID", dataRow(mTableIDColumn)))
    '            .Parameters.Add(New System.Data.SqlClient.SqlParameter("@Column", dataMember))
    '            .Parameters.Add(New System.Data.SqlClient.SqlParameter("@OriginalValue", OriginalValue))
    '            .Parameters.Add(New System.Data.SqlClient.SqlParameter("@Value", IIf(mMode = ActionModeEnum.modeDelete, "Record Deleted", Value)))
    '            .ExecuteNonQuery()
    '        End With
    '    Catch ex As SqlException
    '        Dim Message As String = String.Format("{0} ({1})", ex.Message, ex.GetType.Name)
    '        Message &= vbCrLf & "Do you want to continue without recording this History entry?"
    '        If MessageBox.Show(Me.ActiveForm, Message, ex.GetType.Name, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Throw
    '        If HistoryCommand IsNot Nothing Then HistoryCommand.Dispose() : HistoryCommand = Nothing
    '    End Try
    'End Sub
    Public Function UpdateDB(ByVal RootDir As String, ByVal DatabaseName As String) As Integer
        Dim ConnectionString As String
        Dim frm As frmProgress
        Dim Message As String
        Dim StartTime As Date

        Try
            Dim Root As DirectoryInfo = New DirectoryInfo(RootDir)
            If Not Root.Exists Then
                MessageBox.Show(String.Format("{0} does not exist!", RootDir), mAppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
                Return 1
            End If

            Message = $"{mAppName} - Listing Files From {Root.Name} into {DatabaseName}.Songs"
            StartTime = Now
            mEventLog.WriteEntry(Message)

            frm = New frmProgress(Me)
            With frm
                .Text = Message
                .OKtoClose = False
                .prgProgress.Visible = False
                If UIEnabled Then .Show()

                If UIEnabled Then RaiseEvent ListMessage(String.Format("Connecting to {0}", DatabaseName))
                Application.DoEvents() : If mCancel Then Exit Try

                ConnectionString = $"Application Name={mAppName};Data Source={SystemInformation.ComputerName};Initial Catalog={DatabaseName};Integrated Security=SSPI;Workstation ID={SystemInformation.ComputerName};"
                mSQLConnection = New SqlClient.SqlConnection(ConnectionString)
                mSQLConnection.Open()

                If UIEnabled Then RaiseEvent ListTitle("Determining File Count...")
                Application.DoEvents() : If mCancel Then Exit Try
                Try
                    Dim cntFolders As Long = 0
                    Dim cntFiles As Long = 0
                    CountFiles(Root, cntFolders, cntFiles)
                    mTotalFiles = cntFiles
                Catch ex As Exception
                End Try

                .prgProgress.Visible = True 'Establishes progress bar's maximum value
                mProgressCount = 0
                If UIEnabled Then RaiseEvent ListTitle("Cleaning up deleted files...")
                CleanupFiles(Root)
                .prgProgress.Visible = False

                .prgProgress.Visible = True 'Establishes progress bar's maximum value
                mProgressCount = 0
                If UIEnabled Then RaiseEvent ListTitle($"Processing audio files From {Root.Name} into {DatabaseName}.Songs...")
                ProcessFiles(Root)
                .prgProgress.Visible = False

                .OKtoClose = True
                .Close()
            End With

            Message = $"{mAppName} Complete - {mProgressCount:#,##0} entries written to {DatabaseName}.Songs with {mErrorCount:#,##0} errors.{vbCrLf}{vbCrLf}Elapsed Time: {ElapsedTime(StartTime, Now)}"
            mEventLog.WriteEntry(Message)
            MessageBox.Show(Message, mAppName, MessageBoxButtons.OK)
            Return 0
        Catch ex As Exception
            MessageBox.Show(ex.ToString, mAppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
            Return 1
        Finally
        End Try
    End Function
    Private Sub UpdateSong(ByVal Song As myTag, ByVal ID As Integer)
        Try
            Dim da As New SqlDataAdapter($"Select ID,Genre,Artist,Performer,ArtistID,Year,Album,AlbumID,Disc,Track,Title,Duration,Composer,Comment,Path,Lyrics,Publisher,DateModified From Songs Where ID={ID}", mSQLConnection)
            Dim builder As SqlCommandBuilder = New SqlCommandBuilder(da)
            Dim ds As New Data.DataSet
            Dim drv As DataRowView
            With da
                Dim RecordCount As Integer = .Fill(ds)
                drv = ds.Tables(0).DefaultView(0)

                drv.BeginEdit()
                Dim fUpdate As Boolean = False
                If drv("Genre") <> Song.Genre Then drv("Genre") = Song.Genre : fUpdate = True
                If drv("Artist") <> Song.Artist Then drv("Artist") = Song.Artist : drv("ArtistID") = GetArtist(Song) : fUpdate = True
                If drv("Performer") <> Song.Performer Then drv("Performer") = Song.Performer : fUpdate = True
                If drv("Year") <> Song.Year() Then drv("Year") = Song.Year : fUpdate = True
                If drv("Album") <> Song.Album Then drv("Album") = Song.Album : drv("AlbumID") = GetAlbum(Song) : fUpdate = True
                If drv("Disc") <> Song.Disc Then drv("Disc") = Song.Disc : fUpdate = True
                If drv("Track") <> Song.Track Then drv("Track") = Song.Track : fUpdate = True
                If drv("Title") <> Song.Title Then drv("Title") = Song.Title : fUpdate = True
                If drv("Duration") <> Song.Duration Then drv("Duration") = Song.Duration : fUpdate = True
                If drv("Composer") <> Song.Composer Then drv("Composer") = Song.Composer : fUpdate = True
                If drv("Comment") <> Song.Comment Then drv("Comment") = Song.Comment : fUpdate = True
                If drv("Lyrics") <> Song.Lyrics Then drv("Lyrics") = Song.Lyrics : fUpdate = True
                If drv("Publisher") <> Song.Publisher Then drv("Publisher") = Song.Publisher : fUpdate = True
                If Not fUpdate Then drv.CancelEdit() : Exit Try
                drv("DateModified") = Now()
                drv.EndEdit()
                .Update(ds)
            End With
        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Sub
#End Region
    'Entry point which delegates to C-style main Private Function
    Public Overloads Shared Sub Main()
        'System.Environment.ExitCode = MainCheckFiles(System.Environment.GetCommandLineArgs())
        System.Environment.ExitCode = MainUpdateDB(System.Environment.GetCommandLineArgs())
    End Sub
    Private Overloads Shared Function MainCheckFiles(ByVal args() As String) As Integer
        Dim fl As New clsMP3Update
        Try
            Application.EnableVisualStyles()
            Application.DoEvents()

            If args.Length < 2 Then
                Dim Message As String =
                    "MP3Update Options:" & vbCrLf &
                    vbTab & "args(0) = Full path name of the executable" & vbCrLf &
                    vbTab & "args(1) = Root directory to check"
                MessageBox.Show(Message, fl.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
                Return 1
            End If
            Dim RootDir As String = "D:\Users\kfc12\Music\Rock\Collins, Phil"  'args(1)
            Dim Root As DirectoryInfo = New DirectoryInfo(RootDir)
            If Not Root.Exists Then
                MessageBox.Show(String.Format("{0} does not exist!", RootDir), fl.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
                Return 1
            End If
            fl.CheckFiles(Root)
            Return 0
        Catch ex As Exception
            MessageBox.Show(ex.ToString, fl.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
            Return 1
        Finally
        End Try
    End Function
    Private Overloads Shared Function MainUpdateDB(ByVal args() As String) As Integer
        Dim fl As New clsMP3Update
        Try
            Application.EnableVisualStyles()
            Application.DoEvents()

            If args.Length < 3 Then
                Dim Message As String =
                    "MP3Update Options:" & vbCrLf &
                    vbTab & "args(0) = Full path name of the executable" & vbCrLf &
                    vbTab & "args(1) = Root directory to list files" & vbCrLf &
                    vbTab & "args(2) = Database name"
                MessageBox.Show(Message, fl.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
                Return 1
            End If
            Dim RootDir As String = args(1)
            'RootDir = "D:\Users\kfc12\Music\Soundtrack\Sound of Music, The (1965) - 50th Anniversary Edition"
            Dim DatabaseName As String = args(2)
            Return fl.UpdateDB(RootDir, DatabaseName)
        Catch ex As Exception
            MessageBox.Show(ex.ToString, fl.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
            Return 1
        Finally
        End Try
    End Function
End Class
#Region "Helper Classes"
Friend Class myTag
#Region "Properties"
    Public Property Path As String = ""
    Public Property Genre As String = ""
    Public Property Artist As String = ""
    Public Property Performer As String = ""
    Public Property Year As Integer = 0
    Public Property Album As String = ""
    Public Property Disc As Integer = 1
    Public Property Track As Integer = 0
    Public Property Title As String = ""
    Public Property Duration As Integer = 0
    Public Property Composer As String = ""
    Public Property Comment As String = ""
    Public Property Lyrics As String = ""
    Public Property Publisher As String = ""
#End Region
    Private Function DeDuplicate(ByVal Value As String) As String
        Dim test As String()
        test = Value.Split(";")
        If test.Count > 1 AndAlso test(0).Trim = test(1).Trim Then Value = test(0).Trim
        Return Value
    End Function
    Public Sub New(ByVal Tag As MP3.ID3v1Tag, ByVal Path As String)
        Me.Path = Path
        Me.Genre = Tag.Genre
        Me.Artist = Tag.Artist
        Me.Performer = Tag.Artist
        Me.Year = Int(Tag.Year)
        Me.Album = Tag.Album
        'Me.Disc = Int(IIf(Tag.Disc = 0, 1, Tag.Disc))
        Me.Track = Int(Tag.Track)
        Me.Title = Tag.Title
        'If Not IsNothing(Properties.Duration) Then Duration = Int(Tag.Duration.TotalSeconds)
        Me.Comment = Tag.Comment
    End Sub
    Public Sub New(ByVal tFile As TagLib.File, ByVal Path As String)
        Me.Path = Path
        If Not IsNothing(tFile.Tag.JoinedGenres) Then Me.Genre = DeDuplicate(tFile.Tag.JoinedGenres)
        If Not IsNothing(tFile.Tag.FirstAlbumArtist) Then
            'Seems sometimes MP3's have Artists separated by commas instead of semicolons...
            If tFile.Tag.FirstAlbumArtist.Length > 132 Then
                Dim tempArtists As String() = tFile.Tag.FirstAlbumArtist.Split(",")
                If tempArtists.Count > 0 Then Me.Artist = tempArtists(0) Else Me.Artist = tFile.Tag.FirstAlbumArtist.Substring(0, 132)
            Else
                Me.Artist = tFile.Tag.FirstAlbumArtist
            End If
        End If
        If Not IsNothing(tFile.Tag.JoinedPerformers) Then Me.Performer = DeDuplicate(tFile.Tag.JoinedPerformers)
        If Not IsNothing(tFile.Tag.Year) Then Me.Year = Int(tFile.Tag.Year)
        If Not IsNothing(tFile.Tag.Album) Then Me.Album = tFile.Tag.Album
        If Not IsNothing(tFile.Tag.Disc) Then Me.Disc = Int(IIf(tFile.Tag.Disc = 0, 1, tFile.Tag.Disc))
        If Not IsNothing(tFile.Tag.Track) Then Me.Track = Int(tFile.Tag.Track)
        If Not IsNothing(tFile.Tag.Title) Then Me.Title = tFile.Tag.Title
        If Not IsNothing(tFile.Properties.Duration) Then Me.Duration = Int(tFile.Properties.Duration.TotalSeconds)
        If Not IsNothing(tFile.Tag.JoinedComposers) Then Me.Composer = DeDuplicate(tFile.Tag.JoinedComposers)
        If Not IsNothing(tFile.Tag.Comment) Then Me.Comment = tFile.Tag.Comment
        If Not IsNothing(tFile.Tag.Lyrics) Then Me.Lyrics = tFile.Tag.Lyrics
        If Not IsNothing(tFile.Tag.Publisher) Then Me.Publisher = tFile.Tag.Publisher
    End Sub
    Public Sub Dump()
        Debug.WriteLine($"Path: {Path}")
        Debug.WriteLine($"Genre: {Genre}")
        Debug.WriteLine($"Artist: {Artist}")
        Debug.WriteLine($"Performer: {Performer}")
        Debug.WriteLine($"Year: {Year}")
        Debug.WriteLine($"Album: {Album}")
        Debug.WriteLine($"Disc: {Disc}")
        Debug.WriteLine($"Track: {Track}")
        Debug.WriteLine($"Title: {Title}")
        Debug.WriteLine($"Duration: {Duration}")
        Debug.WriteLine($"Composer: {Composer}")
        Debug.WriteLine($"Comment: {Comment}")
        Debug.WriteLine($"Lyrics: {Lyrics}")
        Debug.WriteLine($"Publisher: {Publisher}")
    End Sub
End Class
#End Region
