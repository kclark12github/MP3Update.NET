'clsMP3Update.vb
'   Populates SQL Server database with contents of Audio files in a given folder...
'   Copyright © 1998-2026, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Description:
'   11/21/26    Corrected file processing (enumerating a specific list of folders rather than the wonky semi-recursive code);
'               Corrected CleanupFiles which had been based on existing files (therefore would never have found anything missing);
'   11/15/26    Added History;
'=================================================================================================================================
Imports System.Collections.Generic
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Security.Policy
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports MailKit
Imports MailKit.Net.Smtp
Imports MailKit.Security
Imports MimeKit
Imports Serilog
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
        InitLog()
    End Sub
    Public Event ListMessage(ByVal Message As String)
    Public Event ListTitle(ByVal Title As String)
    Private UIEnabled As Boolean = False
    Private DBEnabled As Boolean = True
    Private mEventLog As EventLog
    Private mAppName As String
    Private mCancel As Boolean = False
    Private mErrorCount As Long = 0
    Private mErrorList As String = String.Empty
    Private mlogPath As String = String.Empty
    Private mProgressCount As Long = 0
    Private mStartTime As Date = Now()
    Private mTotalFiles As Long = 0
    Private mSQLConnection As SqlClient.SqlConnection
    Private targetFolders As String() = {"Christmas", "Classical", "Comedy", "Country", "Jazz", "New Age", "Radio Shows", "Rock", "Soundtrack"}
    Private allowedExtensions As String() = {".mp3", ".flac"}
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
    Public Sub CleanupFiles(ByVal DatabaseName As String)
        Try
            Dim deletedCount As Integer = 0
            Dim da As New SqlDataAdapter($"Select ID,Path From Songs", mSQLConnection)
            Dim builder As SqlCommandBuilder = New SqlCommandBuilder(da)
            Dim ds As New System.Data.DataSet
            With da
                Dim RecordCount As Integer = .Fill(ds)
                LogInfoMessage($"Processing {RecordCount:#,##0} Songs From {DatabaseName}.Songs...")
                For Each drv As DataRowView In ds.Tables(0).DefaultView
                    Try
                        If Not System.IO.File.Exists(drv("Path")) Then
                            LogInfoMessage($"File not found ""{drv("Path")}"" - deleting from Songs")
                            If DBEnabled Then DoCommand($"Delete From Songs Where Path='{drv("Path").Replace("'", "''").Replace("[", "[[]")}'") : deletedCount += 1
                        End If
                        mProgressCount += 1
                        If mProgressCount Mod 1000 = 0 OrElse mProgressCount = RecordCount Then LogInfoMessage($"Processed {mProgressCount:#,##0} of {RecordCount:#,##0}; Deleted {deletedCount:#,##0} ({mProgressCount / RecordCount:0.00%} w/{mErrorCount:#,##0} Errors - Elapsed: {ElapsedTime(mStartTime, Now)}")
                    Catch ex As Exception
                        mErrorCount += 1
                        LogErrorMessage(ex, $"Error processing {drv("Path")}; {vbCrLf}{vbCrLf}Exception: {ex.ToString}")
                    End Try
                Next
            End With
        Catch ex As Exception
            LogErrorMessage(ex, $"{mAppName}.CleanupFiles")
        Finally
            mProgressCount = 0
        End Try
    End Sub
    Public Sub CloseLog()
        Log.CloseAndFlush()
    End Sub
    Public Function CountFiles(ByVal BaseDir As DirectoryInfo, ByRef audioFiles As Generic.IEnumerable(Of FileInfo), ByVal ListFile As FileInfo) As Long
        Try
            If Not IsNothing(ListFile) Then
                Try
                    Dim commentPrefixes As String() = {"#", ";", "//"}
                    audioFiles = System.IO.File.ReadLines(ListFile.FullName) _
                        .Select(Function(line) New String(line.Where(Function(c) Not Char.IsControl(c) AndAlso
                                                                                 Char.GetUnicodeCategory(c) <> UnicodeCategory.Format).ToArray()).Trim()) _
                        .Where(Function(line) Not String.IsNullOrWhiteSpace(line) AndAlso
                                              Not commentPrefixes.Any(Function(prefix) line.StartsWith(prefix))) _
                        .Select(Function(line) line.Trim(""""c)) _
                        .Where(Function(p) Path.IsPathRooted(p)) _
                        .Select(Function(p) New FileInfo(p)) _
                        .ToList()
                Catch ex As NotSupportedException
                    Dim lineNumber As Integer = 0
                    For Each rawLine As String In System.IO.File.ReadLines(ListFile.FullName)
                        lineNumber += 1
                        Dim cleanPath As String = rawLine.Trim().Trim(""""c)
                        If String.IsNullOrWhiteSpace(cleanPath) Then Continue For
                        Try : Dim fi As New FileInfo(cleanPath) : Catch ex2 As NotSupportedException : Debug.WriteLine($"Line {lineNumber} failed: [{rawLine}]") : End Try
                    Next
                    Throw
                End Try
            ElseIf BaseDir.Name = "Music" Then
                audioFiles = targetFolders _
                    .Select(Function(folder) New DirectoryInfo(Path.Combine(BaseDir.FullName, folder))) _
                    .Where(Function(dir) dir.Exists) _
                    .SelectMany(Function(dir) dir.EnumerateFiles("*.*", SearchOption.AllDirectories)) _
                    .Where(Function(f) allowedExtensions.Contains(f.Extension.ToLower()))
            Else
                audioFiles = BaseDir.EnumerateFiles("*.*", SearchOption.AllDirectories) _
                    .Where(Function(f) allowedExtensions.Contains(f.Extension.ToLower()))
            End If
            Return audioFiles.LongCount
        Catch ex As Exception
            LogErrorMessage(ex, $"{mAppName}.CountFiles")
        End Try
    End Function
    Public Sub DeleteOrphans()
        Try
            'DBEnabled = False
            If DBEnabled Then DoCommand($"Delete From Albums Where Not Exists(Select * From Songs s Where Albums.ID=s.AlbumID);Delete From Artists Where Not Exists(Select * From Songs s Where Artists.ID=s.ArtistID);")
        Catch ex As Exception
            LogErrorMessage(ex, $"{mAppName}.DeleteOrphans")
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
    Private Sub DoMilestone()
        mProgressCount += 1
        If mProgressCount Mod 1000 <> 0 AndAlso mProgressCount <> mTotalFiles Then Return
        Dim Message As String = $"Processed {mProgressCount:#,##0} of {mTotalFiles:#,##0} ({mProgressCount / mTotalFiles:0.00%} w/{mErrorCount:#,##0} Errors - Elapsed: {ElapsedTime(mStartTime, Now)}"
        LogInfoMessage(Message)
        If mProgressCount = mTotalFiles Then mProgressCount = 0
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
    Private Function GetAKA(ByVal Artist As String) As String
        Dim akalist As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"Bachman - Turner Overdrive", "BTO"},
                {"Blue Öyster Cult", "BOC"},
                {"The Charlie Daniels Band", "CDB"},
                {"Creedence Clearwater Revival", "CCR"},
                {"Crosby, Stills & Nash", "CSN"},
                {"Crosby, Stills, Nash & Young", "CSNY"},
                {"Electric Light Orchestra", "ELO"},
                {"Emerson, Lake & Palmer", "ELP"},
                {"Fine Young Cannibals", "FYC"},
                {"Guns N' Roses", "GNR"},
                {"Little River Band", "LRB"},
                {"Stone Temple Pilots", "STP"},
                {"The Steve Miller Band", "SMB"}
            }
        Dim aka As String = Nothing
        If Not akalist.TryGetValue(Artist, aka) Then Return Nothing
        Return aka
    End Function
    Private Function GetAlphaSort(ByVal Artist As String) As String
        Dim exceptionlist As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"Bob Seger & The Silver Bullet Band", "SEGER, BOB, & THE SILVER BULLET BAND"},
            {"Bruce Hornsby & The Range", "HORNSBY, BRUCE, & THE RANGE"},
            {"Bruce Springsteen & The E Street Band", "SPRINGSTEEN, BRUCE, & THE E STREET BAND"},
            {"Edie Brickell & New Bohemians", "BRICKELL, EDIE, & NEW BOHEMIANS"},
            {"Elvis Costello & The Attractions", "COSTELLO, ELVIS, & THE ATTRACTIONS"},
            {"Elvis Costello & The Imposters", "COSTELLO, ELVIS, & THE IMPOSTERS"},
            {"Elvis Costello & The Roots", "COSTELLO, ELVIS, & THE ROOTS"},
            {"George Thorogood & The Destroyers", "THOROGOOD, GEORGE, & THE DESTROYERS"},
            {"Ringo Starr and His All Stard Band", "STARR, RINGO, AND HIS ALL-STARR BAND"},
            {"The Alan Parsons Project", "PARSONS, ALAN, PROJECT"},
            {"The Bob Seger System", "SEGER, BOB, SYSTEM"},
            {"The Steve Miller Band", "MILLER, STEVE, BAND"},
            {"Tom Petty and The Heartbreakers", "PETTY, TOM, AND THE HEARTBREAKERS"},
            {"Wolfgang Amadeus Mozart", "MOZART, WOLFGANG AMADEUS"}
        }
        Select Case Artist
            Case ".38 Special",
                "Adam Ant",
                "Alkaline Trio",
                "American Hi-Fi",
                "Bachman - Turner Overdrive",
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
                "Procol Harum",
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
                Dim test As String = Nothing
                If exceptionlist.TryGetValue(Artist, test) Then Return test
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
                    .Parameters.AddWithValue("@ArtistID", GetArtist(Song))
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
                Dim SQLSource As String = $"Insert Into Artists (Name,AlphaSort,AKA) Values (@Name,@AlphaSort,@AKA);Select SCOPE_IDENTITY();"
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
                    Dim aka As String = GetAKA(Artist)
                    Dim paramValue As Object = If(aka Is Nothing, CObj(DBNull.Value), aka)
                    .Parameters.AddWithValue("@AKA", paramValue)
                    Return .ExecuteScalar
                End With
            End If
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            Song.Dump()
            Throw
        End Try
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
    Public Sub InitLog()
        mStartTime = Now
        mlogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", $"{mAppName}-.log")
        Log.Logger = New LoggerConfiguration() _
            .MinimumLevel.Information() _
            .WriteTo.File(
                path:=mlogPath,
                rollingInterval:=RollingInterval.Day,
                retainedFileCountLimit:=30, ' Automatically deletes logs older than 30 days
                outputTemplate:="{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}"
            ) _
            .CreateLogger()
        LogInfoMessage($"[{mAppName} © {mStartTime:yyyy}, Ken Clark                       {mStartTime:MM/dd/yy hh:mm:ss tt}]")
    End Sub
    Public Sub LogErrorMessage(ByVal ex As Exception, ByVal Message As String)
        Log.Error(ex, Message)
        Debug.WriteLine(Message)
        mEventLog.WriteEntry(Message)
    End Sub
    Public Sub LogErrorMessage(ByVal Message As String)
        Log.Error(Message)
        Debug.WriteLine(Message)
        mEventLog.WriteEntry(Message)
    End Sub
    Public Sub LogFatalMessage(ByVal ex As Exception, ByVal Message As String)
        Log.Fatal(ex, Message)
        mEventLog.WriteEntry(Message)
    End Sub
    Public Sub LogFatalMessage(ByVal Message As String)
        Log.Fatal(Message)
        mEventLog.WriteEntry(Message)
    End Sub
    Public Sub LogInfoMessage(ByVal Message As String, Optional ByVal echo As Boolean = True)
        Log.Information(Message)
        If echo Then Debug.WriteLine(Message)
    End Sub
    Public Sub ProcessFiles(ByVal audioFiles As Generic.IEnumerable(Of FileInfo))
        Try
            For Each fi As FileInfo In audioFiles
                Dim Path As String = fi.FullName
                Try
                    If DBEnabled Then
                        Dim Song As myTag = New myTag(TagLib.File.Create(Path), Path)
                        Dim ID As Integer = DoScalar($"Select ID From Songs Where Path Like '%{Path.Replace("'", "''").Replace("[", "[[]")}'")
                        If IsNothing(ID) OrElse IsDBNull(ID) OrElse ID = 0 Then
                            AddSong(Song, Path)
                        Else
                            UpdateSong(Song, ID)
                        End If
                    End If
                    LogInfoMessage($"{vbTab}{vbTab}{Path}", False)
                    DoMilestone()
                Catch ex As Exception
                    mErrorCount += 1
                    LogErrorMessage($"{vbTab}{vbTab}{ex.GetType().Name}: {ex.Message} - {Path}")
                    mErrorList &= $"{ex.GetType().Name}: {ex.Message} - {Path}{vbCrLf}"
                End Try
            Next
        Catch ex As Exception
            LogErrorMessage(ex, $"{mAppName}.ProcessFiles")
        Finally
        End Try
    End Sub
    Public Sub SendSummaryEmail(subject As String, bodyText As String, Optional attachmentPath As String = Nothing, Optional isHtml As Boolean = False)
        ' Pull the exact same environment variables your PowerShell scripts use
        Dim smtpHost As String = Environment.GetEnvironmentVariable("SMTP_ADDRESS")
        Dim smtpPort As Integer = Integer.Parse(Environment.GetEnvironmentVariable("SMTP_PORT"))
        Dim smtpUser As String = Environment.GetEnvironmentVariable("SMTP_USER")
        Dim smtpPass As String = Environment.GetEnvironmentVariable("SMTP_PW")
        Dim myEmail As String = Environment.GetEnvironmentVariable("My_EMAIL")

        Dim message As New MimeMessage()
        message.From.Add(New MailboxAddress(Environment.UserName, myEmail))
        message.To.Add(New MailboxAddress(Environment.UserName, myEmail))
        message.Subject = subject

        Dim builder As New BodyBuilder()
        If isHtml Then
            builder.HtmlBody = bodyText
        Else
            builder.TextBody = bodyText
        End If

        If Not String.IsNullOrEmpty(attachmentPath) AndAlso System.IO.File.Exists(attachmentPath) Then
            builder.Attachments.Add(attachmentPath)
        End If

        message.Body = builder.ToMessageBody()
        Using client As New SmtpClient()
            ' SecureSocketOptions.Auto automatically negotiates SSL (port 465) or StartTLS (port 587)
            client.Connect(smtpHost, smtpPort, SecureSocketOptions.Auto)
            client.Authenticate(smtpUser, smtpPass)
            client.Send(New SendRequest(message))
            client.Disconnect(True)
        End Using
    End Sub
    Public Function UpdateDB(ByVal RootDir As String, ByVal DatabaseName As String, ByVal ListFile As FileInfo) As Integer
        Dim ConnectionString As String
        Dim Status As String = String.Empty
        Dim Root As DirectoryInfo = New DirectoryInfo(RootDir)
        Dim audioFiles As Generic.IEnumerable(Of FileInfo) = Nothing
        Try
            LogInfoMessage($"{mAppName} - Processing Files From {IIf(IsNothing(ListFile), Root.Name, ListFile.Name)} into {DatabaseName}.Songs; Database Updates Enabled: {DBEnabled}")
            If Not Root.Exists Then LogFatalMessage($"{RootDir} does Not exist!") : Return 1
            If Not IsNothing(ListFile) AndAlso Not ListFile.Exists Then LogFatalMessage($"{ListFile.Name} does Not exist!") : Return 1

            LogInfoMessage("Determining File Count...")
            Try : mTotalFiles = CountFiles(Root, audioFiles, ListFile) : Catch ex As Exception : End Try

            LogInfoMessage($"Connecting To {DatabaseName}")
            ConnectionString = $"Application Name= {mAppName};Data Source={SystemInformation.ComputerName};Initial Catalog={DatabaseName};Integrated Security=SSPI;Workstation ID={SystemInformation.ComputerName};"
            mSQLConnection = New SqlClient.SqlConnection(ConnectionString)
            mSQLConnection.Open()

            LogInfoMessage("Cleaning up deleted files...")
            CleanupFiles(DatabaseName)

            LogInfoMessage($"Processing {mTotalFiles:#,##0} audio files from {IIf(IsNothing(ListFile), Root.Name, ListFile.Name)} into {DatabaseName}.Songs...")
            ProcessFiles(audioFiles)

            LogInfoMessage($"Deleting unreferenced Artists and Albums in {IIf(IsNothing(ListFile), Root.Name, ListFile.Name)} from {DatabaseName}.Songs...")
            DeleteOrphans()

            Status = "Complete"
            Return 0
        Catch ex As Exception
            Status = "Failed"
            LogFatalMessage(ex, "Fatal error occurred during unattended MP3 import.")
            Environment.ExitCode = 1 ' Signals failure to Windows Task Scheduler
        Finally
            Dim Message As String = $"{mAppName} {Status} - {mTotalFiles:#,##0} entries written to {DatabaseName}.Songs with {mErrorCount:#,##0} errors.{vbCrLf}{vbCrLf}See attached log for details.{vbCrLf}{vbCrLf}Elapsed Time: {ElapsedTime(mStartTime, Now)}{vbCrLf}{vbCrLf}{mErrorList}"
            Log.Information(Message) : Debug.WriteLine(Message)
            CloseLog()
            SendSummaryEmail($"{mAppName} ""{Root.Name}"" {Status}", Message, Path.Combine($"{AppDomain.CurrentDomain.BaseDirectory}Logs", $"{mAppName}-{DateTime.Now:yyyyMMdd}.log"))
        End Try
    End Function
    Private Sub UpdateSong(ByVal Song As myTag, ByVal ID As Integer)
        Try
            Dim da As New SqlDataAdapter($"Select ID,Genre,Artist,Performer,ArtistID,Year,Album,AlbumID,Disc,Track,Title,Duration,Composer,Comment,Path,Lyrics,Publisher,DateModified From Songs Where ID={ID}", mSQLConnection)
            Dim builder As SqlCommandBuilder = New SqlCommandBuilder(da)
            Dim ds As New System.Data.DataSet
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
    Private Overloads Shared Function MainUpdateDB(ByVal args() As String) As Integer
        Dim udb As New clsMP3Update
        Try
            If args.Length < 3 Then
                Dim Message As String =
                    $"MP3Update Options:{vbCrLf}" &
                    $"{vbTab}args(0) = Full path name of the executable{vbCrLf}" &
                    $"{vbTab}args(1) = Root directory{vbCrLf}" &
                    $"{vbTab}args(2) = Database name{vbCrLf}" &
                    $"{vbTab}args(3) = List File (optional - overrides Root directory)"
                Throw New ArgumentException(Message)
            End If
            Dim RootDir As String = args(1)
            Dim DatabaseName As String = args(2)
            Dim ListFile As FileInfo = Nothing
            If args.Length = 4 Then ListFile = New FileInfo(args(3))
            Return udb.UpdateDB(RootDir, DatabaseName, ListFile)
        Catch ex As Exception
            Dim Message As String = $"{udb.AppName} Failed - {ex.GetType().Name}{vbCrLf}{vbCrLf}{ex.Message}"
            udb.LogFatalMessage(ex.Message)
            udb.CloseLog()
            udb.SendSummaryEmail($"{udb.AppName} Failed - {ex.GetType().Name}", Message, Path.Combine($"{AppDomain.CurrentDomain.BaseDirectory}Logs", $"{udb.AppName}-{DateTime.Now:yyyyMMdd}.log"))
            Return 1
        End Try
    End Function
End Class
#Region "Helper Classes"
Friend Class myTag
#Region "Properties"
    Dim mAlbum As String = ""
    Dim mArtist As String = ""
    Dim mComment As String = ""
    Dim mComposer As String = ""
    Dim mDisc As Integer = 1
    Dim mDuration As Integer = 0
    Dim mGenre As String = ""
    Dim mLyrics As String = ""
    Dim mPath As String = ""
    Dim mPerformer As String = ""
    Dim mPublisher As String = ""
    Dim mTitle As String = ""
    Dim mTrack As Integer = 0
    Dim mYear As Integer = 0
    Public Property Album As String
        Get
            Return mAlbum
        End Get
        Set(value As String)
            mAlbum = Normalize(value)
        End Set
    End Property
    Public Property Artist As String
        Get
            Return mArtist
        End Get
        Set(value As String)
            mArtist = Normalize(value)
        End Set
    End Property
    Public Property Comment As String
        Get
            Return mComment
        End Get
        Set(value As String)
            mComment = Normalize(value)
        End Set
    End Property
    Public Property Composer As String
        Get
            Return mComposer
        End Get
        Set(value As String)
            mComposer = Normalize(value)
        End Set
    End Property
    Public Property Disc As Integer
        Get
            Return mDisc
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then mDisc = IIf(value = 0, 1, value)
        End Set
    End Property
    Public Property Duration As Integer
        Get
            Return mDuration
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then mDuration = value
        End Set
    End Property
    Public Property Genre As String
        Get
            Return mGenre
        End Get
        Set(value As String)
            mGenre = Normalize(value)
        End Set
    End Property
    Public Property Lyrics As String
        Get
            Return mLyrics
        End Get
        Set(value As String)
            mLyrics = Normalize(value)
        End Set
    End Property
    Public Property Path As String
        Get
            Return mPath
        End Get
        Set(value As String)
            mPath = value
        End Set
    End Property
    Public Property Performer As String
        Get
            Return mPerformer
        End Get
        Set(value As String)
            mPerformer = Normalize(value)
        End Set
    End Property
    Public Property Publisher As String
        Get
            Return mPublisher
        End Get
        Set(value As String)
            mPublisher = Normalize(value)
        End Set
    End Property
    Public Property Track As Integer
        Get
            Return mTrack
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then mTrack = value
        End Set
    End Property
    Public Property Title As String
        Get
            Return mTitle
        End Get
        Set(value As String)
            mTitle = Normalize(value)
        End Set
    End Property
    Public Property Year As Integer
        Get
            Return mYear
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then mYear = value
        End Set
    End Property
#End Region
    Private Function DeDuplicate(ByVal Value As String) As String
        Dim test As String()
        test = Value.Split(";")
        If test.Count > 1 AndAlso test(0).Trim = test(1).Trim Then Value = test(0).Trim

        'Seems sometimes MP3's have Artists separated by commas instead of semicolons (and the TabLib is a bit buggy)...
        If Value.Length > 132 Then
            test = Value.Split(",")
            If test.Count > 0 Then Value = test(0) Else Value = Value.Substring(0, 132)
        End If
        Return Value
    End Function
    Private Function Normalize(ByVal rawValue As String) As String
        If String.IsNullOrEmpty(rawValue) Then Return String.Empty

        'Standardize Unicode composition (fixes macOS/iTunes decomposed accents)
        Dim cleaned As String = rawValue.Normalize(NormalizationForm.FormC)

        'Remove null terminators, zero-width characters, BOMs, and control chars (except whitespace)
        cleaned = Regex.Replace(cleaned, "[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F\u200B-\u200D\uFEFF]", String.Empty)

        'Normalize all Unicode dashes and soft hyphens to standard ASCII hyphen (45)
        cleaned = Regex.Replace(cleaned, "[\p{Pd}\u00AD\u2212]", "-")

        'Normalize smart single quotes/primes to ASCII apostrophe (39)
        cleaned = Regex.Replace(cleaned, "[\u2018\u2019\u201A\u201B\u2032]", "'")

        'Normalize smart double quotes to ASCII double quote (34)
        cleaned = Regex.Replace(cleaned, "[\u201C\u201D\u201E\u201F\u2033]", """")

        'Collapse all whitespace variations (NBSP, tabs, newlines, double spaces) into a single space
        cleaned = Regex.Replace(cleaned, "[\s\p{Z}]+", " ").Trim()

        'Normalizes "ft.", "feat", "featuring" (case-insensitive) to "feat."
        If Not cleaned.Contains("Little Feat") Then cleaned = Regex.Replace(cleaned, "\b(featuring|feat|ft)\.?\b", "feat.", RegexOptions.IgnoreCase)

        Return DeDuplicate(cleaned)
    End Function
    Private Function NormalizeArtist(ByVal Value As String) As String
        If String.IsNullOrEmpty(Value) Then Return ""
        Dim slashWhitelist As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"AC; DC", "AC/DC"},
            {"Tom Petty & The Heartbreakers", "Tom Petty and The Heartbreakers"}
        }
        Dim restoredValue As String = Nothing
        If slashWhitelist.TryGetValue(Value, restoredValue) Then Return restoredValue
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
        Me.Genre = tFile.Tag.JoinedGenres
        Dim AlbumArtist As String = NormalizeArtist(tFile.Tag.JoinedAlbumArtists)
        Me.Artist = IIf(AlbumArtist <> tFile.Tag.FirstAlbumArtist, AlbumArtist, tFile.Tag.FirstAlbumArtist)
        Me.Performer = NormalizeArtist(tFile.Tag.JoinedPerformers)
        Me.Year = Int(tFile.Tag.Year)
        Me.Album = tFile.Tag.Album
        Me.Disc = Int(tFile.Tag.Disc)
        Me.Track = Int(tFile.Tag.Track)
        Me.Title = tFile.Tag.Title
        Me.Duration = Int(tFile.Properties.Duration.TotalSeconds)
        Me.Composer = NormalizeArtist(tFile.Tag.JoinedComposers)
        Me.Comment = tFile.Tag.Comment
        Me.Lyrics = tFile.Tag.Lyrics
        Me.Publisher = tFile.Tag.Publisher
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
