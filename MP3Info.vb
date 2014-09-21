Imports System.IO
Imports System.Text
Namespace MP3
    Public Enum MPEGVersionEnum As Byte
        MPEG25 = 0  'MPEG Version 2.5 (later extension of MPEG 2)
        Reserved = 1
        MPEG2 = 2   'MPEG Version 2 (ISO/IEC 13818-3)
        MPEG1 = 3   'MPEG Version 1 (ISO/IEC 11172-3)
    End Enum
    Public Enum LayerEnum As Byte
        Reserved = 0
        LayerIII = 1    'Layer III
        LayerII = 2     'Layer II
        LayerI = 3      'Layer I
    End Enum
    Public Enum ProtectionEnum As Byte
        CRC = 0     'Protected by CRC (16bit CRC follows header)
        None = 1    'Not protected
    End Enum
    Public Enum ChannelModeEnum As Byte
        Stereo = 0
        JointStereo = 1
        DualChannel = 2     'Dualo channel (2 mono channels)
        SingleChannel = 3   'Single channel (Mono)
    End Enum
    Public Enum EmphasisEnum As Byte
        None = 0
        MS5015 = 1  '50/15 ms
        Reserved = 2
        CCIT = 3    'CCIT J.17
    End Enum
    Public Enum EncodingEnum As Byte
        CBR = 0 'Constant bitrate
        VBR = 1 'Variable bitrate
    End Enum
    Public Enum ID3v1TagVersionEnum As Byte
        Version10 = 10  'Tag version 1.0
        Version11 = 11  'Tag version 1.1
    End Enum
    'Provides methods to retrieve informations about MPEG files and to get and set 
    'different tag formats.

    '<para>Use the <b>MP3Info</b> class for typical operations on MPEG files, like reading 
    'and writing ID3 v1 and v2 tags or getting general MPEG informations.
    '<note type="note">
    'In members that accept a path as an input string, that path must be well-formed 
    'or an exception is raised. For example, if a path is fully qualified but begins 
    'with a space, the path is not trimmed in methods of the class. Therefore, the path 
    'is malformed and an exception is raised. Similarly, a path or a combination of 
    'paths cannot be fully qualified twice. For example, "c:\temp c:\windows" also 
    'raises an exception in most cases. Ensure that your paths are well-formed when 
    'using methods that accept a path string.
    '</note></para>
    '<para>This class was made by using informations from the following websites:
    '<list type="bullet">
    '<item><a href="http://gabriel.mp3-tech.org/mp3infotag.html">http://gabriel.mp3-tech.org/mp3infotag.html</a></item>
    '<item><a href="http://www.multiweb.cz/twoinches/MP3inside.htm#MP3FileStructure">http://www.multiweb.cz/twoinches/MP3inside.htm#MP3FileStructure</a></item>
    '<item><a href="http://www.codeguru.com/vb/gen/vb_multimedia/mp3s/article.php/c4267">http://www.codeguru.com/vb/gen/vb_multimedia/mp3s/article.php/c4267</a></item>
    '<item><a href="http://www.getid3.org/">http://www.getid3.org/</a></item>
    '<item><a href="http://www.mp3-converter.com/mp3codec/mp3_anatomy.htm">http://www.mp3-converter.com/mp3codec/mp3_anatomy.htm</a></item>
    '<item><a href="http://www.mp3-tech.org/">http://www.mp3-tech.org/</a></item>
    '<item><a href="http://www.codeproject.com/audio/MPEGAudioInfo.asp">http://www.codeproject.com/audio/MPEGAudioInfo.asp</a></item>
    '</list></para>
    ''' </remarks>
    ''' <example>
    '''		The following example demonstrates some of the main members of the 
    '''		<b>MP3Info</b> class.
    '''		<code language="VB.NET">
    '''		[Visual Basic]
    '''		Dim objInfo As New MP3Info("c:\test.mp3")
    '''		...
    '''		</code>
    ''' </example>
    Public Class MP3Info
        'Declarations for friend variables holding the property values
        Friend f_strFilename As String
        Friend f_intHeaderPos As Integer
        Friend f_blnVBR As Boolean
        Friend f_intAudioSize As Integer
        Friend f_objID3v1Tag As ID3v1Tag

        'Declarations for internal variables
        Private m_objHeaderBits As New BitArray(24)
        Private m_objXingHeader As XingHeaderStructure
        Private Structure XingHeaderStructure
            Dim Flags As Integer
            Dim FrameCount As Integer
            Dim FileLenght As Integer
            Dim TOC() As Integer
            Dim Quality As Integer
        End Structure
        Public Sub New()
            MyBase.New()
        End Sub
        Public Sub New(ByVal strFilename As String)
            Me.Filename = strFilename
        End Sub
        Public Property Filename() As String
            Get
                Return f_strFilename
            End Get
            Set(ByVal Value As String)
                If (File.Exists(Value)) Then
                    f_strFilename = Value
                    If (Not ReadHeaders()) Then
                        Throw New Exception("The specified file is not a valid MP3 file.")
                    Else
                        f_objID3v1Tag = GetID3v1Tag()
                    End If
                Else
                    Throw New IO.IOException("The specified file doesn't exist.")
                End If
            End Set
        End Property
        Public ReadOnly Property Filesize() As Integer
            Get
                Dim objFI As New FileInfo(f_strFilename)
                Return objFI.Length
            End Get
        End Property
        Public ReadOnly Property HeaderPosition() As Integer
            Get
                Return f_intHeaderPos
            End Get
        End Property
        Public Property ID3v1Tag() As ID3v1Tag
            Get
                Return f_objID3v1Tag
            End Get
            Set(ByVal Value As ID3v1Tag)
                f_objID3v1Tag = Value
            End Set
        End Property
        'The range of these value starts with "0" (best) and ends with "100" (worst).
        'The property will return "-1" if the file was encoded using a constant 
        'bitrate.
        Public ReadOnly Property VBRScale() As Integer
            Get
                If (Me.Encoding = EncodingEnum.CBR) Then
                    Return -1
                Else
                    Return m_objXingHeader.Quality
                End If
            End Get
        End Property
        Public ReadOnly Property Encoding() As EncodingEnum
            Get
                Return IIf(f_blnVBR, EncodingEnum.VBR, EncodingEnum.CBR)
            End Get
        End Property
        'If the file is encoded with a variable bitrate, this value describes the average frame size.
        Public ReadOnly Property FrameSize() As Integer
            Get
                If (Me.Encoding = EncodingEnum.CBR) Then
                    Return ((Me.FrameSamples / 8 * Me.Bitrate) / Me.SamplingRateFrequency) + Me.Padding
                Else
                    Return Math.Round(m_objXingHeader.FileLenght / m_objXingHeader.FrameCount, 0)
                End If
            End Get
        End Property
        Public ReadOnly Property FrameSamples() As Integer
            Get
                Select Case Me.Layer
                    Case LayerEnum.LayerI
                        Return 384
                    Case LayerEnum.LayerII
                        Return 1152
                    Case LayerEnum.LayerIII
                        If (Me.MPEGVersion = MPEGVersionEnum.MPEG1) Then
                            Return 1152
                        Else
                            Return 576
                        End If
                End Select
            End Get
        End Property
        'MPEG Version 2.5 was added lately to the MPEG 2 standard. It is an extension 
        'used for very low bitrate files, allowing the use of lower sampling frequencies. 
        'If your decoder does not support this extension, it is recommended for you to 
        'use 12 bits for synchronization instead of 11 bits.
        Public ReadOnly Property MPEGVersion() As MPEGVersionEnum
            Get
                Select Case BitsToString(m_objHeaderBits, 3, 4)
                    Case "11"
                        Return MPEGVersionEnum.MPEG1
                    Case "10"
                        Return MPEGVersionEnum.MPEG2
                    Case "00"
                        Return MPEGVersionEnum.MPEG25
                    Case "01"
                        Return MPEGVersionEnum.Reserved
                End Select
            End Get
        End Property
        Public ReadOnly Property Layer() As LayerEnum
            Get
                Select Case BitsToString(m_objHeaderBits, 5, 6)
                    Case "01"
                        Return LayerEnum.LayerIII
                    Case "10"
                        Return LayerEnum.LayerII
                    Case "11"
                        Return LayerEnum.LayerI
                    Case "00"
                        Return LayerEnum.Reserved
                End Select
            End Get
        End Property
        'If this property returns CRC, the file is protected by CRC (16bit CRC follows header).
        Public ReadOnly Property Protection() As ProtectionEnum
            Get
                If (Not m_objHeaderBits.Item(7)) Then
                    Return ProtectionEnum.CRC
                Else
                    Return ProtectionEnum.None
                End If
            End Get
        End Property
        'Dual channel files are made of two independant mono channel. Each one uses 
        'exactly half the bitrate of the file. Most decoders output them as stereo, 
        'but it might not always be the case.
        '
        'One example of use would be some speech in two different languages carried 
        'in the same bitstream, and then an appropriate decoder would decode only 
        'the choosen language.
        Public ReadOnly Property ChannelMode() As ChannelModeEnum
            Get
                Select Case BitsToString(m_objHeaderBits, 16, 17)
                    Case "00"
                        Return ChannelModeEnum.Stereo
                    Case "01"
                        Return ChannelModeEnum.JointStereo
                    Case "10"
                        Return ChannelModeEnum.DualChannel
                    Case "11"
                        Return ChannelModeEnum.SingleChannel
                End Select
            End Get
        End Property
        Public ReadOnly Property SamplingRateFrequency() As Integer
            Get
                Select Case BitsToString(m_objHeaderBits, 12, 13)
                    Case "00"
                        Select Case Me.MPEGVersion
                            Case MPEGVersionEnum.MPEG1
                                Return 44100
                            Case MPEGVersionEnum.MPEG2
                                Return 22050
                            Case MPEGVersionEnum.MPEG25
                                Return 11025
                        End Select
                    Case "01"
                        Select Case Me.MPEGVersion
                            Case MPEGVersionEnum.MPEG1
                                Return 48000
                            Case MPEGVersionEnum.MPEG2
                                Return 24000
                            Case MPEGVersionEnum.MPEG25
                                Return 12000
                        End Select
                    Case "10"
                        Select Case Me.MPEGVersion
                            Case MPEGVersionEnum.MPEG1
                                Return 32000
                            Case MPEGVersionEnum.MPEG2
                                Return 16000
                            Case MPEGVersionEnum.MPEG25
                                Return 8000
                        End Select
                    Case "11"
                        Return -1
                End Select
            End Get
        End Property
        'The copyright has the same meaning as the copyright bit on CDs and DAT tapes, 
        'i.e. telling that it is illegal to copy the contents if the bit is set.
        Public ReadOnly Property Copyright() As Boolean
            Get
                Return m_objHeaderBits.Item(20)
            End Get
        End Property
        'The return value "0" means free format. The free bitrate must remain 
        'constant, an must be lower than the maximum allowed bitrate. Decoders are 
        'not required to support decoding of free bitrate streams.
        '
        '	The return value "-1" means that the value is unallowed. 
        '
        '	If the file is encoded with a variable bitrate, this value describes the
        '	average bitrate.
        Public ReadOnly Property Bitrate() As Integer
            Get
                If (Me.Encoding = EncodingEnum.CBR) Then

                    Dim x As Integer = (-m_objHeaderBits.Item(8)) * 8 + (-m_objHeaderBits.Item(9)) * 4 + (-m_objHeaderBits.Item(10)) * 2 + (-m_objHeaderBits.Item(11)) * 1
                    If ((Me.MPEGVersion = MPEGVersionEnum.MPEG1) And (Me.Layer = LayerEnum.LayerI)) Then
                        Return Choose(x + 1, 0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448, -1) * 1000
                    End If
                    If ((Me.MPEGVersion = MPEGVersionEnum.MPEG1) And (Me.Layer = LayerEnum.LayerII)) Then
                        Return Choose(x + 1, 0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384, -1) * 1000
                    End If
                    If ((Me.MPEGVersion = MPEGVersionEnum.MPEG1) And (Me.Layer = LayerEnum.LayerIII)) Then
                        Return Choose(x + 1, 0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, -1) * 1000
                    End If
                    If (((Me.MPEGVersion = MPEGVersionEnum.MPEG2) Or (Me.MPEGVersion = MPEGVersionEnum.MPEG25)) And (Me.Layer = LayerEnum.LayerI)) Then
                        Return Choose(x + 1, 0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256, -1) * 1000
                    End If
                    If (((Me.MPEGVersion = MPEGVersionEnum.MPEG2) Or (Me.MPEGVersion = MPEGVersionEnum.MPEG25)) And ((Me.Layer = LayerEnum.LayerII) Or (Me.Layer = LayerEnum.LayerIII))) Then
                        Return Choose(x + 1, 0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160, -1) * 1000
                    End If

                Else

                    Return Math.Round(((Me.FrameSize * Me.SamplingRateFrequency) / 144) / 1000, 0) * 1000

                End If
            End Get
        End Property
        'Padding is used to exactly fit the bitrate. As an example: 128kbps 44.1kHz 
        'layer II uses a lot of 418 bytes and some of 417 bytes long frames to get 
        'the exact 128k bitrate. For Layer I slot is 32 bits long, for Layer II and 
        'Layer III slot is 8 bits long.
        Public ReadOnly Property Padding() As Integer
            Get
                If (m_objHeaderBits.Item(14)) Then
                    If (Me.Layer = LayerEnum.LayerI) Then
                        Return 4
                    Else
                        Return 1
                    End If
                Else
                    Return 0
                End If
            End Get
        End Property
        Public ReadOnly Property PrivateBit() As Boolean
            Get
                Return m_objHeaderBits.Item(15)
            End Get
        End Property
        'The original bit indicates, if it is set, that the frame is located on its original media.
        Public ReadOnly Property OriginalBit() As Boolean
            Get
                Return m_objHeaderBits.Item(21)
            End Get
        End Property
        'The emphasis indication is here to tell the decoder that the file must 
        'be de-emphasized, ie the decoder must "re-equalize" the sound after a 
        'Dolby-like noise supression. It is rarely used.
        Public ReadOnly Property Emphasis() As EmphasisEnum
            Get
                Select Case BitsToString(m_objHeaderBits, 22, 23)
                    Case "00"
                        Return EmphasisEnum.None
                    Case "01"
                        Return EmphasisEnum.MS5015
                    Case "10"
                        Return EmphasisEnum.Reserved
                    Case "11"
                        Return EmphasisEnum.CCIT
                End Select
            End Get
        End Property
        Public ReadOnly Property Length() As Integer
            Get
                Return Math.Round(f_intAudioSize / Me.Bitrate * 8, 0)
            End Get
        End Property
        Public Sub Update()
            SetID3v1Tag()
        End Sub
        Private Function GetID3v1Tag() As ID3v1Tag

            ''' Declarations
            Dim objFS As FileStream
            Dim objReader As BinaryReader
            Dim bytBytes() As Byte
            Dim strBytes As String
            Dim objTag As ID3v1Tag

            ''' Open the filestream
            objFS = New FileStream(f_strFilename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

            ''' Create the BinaryReader object to read the file
            objReader = New BinaryReader(objFS)

            ''' Create a new object to hold the tag content
            objTag = New ID3v1Tag

            ' By default, there's no tag
            objTag.m_blnTagAvailable = False

            ''' The v1 tag should be found at the end of the file and has a fixed size of 
            ''' 128 bytes, so have a look if there's one
            objFS.Seek(-128, SeekOrigin.End)
            strBytes = BytesToString(objReader.ReadBytes(3))
            If (strBytes.ToUpper = "TAG") Then

                ' We found a tag
                objTag.m_blnTagAvailable = True

                With objTag

                    ''' Read title, remove possible junk bytes and assign value to the property
                    strBytes = BytesToString(objReader.ReadBytes(30)).Replace(Chr(0), "")
                    .m_strTitle = strBytes

                    ''' Read artist, remove possible junk bytes and assign value to the property
                    strBytes = BytesToString(objReader.ReadBytes(30)).Replace(Chr(0), "")
                    .m_strArtist = strBytes

                    ''' Read album, remove possible junk bytes and assign value to the property
                    strBytes = BytesToString(objReader.ReadBytes(30)).Replace(Chr(0), "")
                    .m_strAlbum = strBytes

                    ''' Read year
                    strBytes = BytesToString(objReader.ReadBytes(4)).Replace(Chr(0), "")
                    .m_strYear = strBytes

                    ''' Read 30 bytes for comment and track
                    bytBytes = objReader.ReadBytes(30)

                    ''' If byte 28 is zero and byte 29 non-zero, the tag has version 1.1, 
                    ''' otherwise 1.0
                    If ((bytBytes(28) = 0) And (bytBytes(29) <> 0)) Then
                        .m_bytTagVersion = ID3v1TagVersionEnum.Version11
                        strBytes = BytesToString(bytBytes, 0, 28).Replace(Chr(0), "")
                        .m_bytTrack = bytBytes(29)
                    Else
                        .m_bytTagVersion = ID3v1TagVersionEnum.Version10
                        strBytes = BytesToString(bytBytes).Replace(Chr(0), "")
                        .m_bytTrack = 0
                    End If
                    .m_strComment = strBytes

                    ''' Read genre
                    .m_bytGenre = objReader.ReadByte

                End With

            End If

            ''' Close the reader and the filestream object
            objReader.Close()
            objFS.Close()

            ''' Return the ID3v1 tag object
            Return objTag

        End Function
        Private Sub SetID3v1Tag()

            ''' Declarations
            Dim objFS As FileStream
            Dim objReader As BinaryReader
            Dim objWriter As BinaryWriter
            Dim strBytes As String

            ''' Open the filestream
            objFS = New FileStream(f_strFilename, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)

            ''' Create the BinaryReader object to read the file
            objReader = New BinaryReader(objFS)

            ''' The v1 tag should be found at the end of the file and has a fixed size of 
            ''' 128 bytes, so have a look if there's one
            objFS.Seek(-128, SeekOrigin.End)
            strBytes = BytesToString(objReader.ReadBytes(3))
            If (strBytes.ToUpper = "TAG") Then

                ''' There is an existing tag, so rewind filestream position 3 bytes
                objFS.Seek(-3, SeekOrigin.Current)

            Else

                ''' No tag found, so set filestream position to end of file
                objFS.Seek(0, SeekOrigin.End)

            End If

            ''' Create the BinaryWriter object to write the tag
            objWriter = New BinaryWriter(objFS)

            ''' Write the tag
            objWriter.Write("TAG".ToCharArray)
            objWriter.Write(Me.ID3v1Tag.Title.PadRight(30, Chr(0)).ToCharArray)
            objWriter.Write(Me.ID3v1Tag.Artist.PadRight(30, Chr(0)).ToCharArray)
            objWriter.Write(Me.ID3v1Tag.Album.PadRight(30, Chr(0)).ToCharArray)
            objWriter.Write(Me.ID3v1Tag.Year.PadRight(4, Chr(0)).ToCharArray)
            Select Case Me.ID3v1Tag.TagVersion
                Case ID3v1TagVersionEnum.Version10
                    objWriter.Write(Me.ID3v1Tag.Comment.PadRight(30, Chr(0)).ToCharArray)
                Case ID3v1TagVersionEnum.Version11
                    objWriter.Write(Me.ID3v1Tag.Comment.PadRight(28, Chr(0)).ToCharArray)
                    objWriter.Write(Chr(0))
                    objWriter.Write(Me.ID3v1Tag.Track)
            End Select
            objWriter.Write(Me.ID3v1Tag.Genre)
            objWriter.Flush()

            ''' Close the BinaryReader and BinaryWriter objects and the base stream
            objWriter.Close()
            objReader.Close()
        End Sub
        Private Function ReadHeaders() As Boolean

            ''' Declarations
            Dim objFS As FileStream
            Dim objBitArray As BitArray
            Dim bytBytes(2) As Byte
            Dim bytXingBytes(3) As Byte
            Dim i, j, intIndex, intOffset, intStart As Integer
            Dim objConverter As New System.Text.UTF8Encoding
            Dim blnReturn As Boolean = False
            Dim intXingOffset As Integer
            Dim bytFlags(3), bytFrameCount(3), bytFileLength(3), bytQuality(3) As Byte
            Dim intTOC(99) As Integer
            Dim objBitConverter As BitConverter
            Dim bytXingHeader(116) As Byte

            ''' Open the filestream
            Try
                objFS = New FileStream(f_strFilename, FileMode.Open)
                If (objFS.CanRead) Then
                    objFS.Position = 0
                Else
                    Throw New IO.IOException("Can't read file.")
                End If
            Catch Ex As Exception
                Throw New IO.IOException("An error occurred while trying to open the file.")
            End Try

            ''' Read the headers
            Try

                ''' Read the MPEG header
                While ((objFS.Position + 4) <= objFS.Length)

                    ''' Read a byte from file and check if the bits are all set 
                    If (objFS.ReadByte = 255) Then

                        ''' Read the next 3 bytes (maybe the complete header)
                        objFS.Read(bytBytes, 0, 3)
                        objBitArray = New BitArray(bytBytes)

                        ''' If bits 9 to 11 are set, we found the header
                        If (objBitArray(7)) And (objBitArray(6)) And (objBitArray(5)) Then

                            ''' Store the header position
                            f_intHeaderPos = objFS.Position

                            ''' Change the bit order to a more readable format
                            intIndex = 0
                            intOffset = 0
                            For j = 0 To 2
                                For i = 7 To 0 Step -1
                                    m_objHeaderBits.Item(intIndex) = objBitArray.Item(intOffset + i)
                                    intIndex += 1
                                Next
                                intOffset += 8
                            Next

                            blnReturn = True
                            Exit While

                        Else

                            ''' Rewind 3 bytes
                            objFS.Position -= 3

                        End If

                    End If

                End While

                ''' Adjust audio size
                f_intAudioSize = Me.Filesize - f_intHeaderPos

                ''' Read the extended (Xing) header
                ''' Set position of filestream (Xing header offset)
                If (Me.MPEGVersion = MPEGVersionEnum.MPEG1) Then
                    If (Me.ChannelMode = ChannelModeEnum.SingleChannel) Then
                        intXingOffset = 17
                    Else
                        intXingOffset = 32
                    End If
                Else
                    If (Me.ChannelMode = ChannelModeEnum.SingleChannel) Then
                        intXingOffset = 9
                    Else
                        intXingOffset = 17
                    End If
                End If
                objFS.Position = Me.HeaderPosition + intXingOffset

                ''' Read 4 bytes
                objFS.Read(bytXingBytes, 0, 4)

                ''' Check for Xing header
                If (objConverter.GetString(bytXingBytes).ToLower = "xing") Or (objConverter.GetString(bytXingBytes).ToLower = "info") Then

                    ''' Read the extended (Xing) header
                    objFS.Read(bytXingHeader, 0, 116)

                    ''' Check the encoding
                    If (objConverter.GetString(bytXingBytes).ToLower = "xing") Then
                        f_blnVBR = True
                    End If

                    ''' Adjust audio size
                    f_intAudioSize -= 120

                    ''' Populate the extended (Xing) header
                    For i = 3 To 0 Step -1
                        bytFlags(3 - i) = bytXingHeader(i)
                        bytFrameCount(3 - i) = bytXingHeader(4 + i)
                        bytFileLength(3 - i) = bytXingHeader(8 + i)
                        bytQuality(3 - i) = bytXingHeader(112 + i)
                    Next

                    For i = 0 To 99
                        intTOC(i) = CType(bytXingHeader(12 + i), Integer)
                    Next

                    With m_objXingHeader
                        .Flags = objBitConverter.ToInt32(bytFlags, 0)
                        .FileLenght = objBitConverter.ToInt32(bytFileLength, 0)
                        .FrameCount = objBitConverter.ToInt32(bytFrameCount, 0)
                        .TOC = intTOC
                        .Quality = objBitConverter.ToInt32(bytQuality, 0)
                    End With

                End If

                ''' Read the extended (VBRI) header
                ''' Set position of filestream (VBRI header offset)
                objFS.Position = Me.HeaderPosition + 32

                ''' Read 4 bytes
                objFS.Read(bytXingBytes, 0, 4)

                ''' Check for Xing header
                If (objConverter.GetString(bytXingBytes).ToLower = "vbri") Then

                    ''' TBD

                End If

            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                objFS.Close()
            End Try
            Return blnReturn
        End Function
        'A System.String BitsToString  value...
        Private Function BitsToString(ByVal objBitArray As BitArray, ByVal intStart As Integer, ByVal intEnd As Integer) As String
            Dim strBits As String = ""

            For i As Integer = intStart To intEnd
                strBits &= IIf(objBitArray.Item(i), "1", "0")
            Next
            Return strBits

        End Function
        Private Function BytesToString(ByVal Bytes() As Byte, Optional ByVal StartIndex As Integer = 0, Optional ByVal Length As Integer = 0) As String
            If (Length > 0) Then
                Return System.Text.Encoding.ASCII.GetString(Bytes, StartIndex, Length)
            Else
                Return System.Text.Encoding.ASCII.GetString(Bytes)
            End If
        End Function
    End Class
    Public Class ID3v1Tag
        ''' Declarations for friend variables holding the property values
        Friend m_blnTagAvailable As Boolean = False
        Friend m_bytTagVersion As ID3v1TagVersionEnum = ID3v1TagVersionEnum.Version10
        Friend m_strArtist As String = ""
        Friend m_strTitle As String = ""
        Friend m_strAlbum As String = ""
        Friend m_strYear As String = ""
        Friend m_strComment As String = ""
        Friend m_bytGenre As Byte = 0
        Friend m_strGenreString As String = ""
        Friend m_bytTrack As Byte = 0
        Public ReadOnly Property TagAvailable() As Boolean
            Get
                Return m_blnTagAvailable
            End Get
        End Property
        Public Property TagVersion() As ID3v1TagVersionEnum
            Get
                Return m_bytTagVersion
            End Get
            Set(ByVal Value As ID3v1TagVersionEnum)
                m_bytTagVersion = Value
                If (m_bytTagVersion = ID3v1TagVersionEnum.Version11 And Me.Comment.Length > 28) Then
                    Me.Comment = Me.Comment.Substring(1, 28)
                End If
            End Set
        End Property
        Public Property Title() As String
            Get
                Return m_strTitle
            End Get
            Set(ByVal Value As String)
                If (Value.Length > 30) Then
                    Throw New Exception("The length of the property 'Title' must be equal or less than 30 bytes.")
                Else
                    m_strTitle = Value
                End If
            End Set
        End Property
        Public Property Artist() As String
            Get
                Return m_strArtist
            End Get
            Set(ByVal Value As String)
                If (Value.Length > 30) Then
                    Throw New Exception("The length of the property 'Artist' must be equal or less than 30 bytes.")
                Else
                    m_strArtist = Value
                End If
            End Set
        End Property
        Public Property Album() As String
            Get
                Return m_strAlbum
            End Get
            Set(ByVal Value As String)
                If (Value.Length > 30) Then
                    Throw New Exception("The length of the property 'Album' must be equal or less than 30 bytes.")
                Else
                    m_strAlbum = Value
                End If
            End Set
        End Property
        Public Property Year() As String
            Get
                Return m_strYear
            End Get
            Set(ByVal Value As String)
                If (Value.Length > 4) Then
                    Throw New Exception("The length of the property 'Year' must be equal or less than 4 bytes.")
                Else
                    m_strYear = Value
                End If
            End Set
        End Property
        Public Property Track() As Byte
            Get
                Return m_bytTrack
            End Get
            Set(ByVal Value As Byte)
                m_bytTrack = Value
            End Set
        End Property
        Public Property Comment() As String
            Get
                Return m_strComment
            End Get
            Set(ByVal Value As String)
                Select Case m_bytTagVersion
                    Case ID3v1TagVersionEnum.Version10
                        If (Value.Length > 30) Then
                            Throw New Exception("The length of the property 'Comment' must be equal or less than 30 bytes for tag version 1.0.")
                        Else
                            m_strComment = Value
                        End If
                    Case ID3v1TagVersionEnum.Version11
                        If (Value.Length > 28) Then
                            Throw New Exception("The length of the property 'Comment' must be equal or less than 28 bytes for tag version 1.1.")
                        Else
                            m_strComment = Value
                        End If
                End Select
            End Set
        End Property
        Public Property Genre() As Byte
            Get
                Return m_bytGenre
            End Get
            Set(ByVal Value As Byte)
                m_bytGenre = Value
            End Set
        End Property
        Public Function GetGenreString(ByVal bytGenre As Byte) As String
            Try
                Dim strGenres() As String = {"Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", _
                "Hip - Hop", "Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", "R&B", "Rap", "Reggae", "Rock", "Techno", _
                "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", "Soundtrack", "Euro -Techno", "Ambient", _
                "Trip -Hop", "Vocal", "Jazz Funk", "Fusion", "Trance", "Classical", "Instrumental", "Acid", "House", "Game", _
                "Sound Clip", "Gospel", "Noise", "AlternRock", "Bass", "Soul", "Punk", "Space", "Meditative", _
                "Instrumental Pop", "Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno -Industrial", "Electronic", _
                "Pop -Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult", "Gangsta", "Top 40", "Christian Rap", _
                "Pop/Funk", "Jungle", "Native American", "Cabaret", "New Wave", "Psychadelic", "Rave", "Showtunes", "Trailer", _
                "Lo - Fi", "Tribal", "Acid Punk", "Acid Jazz", "Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", _
                "Folk", "Folk/Rock", "National Folk", "Swing", "Bebob", "Latin", "Revival", "Celtic", "Bluegrass", "Avantgarde", _
                "Gothic Rock", "Progressive Rock", "Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band", "Chorus", _
                "Easy Listening", "Acoustic", "Humour", "Speech", "Chanson", "Opera", "Chamber Music", "Sonata", "Symphony", _
                "Booty Bass", "Primus", "Porn Groove", "Satire", "Slow Jam", "Club", "Tango", "Samba", "Folklore", "Ballad", _
                "Power Ballad", "Rhythmic Soul", "Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Cappella", "Euro - House", _
                "Dance Hall", "Goa", "Drum & Bass", "Club - House", "Hardcore", "Terror", "Indie", "BritPop", "Negerpunk", _
                "Polsk Punk", "Beat", "Christian Gangsta Rap", "Heavy Metal", "Black Metal", "Crossover", "Contemporary Christian", _
                "Christian Rock", "Merengue", "Salsa", "Thrash Metal", "Anime", "JPop", "Synthpop"}
                Return strGenres(bytGenre)
            Catch ex As Exception
                Return bytGenre
            End Try
        End Function
    End Class
End Namespace
