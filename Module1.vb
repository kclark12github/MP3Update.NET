Module Module1
    Sub Main()
        Dim objMP3Info As New MP3.MP3Info()
        'objMP3Info.Filename = "C:\Documents and Settings\All Users\Documents\My Music\Rock\Collins, Phil\1985 - No Jacket Required\10 - Take Me Home.mp3"
        objMP3Info.Filename = "D:\Users\kfc12\Music\Rock\Collins, Phil\1985 - No Jacket Required\10 - Take Me Home.mp3"

        'Console.WriteLine("Filesize:{0}{1:#,##0.0#} KB", vbTab, objMP3Info.Filesize / 1024)
        'Console.WriteLine("Sampling:{0}{1:0.0} KHz", vbTab, objMP3Info.SamplingRateFrequency / 1000)
        'Console.WriteLine("Padding:{0}{1:#,##0} Bytes", vbTab, objMP3Info.Padding)
        'Console.WriteLine("Private:{0}{1}", vbTab, objMP3Info.PrivateBit)
        'Console.WriteLine("Copyright:{0}{1}", vbTab, objMP3Info.Copyright)
        'Console.WriteLine("OriginalBit:{0}", objMP3Info.OriginalBit)
        'Console.WriteLine("Bitrate:{0}{1} kbps", vbTab, objMP3Info.Bitrate \ 1000)
        'Console.WriteLine("FrameSamples:{0}", objMP3Info.FrameSamples)
        'Console.WriteLine("FrameSize:{0} Bytes", objMP3Info.FrameSize)
        'Console.WriteLine("Length:{0}{1}:{2}", vbTab, Int(objMP3Info.Length / 60), objMP3Info.Length Mod 60)

        'Console.WriteLine("MPEGType:{0}{1}/{2}", vbTab, objMP3Info.MPEGVersion.ToString, objMP3Info.Layer.ToString)
        'Console.WriteLine("Protection:{0}{1}", vbTab, objMP3Info.Protection)
        'Console.WriteLine("ChannelMode:{0}{1}", vbTab, objMP3Info.ChannelMode)
        'Console.WriteLine("Emphasis:{0}{1}", vbTab, objMP3Info.Emphasis)
        'Console.WriteLine("Encoding:{0}{1}", vbTab, objMP3Info.Encoding)

        If (objMP3Info.ID3v1Tag.TagAvailable) Then
            With objMP3Info
                Console.WriteLine($"{ .ID3v1Tag.Genre}\{ .ID3v1Tag.Artist}\{ .ID3v1Tag.Year} - { .ID3v1Tag.Album}\{ .ID3v1Tag.Track:00} - { .ID3v1Tag.Title}.mp3{vbTab} { .ChannelMode} {Int(.Length / 60):00}:{ .Length Mod 60:00} @ { .Bitrate \ 1000} kbps/{ .SamplingRateFrequency / 1000:0.0} KHz")
            End With
        End If

        'Update the tag
        'objMP3Info.ID3v1Tag.Title = "Another title"
        'objMP3Info.ID3v1Tag.Update()
        Console.ReadLine()
    End Sub
End Module
