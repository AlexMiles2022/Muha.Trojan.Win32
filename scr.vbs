Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "ms.mp3"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close
