Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "ms.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close
