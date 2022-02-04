Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "ukr.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close
