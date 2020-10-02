Const SAFT48kHz16BitStereo = 39
Const SSFMCreateForWrite = 3 ' Creates file even if file exists and so destroys or overwrites the existing file

Dim oFileStream, oVoice

text = InputBox("Enter the text to be converted text","Somesh's Script")

Set oFileStream = CreateObject("SAPI.SpFileStream")
oFileStream.Format.Type = SAFT48kHz16BitStereo
oFileStream.Open "<your_desired_output_path>", SSFMCreateForWrite   'Add desired path here'

Set oVoice = CreateObject("SAPI.SpVoice")
Set oVoice.Voice = oVoice.GetVoices.Item(0)                          'Set parameter to 0 for male voice and 1 for female voice'
Set oVoice.AudioOutputStream = oFileStream
oVoice.Speak text

oFileStream.Close
