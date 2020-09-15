Dim Message, Speak
Message=InputBox("Ketik kata yang ingin di ucapkan","ILFEN")
Set Speak=CreateObject("sapi.spvoice")
Speak.Speak Message