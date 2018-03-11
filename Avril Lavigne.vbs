Set VObj = CreateObject("SAPI.SpVoice")
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("say",1)
strFileText = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing
MSG = strFileText
If MSG = "" Then WScript.quit: Else
with VObj
	Set .voice = .getvoices.item(2)
	.Rate = 1
    .Volume = 100
    .Speak MSG
end with
WScript.quit