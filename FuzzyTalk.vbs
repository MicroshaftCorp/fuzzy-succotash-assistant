dim speechobject
set speechobject=createobject("sapi.spvoice")
speechobject.speak "Welcome User"
speechobject.speak "I Am Fuzzy And I have An TTS Program"
Dim message, sapi
 message=InputBox("What Do You Want Fuzzy To Say","Fuzzy TTS")
 Set sapi=CreateObject("sapi.spvoice")
 sapi.Speak message
