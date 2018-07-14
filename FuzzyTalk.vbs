@echo off
title Fuzzy Succotash
 ________________________/ O  \___/
 <_O_O_O_O_O_A_O_S_O_O_O_P_____/   \        
dim speechobject
set speechobject=createobject("sapi.spvoice")
speechobject.speak "Welcome User"
speechobject.speak "I Am Fuzzy And I have An TTS Program"
Dim message, sapi
 message=InputBox("What Do You Want Fuzzy To Say","Fuzzy TTS")
 Set sapi=CreateObject("sapi.spvoice")
 sapi.Speak message
