speechtext=inputbox("Enter a text to speak","TEXT TO SPEECH-[Rockingheart]")
set objspeech=createobject("SAPI.spVoice")
objspeech.speak speechtext