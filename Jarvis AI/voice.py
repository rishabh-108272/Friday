import win32com.client

your_voice_index=1
speaker = win32com.client.Dispatch("SAPI.SpVoice")
voices = speaker.GetVoices()
speaker.Voice = speaker.GetVoices().Item(your_voice_index)
for voice in voices:
    print(voice.GetDescription())
    speaker.speak(voice.GetDescription())


