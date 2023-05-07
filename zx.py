import win32com.client as win32
import speech_recognition as sr

listener=sr.Recognizer()

def take_command():
    try:
        with sr.Microphone() as source:
            print('listening...')
            voice = listener.listen(source)
            command = listener.recognize_google(voice)
            command = command.lower()
        
    except:
        pass
    return command

word = win32.Dispatch("Word.Application")
word.Visible = True

doc = word.Documents.Add()

user_input=take_command()

selection = word.Selection
selection.TypeText(user_input)

dn=input("Enter File Name :")
doc.SaveAs(dn+".docx")
doc.Close()

word.Quit()
