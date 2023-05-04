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

# Creating a new doc
doc = word.Documents.Add()

# Taking Input Via Voice
user_input=take_command()

# Adding Input In Doc
selection = word.Selection
selection.TypeText(user_input)

# Saving The Doc
dn=input("Enter File Name :")
doc.SaveAs(dn+".docx")
doc.Close()

# Quit Microsoft Word
word.Quit()
