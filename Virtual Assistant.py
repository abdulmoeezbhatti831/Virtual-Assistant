from win32com.client import Dispatch
from colorama import init, Fore, Style, Back
import pywhatkit as kit
import speech_recognition as sr
import datetime
import os
import shutil
import sys
import time
import webbrowser
import wikipedia
import threading

init(autoreset=True)

def speak(argu: str):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(argu)


def loading(word: str, time_sec: int):
    spinner = ["|", "/", "-", "\\"]
    x = 0
    for i in range(time_sec * 10):
        if i == (time_sec * 10) - 1:
            sys.stdout.write("\r                                                      ")
            break
        if x == 4:
            x = 0
        sys.stdout.write(Fore.CYAN + f"\r{word}..." + spinner[x])
        sys.stdout.flush()
        time.sleep(0.1)
        x += 1


def welcome():
    hour = int(datetime.datetime.now().hour)
    if 4 <= hour < 12:
        speak("Good Morning!")
    elif 12 <= hour < 18:
        speak("Good Afternoon!")
    elif 18 <= hour < 22:
        speak("Good Evening!")
    else:
        speak("It's quite Night! Go to some Sleep! Good Night! And Have SWEET DREAMS!!")
        exit(0)
        
    name = input(Style.BRIGHT + Fore.BLUE + "Name for your Virtual Assistant\n(Default: AMB) > " + Style.RESET_ALL)
    if name == "": name = "AMB"

    speak(f"This is {name}, a Virtual Assistant designed for PERFROMING BASIC TASKS!")
    speak("What can I do for you?")


def user_command():
    r = sr.Recognizer()
    r.pause_threshold = 2
    with sr.Microphone() as source:
        print(Fore.BLUE + Style.BRIGHT + "\n🎤 Say to your Virtual Assistant...")
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
    loading("🔎 Recognizing your voice", 7)
    user_said = r.recognize_google(audio, language="en-US")
    print(Style.BRIGHT + "\n🗣️ You Said: " + Style.NORMAL + f"{user_said}\n")
    time.sleep(2)
    return user_said


if __name__ == "__main__":
    welcome()

    print("""What can I do for you?
    → Open Camera
    → Open Google
    → Open YouTube
    → Google Search
    → YouTube Search
    → Open Website with URL
    → Current Date and Time
    → Wikipedia Search
    → Play music 
    → Whatsapp msg
    → Just Open Whatsapp
    → Send Email
    → Chat Bot
    → CMD Runner
    → Write a note
    → Folder or File Control
    → Some Exciting Programs (Ever I Create, Learn & Have! By the Coder!)
    → Say 'exit program' to close or stop""")

    time.sleep(5)
    speak("Press enter, To EXPLORE me further!")
    input()
        
        
    while True:
        try:
            query = user_command().lower()

            if all(aciton in query for aciton in ["open", "camera"]):
                speak("Opening Camera..., Now give a sweet look!")
                os.system("start microsoft.windows.camera:")
                
            elif all(aciton in query for aciton in ["open", "google"]):
                speak("Going to WORLD's number one SEARCH ENGINE 'GOOGLE'!!!")
                webbrowser.open("https://www.google.com")
                
            elif all(aciton in query for aciton in ["open", "youtube"]):
                speak("Lets explore the WORLD's Most popular Social Platform, YOUTUBE...")
                webbrowser.open("https://www.youtube.com")
                
            elif all(aciton in query for aciton in ["google", "search"]):
                speak("Type to take a look GOOGLE!")
                kit.search(input(Style.BRIGHT + Fore.CYAN + "Search on Google Here: " + Style.RESET_ALL))
                
            elif all(aciton in query for aciton in ["youtube", "search"]):
                speak("Just type and get a video on it on YOUTUBE!")
                webbrowser.open(f"https://www.youtube.com/results?search_query={input(Style.BRIGHT + Fore.CYAN + "Search on YouTube Here: " + Style.RESET_ALL).replace(" ", "+")}")
                
            elif all(aciton in query for aciton in ["website", "url"]):
                speak("Give me the URL and I'll search upon it...")
                webbrowser.open(input(Style.BRIGHT + Fore.CYAN + "Website URL: " + Style.RESET_ALL))
                
            elif any(aciton in query for aciton in ["date", "time"]):
                speak("Okay, I also wanted to know the Date and Time...")
                print(Fore.LIGHTGREEN_EX + datetime.datetime.now().strftime("Current Date and Time:\n\nDate: %B %d, %Y.\nDay: %A\nTime: %I:%M:%S %p"))
            
            elif all(aciton in query for aciton in ["wikipedia", "search"]):
                speak("Get information on anything here. Just tell me the TOPIC...")
                print(wikipedia.summary(input(Style.BRIGHT + Fore.CYAN + "Topic: " + Style.RESET_ALL), sentences=input(Fore.CYAN + "No. of sentences: " + Style.RESET_ALL)))
                
            elif all(aciton in query for aciton in ["cmd", "run"]):
                speak("CMD initiating...")
                os.system(input(Style.BRIGHT + Fore.CYAN + "Command: " + Style.RESET_ALL))
                
            elif all(aciton in query for aciton in ["write", "note"]):
                speak("Making you ready the environment...")
                
                def write_note():
                    speak("Write or speak?")
                    user = input(Style.BRIGHT + Fore.CYAN + "> " + Style.RESET_ALL).lower()
                    if user == "speak":
                        r = sr.Recognizer()
                        r.pause_threshold = 3
                        time.sleep(5)
                        print(Fore.CYAN + Style.BRIGHT + "\nSay your message... 🎤 ")
                        with sr.Microphone() as source:
                            r.adjust_for_ambient_noise(source)
                            audio = r.listen(source)
                        loading("🔎 Recognizing", 5)
                        user_note = r.recognize_google(audio, language="en-US")
                        print(Style.BRIGHT + "\nYour message: " + Style.NORMAL + user_note)
                        return user_note
                    elif user == "write":
                        speak("Type your message. Write 'exit' at the last line to stop and get the complete message!")
                        print(Style.BRIGHT + "Message: ", end="")
                        user_note = []
                        while True:
                            line = input().lower()
                            if line == "exit":
                                break
                            user_note.append(line)
                        return "\n".join(user_note)
                    else:
                        speak(f"Sorry, {user} is not available!")
                        return write_note()
                
                while True: 
                    Path = input(Style.BRIGHT + Fore.CYAN + "Path: " + Style.RESET_ALL)
                    note = input(Fore.CYAN + "Create Text File: " + Style.RESET_ALL)
                    if not note.endswith(".txt"): note = f"{Path}\\{note}.txt"
                    else: note = f"{Path}\\{note}"

                    if not os.path.exists(note):
                        with open(note, "w") as f:
                            f.write(write_note())
                        speak("Task done!")
                        print(Fore.GREEN + "✅ Task Done!")
                        break
                    else:
                        speak("Such name already have used try another!")
                    
                        
            elif all(aciton in query for aciton in ["play", "music"]):
                speak("Hahahahahahahahah! Music Time!")
                speak("Frist Let's Create a songs list do want to play (Type 'exit' to exit)")
                
                print(Fore.MAGENTA + "🎵 Create the list: ")
                songs = []
                while True:
                    song = input(f"» ").lower()
                    if song == "exit":
                        break
                    songs.append(song)
                    
                found = False
                
                def search_animation():
                    x = 1
                    i = 0
                    while True:
                        if found:
                            break
                        if i % 3 == 0:
                            sys.stdout.write(Fore.BLUE + "\r🔎 Finding Your Songs    ")
                            x = 1
                        sys.stdout.write(Fore.BLUE + "\r🔎 Finding Your Songs" + "." * x)
                        sys.stdout.flush()
                        time.sleep(0.2)
                        x += 1
                        i += 1
                    
                for song in songs:
                    t = threading.Thread(target=search_animation)
                    t.start()
                    for drive in os.listdrives():
                        for root, dir, files in os.walk(drive):
                            for file in files:
                                if (song.lower() in file.lower() or song.lower().replace(" ", "_") in file.lower()) and file.endswith(".mp3"):
                                    found = True
                                    t.join()
                                    print(Style.BRIGHT + "\nSong Found: " + Back.LIGHTBLUE_EX + os.path.join(root, file))
                                    os.startfile(os.path.join(root, file))
                                    break
                            if found:
                                break
                        if found:
                            break
                    if not found:
                        found = True
                        t.join()
                        print(Fore.RED + f"\n❌ {song} Not Found!\n")
                    found = False
                print(Fore.GREEN + "✅ Enjoy!\n")

            
            elif all(aciton in query for aciton in ["just", "open", "whatsapp"]):
                speak("Make sure Your whatsapp is connected with WEB WHATSAPP!")
                kit.open_web()
                
            elif all(aciton in query for aciton in ["whatsapp", "message"]):
                speak("Oh YES! It's CHIT CHAT TIME!")
                
                def msg():
                    speak("Write or speak?")
                    user = input("> ").lower()
                    if user == "speak":
                        r = sr.Recognizer()
                        r.pause_threshold = 3
                        time.sleep(5)
                        print(Fore.CYAN + Style.BRIGHT + "\nSay your message... 🎤 ")
                        with sr.Microphone() as source:
                            r.adjust_for_ambient_noise(source)
                            audio = r.listen(source, phrase_time_limit=2)
                        message = r.recognize_google(audio, language="en-US")
                        print(Style.BRIGHT + "Your message: " + Style.NORMAL + message)
                        input()
                        return message
                    elif user == "write":
                        speak("Type your message. Write 'exit' at the last line to stop and get the complete message!")
                        print(Style.BRIGHT + "Message: ", end="")
                        message = []
                        while True:
                            line = input().lower()
                            if line == "exit":
                                break
                            message.append(line)
                        return "\n".join(message)
                    else:
                        speak(f"Sorry, {user} is not available!")
                        return msg()
                
                receiver = input(Style.BRIGHT + Fore.YELLOW + "Receiver Phone No. or Group ID: " + Style.RESET_ALL)
                if any(char.lower() in "abcdefghijklmnopqrstuvwxyz" for char in receiver):
                    kit.sendwhatmsg_to_group_instantly(receiver, msg(), tab_close=True)
                    print(Fore.GREEN + "✅ Task Done!")
                else:
                    kit.sendwhatmsg_instantly(receiver, msg(), tab_close=True)
                    print(Fore.GREEN + "✅ Task Done!")
                    
            elif all(aciton in query for aciton in ["send", "mail"]):
                speak("Make sure your message should well written! And one thing more here you need your email app pass not original password! If don't know search about it!")
                EMAIL_ID = input(Style.BRIGHT + Fore.CYAN + "Your Email ID: " + Style.RESET_ALL)
                EMAIL_PWD = input(Fore.CYAN + "Your Apps Password: " + Style.RESET_ALL)
                EMAIL_RECEIVER = input(Fore.CYAN + "Receiver Email ID: " + Style.RESET_ALL)
                kit.send_mail(EMAIL_ID, EMAIL_PWD, input(Style.BRIGHT + "Subject: " + Style.RESET_ALL), input(Style.BRIGHT + "Message: " + Style.RESET_ALL), EMAIL_RECEIVER)
                    
            elif all(aciton in query for aciton in ["chat", "bot"]):
                speak("Okay, go to do some chat with a real-time AI!")
                webbrowser.open("https://www.chatgpt.com")
                
            elif "control" in query and any(aciton in query for aciton in ["file", "folder"]):
                speak("Let's Control your PC from here! Tell me what to do to?")
                print(Fore.LIGHTGREEN_EX + Style.BRIGHT + "↪ Create\t↪ Delete\t↪ Move\t↪ Rename\t↪ Copy\t↪ Compress\t↪ List")
                
                OG_Path = os.getcwd()
                
                Path = input(Style.BRIGHT + Fore.GREEN + "Set the Path where to perform the action: " + Style.RESET_ALL)
                if os.path.exists(Path):
                    speak("PATH EXIST!! Access Granted!")
                    print(Fore.GREEN + "✅ Access Granted!")
                    os.chdir(Path)
                    
                    while True:
                        command = input(Style.BRIGHT + Fore.LIGHTGREEN_EX + "\nSo, tell me -> " + Style.RESET_ALL).lower()
                        
                        if command == "list":
                            for root, dir, files in os.walk(Path):
                                if files == []:
                                    print(Back.LIGHTBLUE_EX + Fore.WHITE + root)
                                for file in files:
                                    print(Back.LIGHTBLUE_EX + Fore.WHITE + os.path.join(root, file))
                            speak("Task done!")
                            print(Fore.GREEN + "✅ Task Done!")
                            
                        elif command == "create":
                            speak("This can also create folder in folder! Just give correct info!")
                            create = input(Fore.LIGHTCYAN_EX + "Folder or File or if sub-folder (Folder\\sub-folder\\File): " + Style.RESET_ALL)
                            if "\\" in create:
                                os.makedirs(create)
                            elif "." in create:
                                open(create, "w")
                            else:
                                os.mkdir(create)
                            speak("Task done!")
                            print(Fore.GREEN + "✅ Task Done!")
                                
                        elif command == "delete":
                            speak("Make sure to give right info or it may delete some important stuff!")
                            delete = input(Fore.LIGHTCYAN_EX + "Folder or File or if sub-folder (Folder\\sub-folder\\File): " + Style.RESET_ALL)
                            if "." in delete:
                                os.remove(delete)
                            else:
                                shutil.rmtree(delete)
                            speak("Task done!")
                            print(Fore.GREEN + "✅ Task Done!")
                                
                        elif command == "move":
                            speak("Make sure to give right info or it might moved some important stuff!")
                            move = input(Fore.LIGHTCYAN_EX + "Folder or File or if sub-folder (Folder\\sub-folder\\File): " + Style.RESET_ALL)
                            shutil.move(move, input(Fore.LIGHTCYAN_EX + "Destination with full path (C:\\): " + Style.RESET_ALL) + "\\" + move.split("\\")[-1])
                            speak("Task done!")
                            print(Fore.GREEN + "✅ Task Done!")
                            
                        elif command == "rename":
                            speak("Don't like the current name. It's okay, now rename it!")
                            rename = input(Fore.LIGHTCYAN_EX + "Folder or File or if sub-folder (Folder\\sub-folder\\File): " + Style.RESET_ALL)
                            os.rename(rename, input(Fore.YELLOW + "New name: " + Style.RESET_ALL))
                            speak("Task done!")
                            print(Fore.GREEN + "✅ Task Done!")
                            
                        elif command == "copy":
                            speak("Just make sure don't you put the copy file into original file path!")
                            copy = input(Fore.LIGHTCYAN_EX + "Folder or File or if sub-folder (Folder\\sub-folder\\File): " + Style.RESET_ALL)
                            if "." in copy:
                                shutil.copy(copy, input(Fore.LIGHTCYAN_EX + "Destination with full path (C:\\): " + Style.RESET_ALL) + "\\" + copy.split("\\")[-1])
                            else:
                                shutil.copytree(copy, input(Fore.LIGHTCYAN_EX + "Destination with full path (C:\\): " + Style.RESET_ALL) + "\\" + copy.split("\\")[-1])
                            speak("Task done!")
                            print(Fore.GREEN + "✅ Task Done!")
                                
                        elif command == "compress":
                            speak("That's a good idea to compress the big stuff! Make sure it couldn't be performed on files")
                            compress = input(Fore.LIGHTCYAN_EX + "Folder or if sub-folder (Folder\\sub-folder): " + Style.RESET_ALL)
                            if "." in compress:
                                print(Fore.RED + "\n❌ Operation could not be Done!")
                            else:
                                shutil.make_archive(compress, "zip", compress)
                                speak("Task done!")
                                print(Fore.GREEN + "✅ Task Done!")
                        
                        else:
                            speak(f"Sorry, {command} is not available!")
                            continue

                        if input(Fore.BLUE + Style.BRIGHT + "More actions in this Path (y/n)? > " + Style.RESET_ALL).lower() == "y":
                            continue
                        else:
                            break
                        
                    os.chdir(OG_Path)
                        
                else:
                    speak("Sorry, I can't perform the action as the path you have provided doesn't exit!")
                    print(Fore.RED + "❌ Path doesn't exists!")
                    
            elif all(aciton in query for aciton in ["exciting", "program"]):
                speak("Let me show you, the exciting programs, ever created, learnt, or had, by the CODER, of this CODE!!")
                webbrowser.open("https://www.github.com/abdulmoeezbhatti831")
                
            elif all(aciton in query for aciton in ["exit", "program"]):
                speak("WELL, Thanks for using me as your Virtual Assistant!!")
                speak("NOW Exiting... In")
                for i in range(1, 4):
                    speak(4-i)
                print(Fore.RED + "❌ EXIT!")
                speak("EXIT")
                break
                
            else:
                speak("I have mentioned my capabilities. So, please ask the stuff related to it! Thanks...")
                print("""What can I do for you?
                → Open Camera
                → Open Google
                → Open YouTube
                → Google Search
                → YouTube Search
                → Open Website with URL
                → Current Date and Time
                → Wikipedia Search
                → Play music 
                → Whatsapp msg
                → Just Open Whatsapp
                → Send Email
                → Chat Bot
                → CMD Runner
                → Write a note
                → Folder or File Control
                → Some Exciting Programs (Ever I Create, Learn & Have! By the Coder!)
                → Say 'exit program' to close or stop""")
                
        except sr.UnknownValueError:
            speak("Hmmmmm, I got you, you didn't say anything!!")
            
        except Exception as e:
            speak("It looks like we are facing an ERROR!")
            print(Fore.MAGENTA + Style.BRIGHT + "\n[Error]: " + Style.NORMAL + e)
            
        input(Style.DIM + "\nPress any key to continue...\n" + Style.RESET_ALL)