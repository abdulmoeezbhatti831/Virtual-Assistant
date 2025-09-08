from win32com.client import Dispatch # pip install pywin32
from colorama import init, Fore, Style, Back 
from openai import OpenAI
import pywhatkit as kit
import speech_recognition as sr # pip install speechRecognition
import datetime
import os
import shutil
import sys
import time
import webbrowser
import wikipedia
import threading
import random

# Autoreset the console color after each print function
init(autoreset=True)

# PC voice / speaking by PC
def speak(argu: str):
    global name
    speak = Dispatch("SAPI.SpVoice")
    if name == "friday":
        speak.Voice = speak.GetVoices()[1]
        speak.Speak(argu)
    else:
        speak.Speak(argu)

# Loading Animation
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

# Welcome Screen
def welcome():
    global name
    name = input(Style.BRIGHT + Fore.BLUE + "Name for your Virtual Assistant (Jarvis / Friday) > " + Style.RESET_ALL).lower()
    # Automatic name selection if user didn't enter or give other then given ones
    if name == "" or not any(name == voice for voice in ["jarvis", "friday"]): name = random.choice(["jarvis", "friday"])
    
    # Welcome acccording to current time
    hour = int(datetime.datetime.now().hour)
    if 4 <= hour < 12: 
        speak("Good Morning!")
    elif 12 <= hour < 18:
        speak("Good Afternoon!")
    elif 18 <= hour < 22:
        speak("Good Evening!")
    else: # If it's quite night then program exit as Parents said to sleep quickly
        speak("It's quite late at night! You should get some rest. Good night and have sweet dreams! Take care!")
        exit(0)

    speak(f"Hello! I am {name.title()}, your Virtual Assistant, ready to help you with a variety of tasks.")
    speak("Just tell me what you need, and I'll do my best to assist you!")

# Use voice command to take input from the user
def user_command():
    global name
    r = sr.Recognizer()
    r.pause_threshold = 2
    with sr.Microphone() as source:
        print(Fore.BLUE + Style.BRIGHT + f"\n🎤 Say to {name.title()}...")
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
        loading("🔎 Recognizing your voice", 5)
        user_said = r.recognize_google(audio, language="en-in")
        print(Style.BRIGHT + "\n🗣️ You Said: " + Style.NORMAL + f"{user_said}\n")
        time.sleep(2)
        return user_said

# Main Body
if __name__ == "__main__":
    name = "Jarvis"
    welcome()

    # Menu
    print(Style.BRIGHT + Fore.CYAN + "\n✨ Here are some things I can do for you:")
    print(      Fore.YELLOW + """
    ┌─────────────────────────────────────────────────────────────┐
    │  📷  Open Camera                                            │
    │  🌐  Open Google                                            │
    │  📺  Open YouTube                                           │
    │  🔎  Google Search                                          │
    │  🎬  YouTube Search                                         │
    │  🌍  Open Website with URL                                  │
    │  📅  Current Date and Time                                  │
    │  📚  Wikipedia Search                                       │
    │  🎶  Play music (anywhere in your PC)                       │
    │  💬  Whatsapp msg                                           │
    │  🟢  Just Open Whatsapp                                     │
    │  ✉️   Send Email                                             │
    │  🤖  Simple Chatting (with real-time AI)                    │
    │  💻  CMD Runner                                             │
    │  📝  Write a note                                           │
    │  📁  Folder or File Control                                 │
    │  🚀  Some Exciting Programs (Created & Learnt by the Coder) │
    │  ❌  Say 'exit program' to close or stop                    │
    └─────────────────────────────────────────────────────────────┘
                """)
    print(Style.RESET_ALL)

    # Gap for user to read menu easily then started
    loading("⌚ Wait a while", 5)
    speak("Press Enter, To EXPLORE me further!")
    input()
        
    # While Loop to give infinite command
    while True:
        try:
            # Taking the user input as query
            query = user_command().lower()

            # Open Camera
            if all(aciton in query for aciton in ["open", "camera"]):
                speak("Opening Camera..., Now give a sweet look!")
                os.system("start microsoft.windows.camera:")
                
            # Open Google
            elif all(aciton in query for aciton in ["open", "google"]):
                speak("Going to WORLD's number one SEARCH ENGINE 'GOOGLE'!!!")
                webbrowser.open("https://www.google.com")
                
            # Open Youtube
            elif all(aciton in query for aciton in ["open", "youtube"]):
                speak("Lets explore the WORLD's Most popular Social Platform, YOUTUBE...")
                webbrowser.open("https://www.youtube.com")
                
            # Google Search
            elif all(aciton in query for aciton in ["google", "search"]):
                speak("Type to take a look GOOGLE!")
                kit.search(input(Style.BRIGHT + Fore.CYAN + "Search on Google Here: " + Style.RESET_ALL))
                
            # Youtube Search
            elif all(aciton in query for aciton in ["youtube", "search"]):
                speak("Just type and get a video on it on YOUTUBE!")
                webbrowser.open(f"https://www.youtube.com/results?search_query={input(Style.BRIGHT + Fore.CYAN + "Search on YouTube Here: " + Style.RESET_ALL).replace(" ", "+")}")
                
            # Open Website through URL
            elif all(aciton in query for aciton in ["website", "url"]):
                speak("Give me the URL and I'll search upon it...")
                webbrowser.open(input(Style.BRIGHT + Fore.CYAN + "Website URL: " + Style.RESET_ALL))
                
            # Date and Time
            elif any(aciton in query for aciton in ["date", "time"]):
                speak("Okay, I also wanted to know the Date and Time...")
                print(Fore.LIGHTGREEN_EX + datetime.datetime.now().strftime("Current Date and Time:\n\nDate: %B %d, %Y.\nDay: %A\nTime: %I:%M:%S %p"))
            
            # Wikipedia Search
            elif all(aciton in query for aciton in ["wikipedia", "search"]):
                speak("Get information on anything here. Just type the TOPIC...")
                result = wikipedia.summary(input(Style.BRIGHT + Fore.CYAN + "Topic: " + Style.RESET_ALL), sentences=input(Fore.CYAN + "No. of sentences: " + Style.RESET_ALL))
                speak("Searched Complete!")
                print(Fore.GREEN + result)
                time.sleep(3)
                speak(result)
                
            # CMD Runner
            elif all(aciton in query for aciton in ["cmd", "run"]):
                speak("CMD initiating...")
                os.system(input(Style.BRIGHT + Fore.CYAN + "Command: " + Style.RESET_ALL))
                
            # Writing a note + saving 
            elif all(aciton in query for aciton in ["write", "note"]):
                speak("Making you ready the environment...")
                
                # Function for writing
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
                        user_note = r.recognize_google(audio, language="en-in")
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
                
                # Loop for if given path already exists
                while True: 
                    Path = input(Style.BRIGHT + Fore.CYAN + "Path: " + Style.RESET_ALL)
                    note = input(Fore.CYAN + "Create Text File: " + Style.RESET_ALL)
                    
                    # Making sure to create .txt format
                    if not note.endswith(".txt"): note = f"{Path}\\{note}.txt"
                    else: note = f"{Path}\\{note}"

                    # Verfying the given path
                    if not os.path.exists(note):
                        with open(note, "w") as f:
                            f.write(write_note())
                        speak("Task done!")
                        print(Fore.GREEN + "✅ Task Done!")
                        break
                    else:
                        speak("Such name already have used try another!")
                    
            # Play Music (.mp3) anywhere in your PC
            elif all(aciton in query for aciton in ["play", "music"]):
                speak("🎶 It's Music Time! Let's get the party started!")
                speak("First, let's create a list of songs you want to play. Type 'exit' at the end when you're done.")
                
                # Creating the list 
                print(Fore.MAGENTA + "🎵 Create the list: ")
                songs = []
                while True:
                    song = input(f"» ").lower()
                    if song == "exit":
                        break
                    songs.append(song)
                    
                found = False # Initially found is false
                
                # Finding Animation
                def search_animation():
                    x = 1
                    i = 0
                    while True:
                        if found: # break when song found
                            break
                        if i % 3 == 0:
                            sys.stdout.write(Fore.BLUE + "\r🔎 Finding Your Songs    ")
                            x = 1
                        sys.stdout.write(Fore.BLUE + "\r🔎 Finding Your Songs" + "." * x)
                        sys.stdout.flush()
                        time.sleep(0.2)
                        x += 1
                        i += 1
                    
                # Searching for each song in the list from entire PC
                for song in songs:
                    t = threading.Thread(target=search_animation)
                    t.start()
                    for drive in os.listdrives():
                        for root, dir, files in os.walk(drive):
                            for file in files:
                                if (song.lower() in file.lower() or song.lower().replace(" ", "_") in file.lower()) and file.endswith(".mp3"):
                                    found = True # If found then play and break
                                    t.join()
                                    print(Style.BRIGHT + "\nSong Found: " + Back.LIGHTBLUE_EX + os.path.join(root, file))
                                    os.startfile(os.path.join(root, file))
                                    break
                            if found:
                                break
                        if found:
                            break
                    if not found: # if not found then will go for the next one
                        found = True # true for breaking the search animation
                        t.join()
                        print(Fore.RED + f"\n❌ {song} Not Found!\n")
                    found = False # Reset to false
                print(Fore.GREEN + "✅ Enjoy!\n")

            # Just Open Whatsapp -> Done through web whatsapp
            elif all(aciton in query for aciton in ["just", "open", "whatsapp"]):
                speak("Just ensure your WhatsApp is connected to WhatsApp Web in your browser. Opening WhatsApp Web for you now!")
                kit.open_web()
                
            # Whatsapp automated messaging -> Make sure to not use it much. As to avoid from spamming which can result suspend account
            elif all(aciton in query for aciton in ["whatsapp", "message"]):
                speak("Oh YES! It's CHIT CHAT TIME! Let's send a WhatsApp message. Make sure your WhatsApp Web is connected.")
                
                # Taking messages
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
                        message = r.recognize_google(audio, language="en-in")
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
                
                # If receiver is a Group ID in which user is joined and able to send msg
                if any(char.lower() in "abcdefghijklmnopqrstuvwxyz" for char in receiver):
                    kit.sendwhatmsg_to_group_instantly(receiver, msg(), tab_close=True)
                    print(Fore.GREEN + "✅ Task Done!")
                # If receiver is a Phone no.
                else:
                    kit.sendwhatmsg_instantly(receiver, msg(), tab_close=True)
                    print(Fore.GREEN + "✅ Task Done!")
                    
            # Sending Email
            elif all(aciton in query for aciton in ["send", "mail"]):
                speak("Make sure your message should well written!")
                EMAIL_ID = input(Style.BRIGHT + Fore.CYAN + "Your Email ID: " + Style.RESET_ALL)
                EMAIL_PWD = input(Fore.CYAN + "Your Apps Password: " + Style.RESET_ALL) # Here you need your email account app pass not original password! If don't know search about it!
                EMAIL_RECEIVER = input(Fore.CYAN + "Receiver Email ID: " + Style.RESET_ALL)
                kit.send_mail(EMAIL_ID, EMAIL_PWD, input(Style.BRIGHT + "Subject: " + Style.RESET_ALL), input(Style.BRIGHT + "Message: " + Style.RESET_ALL), EMAIL_RECEIVER)
                    
            # Simple Chatting -> Real Time AI chatting
            elif all(aciton in query for aciton in ["simple", "chat"]):
                speak("To start chatting, please provide your OpenAI API KEY.")
                API_KEY = input(Style.BRIGHT + Fore.CYAN + "Enter your OpenAI API KEY: " + Style.RESET_ALL)
                
                # Verifying the API KEY
                try:
                    client = OpenAI(api_key=API_KEY)
                    Verify_response = client.responses.create(
                        model = "gpt-5-mini",
                        input = "Nothing"
                    )
                except Exception as e:
                    speak("It looks like there's some error with your API KEY!")
                    print(Fore.MAGENTA + Style.BRIGHT + "\n[Error]: " + Style.NORMAL + str(e))
                    
                    input(Style.DIM + "\nPress Enter to continue..." + Style.RESET_ALL) # Gap between user inputs
                    continue
                
                # Function to get response from AI
                def AI(user_input):
                    client = OpenAI(api_key=API_KEY)
                    response = client.responses.create(
                        model = "gpt-5-mini",
                        input = user_input
                    )
                    return response.output_text
                
                # Separate speech function for just chatting with AI
                def speech(): 
                    global name
                    r = sr.Recognizer()
                    r.pause_threshold = 2
                    with sr.Microphone() as source:
                        print(Fore.BLUE + Style.BRIGHT + f"\n🎤 Say to chat with {name.title()}...")
                        r.adjust_for_ambient_noise(source)
                        audio = r.listen(source)
                        loading("🔎 Recognizing your voice", 5)
                        user_said = r.recognize_google(audio, language="en-in")
                        return user_said
                    
                # Setting the Parameters
                chat = ""
                conversation_done = False
                speak("Start chatting! Say something to Start Chatting. You can say 'exit' to stop chatting.")
                
                # Loop to chat continuously
                while True:
                    try:
                        user_input = speech()
                        chat += f"You: {user_input}"
                        print(Fore.YELLOW + Style.BRIGHT + "\nYou: " + Style.RESET_ALL + user_input)
                        
                        # Handling the chat leaving stiuation
                        if user_input.lower() == "exit":
                            speak("Store your conversation as a memory!")
                            store = input(Fore.BLUE + Style.BRIGHT + "Store chat? (y/n) > " + Style.RESET_ALL).lower()
                            if store == "y":
                                speak("That's great to store your chatting as memory!")
                                while True:
                                    path = input(Fore.BLUE + Style.BRIGHT + "Path: " + Style.RESET_ALL)
                                    if os.path.exists(path):
                                        speak("PATH EXIST!! Access Granted!")
                                        with open(f"{path}\\Chat - {datetime.datetime.now().strftime("%Y-%m-%d %I.%M.%S %p")}.txt", "w") as f:
                                            f.write(chat)
                                        speak("Your Conversation have been saved Successfully!")
                                        conversation_done = True
                                        print(Fore.GREEN + Style.BRIGHT + "\n✅ Chat Saved!")
                                        break
                                    else:
                                        speak("Sorry, I can't perform the action as the path you have provided doesn't exit!")
                                        print(Fore.RED + "❌ Path doesn't exists!")
                            else:
                                conversation_done = True
                                speak("It's good to have a chat with you!")
                            
                        # If done conversation then break
                        if conversation_done:
                            break
                        
                        # Giving + showing + speaking the chat with AI
                        AI_reply = AI(user_input)
                        chat += f"\n{name.title()}: {AI_reply}\n"
                        print(Fore.CYAN + Style.BRIGHT + f"\n{name.title()}: " + Style.RESET_ALL + f"{AI_reply}\n")
                        speak(AI_reply)
                    
                    # Handling "if user said nothing" error
                    except sr.UnknownValueError:
                        speak("Sorry, I didn't understand you properly! Please press enter and say again!")
                        input(Style.DIM + "\nPress Enter to continue..." + Style.RESET_ALL)
                
            # Controlling Files and Folder in your PC
            elif "control" in query and any(aciton in query for aciton in ["file", "folder"]):
                speak("Ready to manage your files and folders! Let's Control your system!")
                print(Fore.LIGHTYELLOW_EX + "↪ Create  ↪ Delete  ↪ Move  ↪ Rename  ↪ Copy  ↪ Compress  ↪ List")
                
                # Current Directory Path
                OG_Path = os.getcwd()
                
                # Setting the Path in which actions are to be perfromed
                speak("Make sure to set the path First!")
                Path = input(Style.BRIGHT + Fore.GREEN + "Set the Path where to perform the action: " + Style.RESET_ALL)
                
                # if Path exists
                if os.path.exists(Path):
                    speak("PATH EXIST!! Access Granted!")
                    print(Fore.GREEN + "✅ Access Granted!")
                    
                    # Changing current directory to this path to perform actions more easily
                    os.chdir(Path)
                    
                    # Loop for many commands one after one
                    while True:
                        command = input(Fore.YELLOW + "\nSo, tell me what to do -> " + Style.RESET_ALL).lower()
                        
                        try:
                            if command == "list":
                                for root, dir, files in os.walk(Path):
                                    if files == []:
                                        print(Back.LIGHTBLUE_EX + Style.BRIGHT + root)
                                    for file in files:
                                        print(Back.LIGHTBLUE_EX + Style.BRIGHT + os.path.join(root, file))
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
                                speak("Make sure to give right info, or it may delete some important stuff!")
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
                                os.rename(rename, input(Fore.LIGHTCYAN_EX + "New name: " + Style.RESET_ALL))
                                speak("Task done!")
                                print(Fore.GREEN + "✅ Task Done!")
                                
                            elif command == "copy":
                                speak("Just make sure don't you put the copy file into original file path!")
                                copy = input(Fore.LIGHTCYAN_EX + "Folder or File or if sub-folder (Folder\\sub-folder\\File): " + Style.RESET_ALL)
                                if "." in copy:
                                    shutil.copy(copy, input(Fore.LIGHTCYAN_EX + "Destination with full path (C:\\): " + Style.RESET_ALL) + "\\" + "(Copy) " + (copy.split("\\")[-1] if True else copy))
                                else:
                                    shutil.copytree(copy, input(Fore.LIGHTCYAN_EX + "Destination with full path (C:\\): " + Style.RESET_ALL) + "\\" + "(Copy) " + (copy.split("\\")[-1] if True else copy))
                                speak("Task done!")
                                print(Fore.GREEN + "✅ Task Done!")
                                    
                            elif command == "compress":
                                speak("That's a good idea to compress the big stuff! Make sure it couldn't be performed on files")
                                compress = input(Fore.LIGHTCYAN_EX + "Folder or if sub-folder (Folder\\sub-folder): " + Style.RESET_ALL)
                                if "." in compress:
                                    speak("Task could not be done!")
                                    print(Fore.RED + "\n❌ Task could not be Done!")
                                else:
                                    shutil.make_archive(compress, "zip", compress)
                                    speak("Task done!")
                                    print(Fore.GREEN + "✅ Task Done!")
                            
                            else:
                                speak(f"Sorry, {command} is not available!")
                                continue
                        
                        # Handling any error while controlling files and folders
                        except Exception as e:
                            speak("Path File or Folder ERROR!")
                            print(Fore.MAGENTA + Style.BRIGHT + "\n[Error]: " + Style.NORMAL + str(e))

                        # Asking for more actions or done actions
                        if input(Fore.BLUE + Style.BRIGHT + "\nMore actions in this Path (y/n)? > " + Style.RESET_ALL).lower() == "y":
                            continue
                        else:
                            break
                        
                    # Get back to current working directory
                    os.chdir(OG_Path)
                        
                # If path not exists
                else:
                    speak("Sorry, I can't perform the action as the path you have provided doesn't exit!")
                    print(Fore.RED + "❌ Path doesn't exists!")
                    
            # Exciting Programs of the Coder
            elif all(aciton in query for aciton in ["excit", "program"]):
                speak("Let me show you, the exciting programs, ever created, learnt, or had, by MY CODER!!")
                webbrowser.open("https://www.github.com/abdulmoeezbhatti831")
                
            # For Exiting the program or this Virtual Assistant
            elif all(aciton in query for aciton in ["exit", "program"]):
                speak("Thank you for choosing me as your Virtual Assistant! I hope I was able to help you today. Have a wonderful day ahead!")
                speak("Exiting now. In:")
                for i in range(3, 0, -1):
                    speak(i)
                speak("Goodbye!")
                print(Fore.RED + "❌ Exit!")
                break
                
            # if query doesn't match any of all the above ones
            else:
                speak("Sorry, I couldn't recognize your command. Please try again with one of the listed options or rephrase your request.")
                
                # Again dislaying the Menu
                print(Style.BRIGHT + Fore.CYAN + "\n✨ Here are some things I can do for you:")
                print(Fore.YELLOW + """
    ┌─────────────────────────────────────────────────────────────┐
    │  📷  Open Camera                                            │
    │  🌐  Open Google                                            │
    │  📺  Open YouTube                                           │
    │  🔎  Google Search                                          │
    │  🎬  YouTube Search                                         │
    │  🌍  Open Website with URL                                  │
    │  📅  Current Date and Time                                  │
    │  📚  Wikipedia Search                                       │
    │  🎶  Play music (anywhere in your PC)                       │
    │  💬  Whatsapp msg                                           │
    │  🟢  Just Open Whatsapp                                     │
    │  ✉️   Send Email                                             │
    │  🤖  Simple Chatting (with real-time AI)                    │
    │  💻  CMD Runner                                             │
    │  📝  Write a note                                           │
    │  📁  Folder or File Control                                 │
    │  🚀  Some Exciting Programs (Created & Learnt by the Coder) │
    │  ❌  Say 'exit program' to close or stop                    │
    └─────────────────────────────────────────────────────────────┘
                """)
                print(Style.RESET_ALL)
                
        # Handling "if user said nothing" error
        except sr.UnknownValueError:
            speak("Hmmmmm, I got you, you didn't say anything!!")
            
        # Handling any other remaining error 
        except Exception as e:
            speak("It looks like we are facing an ERROR!")
            print(Fore.MAGENTA + Style.BRIGHT + "\n[Error]: " + Style.NORMAL + str(e))
            
        # Gap to next query / for not running continuously / if user leaves the program ON
        input(Style.DIM + "\nPress Enter to continue..." + Style.RESET_ALL)