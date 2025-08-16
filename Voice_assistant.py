import tkinter as tk  # GUI
from tkinter import font, ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading  # for voice listening in background so that gui doesnt freeze
import queue  # pass message b/w background threads and GUI
import os  # to intreact python with your windows
import time
import webbrowser  # to open websites in default browsers
import subprocess  # use to run your system programs with more control than os
import speech_recognition as sr  # voice command recognization
import pyttsx3  # converts text to speech
import pyautogui  # screenshots
import requests  # get and provode news headlines
import xml.etree.ElementTree as et
from datetime import datetime
import platform  # system info (CPU,RAM ,Battery)
import re  # detect commands
import sys  # controls the python enviorment

# try/except feature checks:
try:
    import psutil
    has_psutil = True
except Exception:
    has_psutil = False

try:
    has_pyautogui = True
except Exception:
    has_pyautogui = False

# to take screenshots
try:
    from PIL import ImageGrab
    has_pyautogui = True
    has_imagegrab = True
except Exception:
    has_pyautogui = False
    has_imagegrab = False

# youtube first result scraping (beautiful Soap)
try:
    from bs4 import BeautifulSoup
    has_bs4 = True
except Exception:
    has_bs4 = False

# configuration:
app_paths = {  # add paths acc to your operating system
    "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "calculator": r"C:\Windows\System32\calc.exe",
    # Add more explicit app paths here if needed
}
spotify_path = r"C:\Users\Hasnain Memon\AppData\Roaming\Spotify\Spotify.exe"
news = "https://news.google.com/rss"

# checks screenshot folder exists
# os,getcwd: gets the current working directory , os.path.join: safely join folders name to create a full path
ss_folder = os.path.join(os.getcwd(), "Screenshots_Folder")
os.makedirs(ss_folder, exist_ok=True)

# text to speech:
tts_engine = pyttsx3.init()  # start text to speech engine
tts_engine.setProperty("rate", 150)


def speak(text):
    """Speak text without blocking UI."""
    def _s(t):
        try:
            print(f"[Assistant Speaking]: {t}")
            tts_engine.say(t)
            tts_engine.runAndWait()
        except Exception as e:
            print("Text to speech Error:", e)
            ui_log(f"TTS Error: {e}", "assistant")

    # without threading program will be pause until your speech finishes u cant control any button in the progarm
    threading.Thread(target=_s, args=(text,), daemon=True).start()


# UI: it will add msg to the queue to that GUI can display them later
ui_queue = queue.Queue()


def ui_log(msg, kind="info"):
    ui_queue.put((msg, kind))


# Command Functions:
def play_spotify(song_name):
    if not song_name:
        speak("Please tell me the song name")
        ui_log("No Song Specified:", "assistant")
        return
    encoded = song_name.strip().replace(" ", "%20")
    # song_name.strip(): remove extra space in the start or end ,
    # .replace(" ","%20"):Converts spaces into %20 so that the name is uri friendly.
    uri = f"spotify:search:{encoded}"
    ui_log(f"Launching Spotify search: {song_name}", "assistant")

    try:
        os.system(f'start {uri}')  # open through uri if registered
        # checks if spotify path is given so launch through path
        if os.path.exists(spotify_path):
            subprocess.Popen(spotify_path)
        speak(f"playing {song_name} on spotify")
    except FileNotFoundError:
        ui_log("Error: spotify not found please update your path! ", "assistant")
        speak("I couldn't find spotify on this system.")
    except Exception as e:
        ui_log(f"Spotify open error: {e}", "assistant")
        speak("I couldn't open Spotify.")


def search_google(query):
    if not query:
        speak("Please tell me what to search.")
        ui_log("No search query.", "assistant")
        return
    ui_log(f"Searching Google for: {query}", "assistant")
    speak(f"Searching Google for {query}")
    webbrowser.open(f"https://www.google.com/search?q={query}")


def read_news():
    ui_log("Fetching latest news...", "assistant")
    speak("Fetching the latest headlines.")
    try:
        r = requests.get(news, timeout=10)  # if server doesnt respond withi 10 secs, it will stop trying
        r.raise_for_status()  # if server returns an error it raises an exception
        root = et.fromstring(r.content)
        items = root.findall("./channel/item")
        if not items:
            ui_log("No news item found: ", "assistant")
            speak("I couldn't find news.")
            return
        # loop to take top 5 headlines
        for i, item in enumerate(items[:5], start=1):
            title = item.find("title").text  # extract the headlines text
            ui_log(f"News {i}: {title}", "assistant")
            speak(title)
            # small time delay between headines so that it can be understandable
            time.sleep(3)
    except Exception as e:
        ui_log(f"News fetch Error! {e}", "assistant")
        speak("Unable to fetch news right now.")


def open_app(app_name):
    name = app_name.strip().lower()  # remove extra space .strip() and convert it into lowercase .lower()
    if not name:
        speak("Which application should I open")
        ui_log("App name is missing. ", "assistant")
        return
    if name in app_paths:
        path = app_paths[name]
        try:
            ui_log(f"Openning {name}...", "assistant")
            speak(f"Openning {name}")
            os.startfile(path)  # used for direct app launching through path
            return
        except Exception:
            ui_log(f"Sorry: Failed to open {name}.", "assistant")
    try:
        ui_log(f"Attempting to launch {name}...", "assistant")
        if name in ("chrome", "google chrome") and os.path.exists(app_paths.get("chrome", "")):
            subprocess.Popen([app_paths["chrome"]])
        else:
            subprocess.Popen(name)
    except Exception as e:
        ui_log(f"could not open {name}: {e}", "assistant")
        speak(f"Unable to open app {name}. please update the app path")


def search_yt(query, play_first_video=False):
    if not query:
        speak("Please tell me what to search on youtube")
        return
    ui_log(f"Searching youtube for {query}...", "assistant")
    speak(f"Searching youtube for {query}")
    search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
    if play_first_video:
        try:
            r = requests.get(search_url, timeout=10)
            # only continues to search first video if the respose is ok
            r.raise_for_status()
            html = r.text  # contains entire html code of that page
            video_id = None
            if has_bs4:  # if beautifulsoap lib is installed
                soup = BeautifulSoup(html, "html.parser")
                a = soup.find("a", href=re.compile(r"^/watch"))
                # find first tag where an href look like /watch
                if a and a.get("href"):
                    video_id = a.get("href")
            else:
                # uses a regex(regular expression) to find yt first video
                b = re.search(r'href=\"(/watch\?v=[^\"]+)\"', html)
                if b:
                    video_id = b.group(1)
            if video_id:  # if video is found
                webbrowser.open(f"https://www.youtube.com{video_id}")
                return
        except Exception as e:
            ui_log(f"YouTube scrape error: {e}", "assistant")
    webbrowser.open(search_url)


def system_info():
    try:
        uname = platform.uname()  # platform.uname give system details
        # uses psutil to get CPU usage percentage over 10 sec
        cpu = f"{psutil.cpu_percent(interval=1)}%" if has_psutil else "psutil not installed"
        # get RAM usage
        ram = f"{psutil.virtual_memory().percent}%" if has_psutil else "psutil not installed"
        # get Disk usage
        disk = f"{psutil.disk_usage('/').percent}%" if has_psutil else "psutil not installed"
        battery = "N/A"
        if has_psutil:
            try:
                bat = psutil.sensors_battery()  # checks battery percenatge
                # if battery found show %
                battery = f"{bat.percent}%" if bat else "N/A"
            except Exception:
                # if battery is not shown on desktop then it will show N/A
                battery = "N/A"
        msg = (
            f"System: {uname.system} {uname.release}\n"  # uname.system = OS, uname.release = Version
            f"Machine: {uname.machine}\n"
            f"Processor: {uname.processor}\n"
            f"CPU Usage: {cpu}\n"
            f"RAM Usage: {ram}\n"
            f"Disk Usage: {disk}\n"
            f"Battery: {battery}"
        )
        ui_log(msg, "assistant")
        speak("Here is your system information. " + msg.replace("\n", ", "))    
        # show a messagebox from the main thread
        def _show():
            messagebox.showinfo("system Information", msg)
        # Schedule on the Tkinter event loop (thread-safe)
        ui_queue.put(("__SHOW_SYSINFO__", _show))
    except Exception:
        ui_log("Error in fetching system information", "assistant")
        speak("Sorry I couldn't get the system information.")


def screenshot():
    try:
        filename = os.path.join(
            ss_folder,
            f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"  # stores every ss with unique name
        )
        if has_pyautogui:
            ss = pyautogui.screenshot()
            ss.save(filename)
        else:
            ui_log("Screen shots features required pyautogui to run.", "assistant")
            speak("Screen shot is not avaliable please import pyautogui lib in your system")
            return

        ui_log(f"Screenshot saved: {filename}", "assistant")
        speak("Screenshot taken and saved.")
    except Exception as e:
        ui_log(f"Screenshot Error: {e}", "assistant")
        speak("Sorry, I couldn't take the screenshot.")


def sys_control(action):
    try:
        a = action.lower()
        if a == "lock":
            speak("Locking the computer")
            if platform.system().lower() == "windows":
                ctypes_cmd = "rundll32.exe user32.dll,LockWorkStation"  # windows lock attempt
                os.system(ctypes_cmd)
            else:
                # linux lock attempt
                os.system("gnome-screensaver-command -l || loginctl lock-session")

        elif a == "shutdown":
            speak("Shutting down the computer.")
            if platform.system().lower() == "windows":
                os.system("shutdown /s /t 1")
            else:
                os.system("shutdown now")

        elif a == "restart":
            speak("Restarting the computer.")
            if platform.system().lower() == "windows":
                os.system("shutdown /r /t 1")
            else:
                os.system("reboot")
        else:
            speak("Unknown system command.")

    except Exception as e:
        ui_log(f"System control error: {e}", "assistant")
        speak("I couldn't perform that action.")


def tell_date():
    today = datetime.now().strftime("%A, %B %d, %Y")
    speak(f"Today is {today}")


def tell_time():
    current_time = datetime.now().strftime("%I:%M %p")
    speak(f"The time is {current_time}")


# command processr (central_brain of the program)
def process_command(text):
    if not text:
        return
    # converts the command to lowercase and remove extra spaces
    command = text.strip().lower()
    ui_log(command, "user")

    # check if user said play or play music
    if command.startswith("play") or command.startswith("play music"):
        if "youtube" in command:  # if it contains yt in command
            query = command.replace("play youtube", "").strip()
            # searches first video in yt
            search_yt(query, play_first_video=True)
        else:
            song = command.replace("play music", "").replace("play", "", 1).strip()
            if song:  # play song on spotify
                ui_log(f"Playing {song} on Spotify...", "assistant")
                speak(f"Playing {song} on Spotify")
                play_spotify(song)
            else:
                # if no song is given ask song name
                speak("Which song should I play?")
                ui_log("No song name detected.", "assistant")

    elif command.startswith("search") or "search for " in command:
        if "youtube" in command:
            query = command.replace("search youtube", "").replace("youtube", "").strip()
            search_yt(query, play_first_video=False)
        else:
            if "search for " in command:
                # if search for is in command it will search
                query = command.split("search for", 1)[1].strip()
            else:
                # else if search is in command it will also search
                query = command.replace("search", "", 1).strip()
            search_google(query)

    elif "news" in command or "read news" in command:
        read_news()

    elif command.startswith("open"):
        app = command.split("open", 1)[1].strip()
        if app in ("chrome", "google chrome", "browser"):
            if os.path.exists(app_paths.get("chrome", "")):
                try:
                    os.startfile(app_paths["chrome"])
                except Exception:
                    webbrowser.open("https://www.google.com")
            else:
                webbrowser.open("https://www.google.com")
            speak("Opening Chrome")
        elif app in ("spotify", "open spotify"):
            play_spotify("")
            open_app("spotify")
            speak("Opening Spotify")
        elif app in ("calculator", "calc"):
            open_app("calculator")
            speak("Opening Calculator")
        else:
            if app:
                open_app(app)
            else:
                speak("please tell me the app to open")

    elif "time" in command:
        now = datetime.now().strftime("%I:%M %p")  # %I hours ,%M Mins , %p AM or PM
        ui_log(f"The time is {now}", "assistant")
        speak(f"The time is {now}")

    elif "date" in command:
        today = datetime.now().strftime("%B %d, %Y")  # %B month ,%d day %Y year
        ui_log(f"Today's date is {today}", "assistant")
        speak(f"Today's date is {today}")

    elif "youtube" in command:
        q = command.replace("youtube", "").replace("search", "").replace("play", "").strip()
        if q:  # if user say youtube search for --- it will directly open first video
            search_yt(q, play_first_video=True)
        else:  # if user says youtube it will open youtube browser
            webbrowser.open("https://www.youtube.com")
            speak("Openning youtube")

    elif "system info" in command or "system information" in command:
        system_info()

    elif "screenshot" in command or "take screenshot" in command:
        screenshot()

    # check the command is of one length ex lock correct lock computer wrong
    elif command == "lock":
        sys_control("lock")

    elif "shutdown" in command:
        sys_control("shutdown")

    elif "restart" in command:
        sys_control("restart")

    elif command in ("exit", "quit", "stop"):
        ui_log("shutting down the progarm....", "assistant")
        speak("Goodbye.")
        # assistant is sending special message to programs UI that its time to quit
        ui_queue.put(("__QUIT__", "info"))

    else:
        ui_log("Sorry, I don't know that command yet.", "assistant")
        speak("Sorry, I don't know that command yet.")


# Background listner:
listening_flag = threading.Event()
listen_thread = None
listen_lock = threading.Lock()


def listen_worker():
    """function runs in background and continously listens user command"""
    recognizer = sr.Recognizer()  # creates a recognizer object
    try:
        mic = sr.Microphone()
    except Exception as e:
        ui_log(f"Microphone error: {e}", "assistant")
        speak("Microphone not available.")
        return

    with mic as source:
        # listen to background noise for 1 sec
        recognizer.adjust_for_ambient_noise(source, duration=1)

    while listening_flag.is_set():  # keeps listening until flag is ON is_set is use to set status
        try:
            with mic as source:
                ui_log("Listening...", "assistant")
                # timeout it will w8 6 sec to listen you if u speak within 6 sec ok otherwise exception
                # phrase_time_limit : it will listen you until 12 sec after 12 sec it will stop listening.
                audio = recognizer.listen(source, timeout=6, phrase_time_limit=12)
            try:
                # take the voice as audio of user send it to the google speech API to get back text of your audio
                text = recognizer.recognize_google(audio)
            except sr.UnknownValueError:  # if google cant figure out audio it just skip it and go back to listening.
                continue
            except sr.RequestError as e:
                ui_log(f"Speech service error: {e}", "assistant")
                speak("Speech service error.")
                time.sleep(1)
                continue
            if text:
                process_command(text)
        except Exception as e:
            ui_log(f"Listener error: {e}", "assistant")
            time.sleep(0.5)


def start_listenning():
    global listen_thread
    if listening_flag.is_set():
        ui_log("Already Listening...", "assistant")
        return
    listening_flag.set()  # ON the listening_flag
    with listen_lock:
        listen_thread = threading.Thread(target=listen_worker, daemon=True)
        listen_thread.start()
    ui_log("Listening started.", "assistant")


def stop_listening():
    if listening_flag.is_set():
        listening_flag.clear()  # off the listening_flag
        ui_log("Stopped listening.", "assistant")
    else:
        ui_log("Was not listening.", "assistant")


# COLOR PALETTE 
BG_MAIN   = "#000000"   # Pure Black background
BG_PANEL  = "#0A0A1F"   # Dark Navy for panels
ACCENT    = "#4169E1"   # Royal Blue
TEXT      = "#FFFFFF"   # White text
SUBTEXT   = "#B0B3C1"   # Grayish secondary text

ui_queue = queue.Queue()
listening_flag = threading.Event()


#GUI CLASS:
class GUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Voice Assistant Console")
        self.geometry("850x580")
        self.configure(bg=BG_MAIN)
        self.resizable(False, False)

        # Fonts
        self.h1 = font.Font(family="Times New Roman", size=18, weight="bold")  # Headings
        self.h2 = font.Font(family="Calibri", size=12)                         # Normal text
        self.small = font.Font(family="Calibri", size=10)

        # Main Container 
        container = tk.Frame(self, bg=BG_PANEL, bd=2, relief="solid")
        container.place(relx=0.02, rely=0.03, relwidth=0.96, relheight=0.94)

        # Left Panel: Chat/Logs 
        left = tk.Frame(container, bg=BG_PANEL)
        left.place(relx=0.02, rely=0.03, relwidth=0.66, relheight=0.94)

        title = tk.Label(left, text="Voice Assistant", bg=BG_PANEL, fg=ACCENT, font=self.h1)
        title.pack(anchor="nw", padx=12, pady=(8, 2))

        self.chat_area = ScrolledText(
            left, bg=BG_MAIN, fg=TEXT, insertbackground=TEXT,
            font=self.h2, bd=0, padx=10, pady=10
        )
        self.chat_area.pack(fill="both", expand=True, padx=12, pady=6)
        self.chat_area.tag_config("user", foreground=ACCENT, justify="right")
        self.chat_area.tag_config("assistant", foreground=SUBTEXT, justify="left")
        self.chat_area.tag_config("center", justify="center")
        self.chat_area.config(state=tk.NORMAL)
        self.chat_area.insert(tk.END, "Welcome...\nHow can I help you??", "center")
        self.chat_area.config(state=tk.DISABLED)
        

        # Right Panel: Mic + Controls 
        right = tk.Frame(container, bg=BG_PANEL)
        right.place(relx=0.70, rely=0.03, relwidth=0.28, relheight=0.94)

        # Mic Canvas
        self.mic_canvas = tk.Canvas(right, bg=BG_PANEL, width=210, height=210, highlightthickness=0, bd=0)
        self.mic_canvas.place(relx=0.5, rely=0.25, anchor="center")


        # Perfect Circle with mic emoji in center
        self.mic_circle = self.mic_canvas.create_oval(
            10, 10, 190, 190, fill=ACCENT, outline=TEXT, width=3
        )
        self.mic_label = self.mic_canvas.create_text(
            100, 100, text="🎤", font=("Segoe UI Emoji", 48), fill=BG_MAIN
        )

        # Buttons Frame
        btn_frame = tk.Frame(right, bg=BG_PANEL)
        btn_frame.place(relx=0.10, rely=0.55, relwidth=0.80, relheight=0.35)

        self.start_btn = tk.Button(
            btn_frame, text="▶ Start Listening", font=self.h2,
            fg=TEXT, bg=BG_MAIN, activebackground=ACCENT,
            activeforeground=TEXT, relief="solid", bd=2,
            command=self.start_listen
        )
        self.stop_btn = tk.Button(
            btn_frame, text="■ Stop Listening", font=self.h2,
            fg=TEXT, bg=BG_MAIN, activebackground=ACCENT,
            activeforeground=TEXT, relief="solid", bd=2,
            command=self.stop_listen
        )
        self.one_shot_btn = tk.Button(
            btn_frame, text="● One-Shot Listening", font=self.h2,
            fg=TEXT, bg=BG_MAIN, activebackground=ACCENT,
            activeforeground=TEXT, relief="solid", bd=2,
            command=self.one_shot_command
        )
       
        # Pack buttons with uniform style
        for btn in (self.start_btn, self.stop_btn, self.one_shot_btn):
            btn.pack(fill="x", pady=6, ipady=4)


        # State & Animations 
        self.after(150, self._poll_ui_queue)

        # Bind mic circle to toggle listening
        self.mic_canvas.tag_bind(self.mic_circle, "<Button-1>", lambda e: self._toggle_listen())
        self.mic_canvas.tag_bind(self.mic_label, "<Button-1>", lambda e: self._toggle_listen())

        # Welcome Text
        self._drain_startup()


    # Utility Functions 
    def _drain_startup(self):
        self._poll_ui_queue(initial=True)

    def _append_chat(self, text, kind="assistant"):
        self.chat_area.config(state=tk.NORMAL)
        tag = "assistant" if kind != "user" else "user"
        timestamp = datetime.now().strftime("%I:%M %p")
        prefix = "You" if kind == "user" else "Assistant"
        self.chat_area.insert(tk.END, f"\n{prefix} ({timestamp}): {text}\n", tag)
        self.chat_area.see(tk.END)
        self.chat_area.config(state=tk.DISABLED)

    def _poll_ui_queue(self, initial=False):
        try:
            while True:
                item = ui_queue.get_nowait()
                if isinstance(item, tuple) and item and item[0] == "__QUIT__":
                    self.destroy()
                    return
                if isinstance(item, tuple):
                    self._append_chat(item[0], item[1])
                else:
                    self._append_chat(str(item), "assistant")
        except queue.Empty:
            pass
        self.after(150, self._poll_ui_queue)

    #Button Commands
    def start_listen(self): start_listenning()
    def stop_listen(self): stop_listening()
    def one_shot_command(self): threading.Thread(target=one_shot_listen, daemon=True).start()
    def _toggle_listen(self):
        if listening_flag.is_set(): stop_listening()
        else: start_listenning()
    def quit_confirm(self):
        if messagebox.askokcancel("Exit", "Do you want to exit the assistant?"):
            self.destroy()

# One shot listen:
def one_shot_listen():
    r = sr.Recognizer()  # creates speech recognization obj comes from speechrecognition lib
    try:
        with sr.Microphone() as source:  # open microphone as input taker
            ui_log("Listening for one command...", "assistant")
            r.adjust_for_ambient_noise(source, duration=0.5)
            audio = r.listen(source, timeout=6, phrase_time_limit=10)
            try:
                text = r.recognize_google(audio)
                process_command(text)
            except sr.UnknownValueError:
                ui_log("Sorry, I didn't catch that.", "assistant")
                speak("Sorry, I didn't catch that.")
            except sr.RequestError as e:
                ui_log(f"Speech service error: {e}", "assistant")
                speak("Speech service error.")
    except Exception as e:
        ui_log(f"Microphone error: {e}", "assistant")
        speak("Microphone error.")


# Main:
def main():
    app = GUI()
    app.mainloop()  # loop which keeps your app running without it app will open and close instantly


if __name__ == "__main__":  # it ensures it only runs when you actually execute the file, not when you import it somewhere else.
    main()
