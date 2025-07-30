import sys
import os
import psutil
import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QVBoxLayout, QHBoxLayout
)
from PyQt5.QtGui import QPixmap, QPainter, QFont
from PyQt5.QtCore import Qt, QTimer, QPoint
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
import ctypes


class AnimatedAvatar(QWidget):
    def __init__(self, idle_path, mouth_open_path, blink_path, size=(200, 200)):
        super().__init__()

        self.idle = QPixmap(idle_path).scaled(*size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.mouth_open = QPixmap(mouth_open_path).scaled(*size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.blink = QPixmap(blink_path).scaled(*size, Qt.KeepAspectRatio, Qt.SmoothTransformation)

        self.current = self.idle

        self.setFixedSize(self.idle.size())
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)

        self.blink_timer = QTimer(self)
        self.blink_timer.timeout.connect(self.blink_once)
        self.blink_timer.start(5000)


    def paintEvent(self, event):
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self.current)

    def blink_once(self):
        # Show blink sprite for 300 ms, then back to idle
        self.current = self.blink
        self.update()
        QTimer.singleShot(300, self.return_to_idle)

    def return_to_idle(self):
        self.current = self.idle
        self.update()

    def talk(self):
        # Show mouth open sprite for 700 ms, then back to idle
        self.current = self.mouth_open
        self.update()
        QTimer.singleShot(700, self.return_to_idle)


class TextBubble(QLabel):
    def __init__(self):
        super().__init__()
        self.setWordWrap(True)
        self.setFont(QFont("Arial", 11))
        self.setStyleSheet("""
            background-color: rgba(255, 255, 255, 220);
            color: black;
            border: 1px solid gray;
            border-radius: 12px;
            padding: 10px;
        """)
        self.setVisible(False)
        self.setFixedWidth(250)

        self.full_text = ""
        self.displayed_text = ""
        self.char_index = 0
        self.timer = QTimer()
        self.timer.timeout.connect(self._next_character)

    def show_animated_text(self, message, interval=30):
        self.full_text = message
        self.displayed_text = ""
        self.char_index = 0
        self.setText("")
        self.setVisible(True)
        self.timer.start(interval)

    def _next_character(self):
        if self.char_index < len(self.full_text):
            self.displayed_text += self.full_text[self.char_index]
            self.setText(self.displayed_text)
            self.adjustSize()
            self.char_index += 1
        else:
            self.timer.stop()
            # Hide after delay
            QTimer.singleShot(6000, lambda: self.setVisible(False))





class DesktopAssistant(QWidget):
    def __init__(self, idle_path, mouth_open_path, blink_path):
        super().__init__()

        self.avatar = AnimatedAvatar(idle_path, mouth_open_path, blink_path)
        self.bubble = TextBubble()
        self.input_line = QLineEdit()
        self.input_line.setPlaceholderText("Ask me something...")
        self.input_line.returnPressed.connect(self.handle_input)

        self.init_ui()
        self.drag_pos = None

        self.fun_facts = [
            "Honey never spoils.",
            "Octopuses have three hearts.",
            "Bananas are berries, but strawberries aren't."
        ]

        self.fix_tips = {
            "slow": "Try restarting your PC and closing unused apps.",
            "internet": "Check your router or try resetting your network adapter.",
            "battery": "Try calibrating your battery or replace if it's old.",
            "crash": "Update your drivers and check for overheating.",
            "disk": "You can free up space by running 'cleanmgr' (Disk Cleanup) or deleting temp files using Win+R → %temp%.",
        }

        self.user_profile = {
            "name": None,
            "favorite_color": None,
            "birthday": None
        }


    def init_ui(self):
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setMinimumSize(400, 250)  # Set a reasonable minimum
        self.adjustSize()              # Allow it to grow for bubble


        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(5)

        # Left vertical layout: avatar on top, input directly below with avatar width
        left_layout = QVBoxLayout()
        left_layout.setSpacing(5)
        left_layout.addWidget(self.avatar, alignment=Qt.AlignLeft)

        self.input_line.setFixedWidth(self.avatar.width())  # Match avatar width
        left_layout.addWidget(self.input_line, alignment=Qt.AlignLeft)

        # Right side: bubble (aligned top-left)
        right_layout = QVBoxLayout()
        right_layout.addWidget(self.bubble, alignment=Qt.AlignTop)

        main_layout.addLayout(left_layout)
        main_layout.addLayout(right_layout)

        self.setLayout(main_layout)
        self.summary_timer = QTimer(self)
        self.summary_timer.timeout.connect(self.hourly_summary)
        self.summary_timer.start(3600000)  # 1 hour in milliseconds

        self.game_active = False
        self.secret_number = None
        self.rps_active = False




    def hourly_summary(self):
        now = datetime.datetime.now().strftime("It's %H:%M on %A, %B %d, %Y.")

        batt_info = "Battery info not available."
        batt = psutil.sensors_battery()
        if batt:
            plugged = "charging" if batt.power_plugged else "on battery"
            batt_info = f"Battery at {batt.percent}%, {plugged}."

        cpu = psutil.cpu_percent(interval=0.5)
        mem = psutil.virtual_memory().percent
        usage = f"CPU: {cpu}%, Memory: {mem}%."

        import random
        fact = f"Fun fact: {random.choice(self.fun_facts)}"

        summary = f"{now}\n{batt_info}\n{usage}\n{fact}"
        self.avatar.talk()
        self.show_bubble(summary)



    def handle_input(self):
        user_text = self.input_line.text().strip().lower()
        if not user_text:
            return

        response = self.respond(user_text)
        self.avatar.talk()
        self.show_bubble(response)
        self.input_line.clear()


    def respond(self, text):
        def toggle_mute(mute: bool):
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = ctypes.cast(interface, ctypes.POINTER(IAudioEndpointVolume))
            volume.SetMute(mute, None)

        if any(kw in text for kw in ["time", "date"]):
            now = datetime.datetime.now()
            return now.strftime("It's %H:%M on %A, %B %d, %Y.")
        if "battery" in text:
            batt = psutil.sensors_battery()
            if batt:
                plugged = "charging" if batt.power_plugged else "on battery"
                return f"Battery is at {batt.percent}% and {plugged}."
            return "Battery info not available."
        if "health" in text or "cpu" in text or "memory" in text:
            cpu = psutil.cpu_percent(interval=0.5)
            mem = psutil.virtual_memory().percent
            return f"CPU usage is {cpu}%. Memory usage is {mem}%."
        if "fix" in text or "slow" in text or "problem" in text or "crash" in text  or "disk" in text:
            for key, tip in self.fix_tips.items():
                if key in text:
                    return tip
            return "Try restarting your computer or checking for updates."
        import subprocess

        if "run cleanup" in text or "open cleanup" in text or "clean" in text:
            try:
                subprocess.Popen("cleanmgr")
                return "Launching Disk Cleanup..."
            except Exception as e:
                return f"Failed to launch Disk Cleanup: {e}"

        if "fun fact" in text or "fact" in text:
            import random
            return random.choice(self.fun_facts)
        
        if any(kw in text for kw in ["summary", "status report", "how am i doing"]):
            now = datetime.datetime.now().strftime("It's %H:%M on %A, %B %d, %Y.")

            batt_info = "Battery info not available."
            batt = psutil.sensors_battery()
            if batt:
                plugged = "charging" if batt.power_plugged else "on battery"
                batt_info = f"Battery at {batt.percent}%, {plugged}."

            cpu = psutil.cpu_percent(interval=0.5)
            mem = psutil.virtual_memory().percent
            usage = f"CPU: {cpu}%, Memory: {mem}%."

            import random
            fact = f"Fun fact: {random.choice(self.fun_facts)}"

            return f"{now}\n{batt_info}\n{usage}\n{fact}"

        if text.startswith("open "):
            app_name = text.replace("open ", "").strip()
            return self.launch_app(app_name)
        
        if "clipboard" in text or "read clipboard" in text:
            return self.read_clipboard()
        
        if text in ["exit", "quit", "close", "bye"]:
            self.avatar.talk()
            self.show_bubble("Goodbye! Shutting down...")
            QTimer.singleShot(2000, QApplication.quit)  # 2 sec delay to show message
            return ""

        import random

        # Start the game
        if "play guess" in text or "guess number" in text:
            self.game_active = True
            self.secret_number = random.randint(1, 10)
            return "I'm thinking of a number between 1 and 10. Try to guess it!"

        # During the game
        if self.game_active and text.startswith("guess"):
            try:
                guess = int(text.split()[1])
                if guess < self.secret_number:
                    return "Too low! Try again."
                elif guess > self.secret_number:
                    return "Too high! Try again."
                else:
                    self.game_active = False
                    return "Correct! You guessed it!"
            except:
                return "Please type like: guess 5"

        # Start Rock Paper Scissors
        if "rock paper scissors" in text or "play rps" in text:
            self.rps_active = True
            return "Let's play Rock, Paper, Scissors! Type your move: rock, paper, or scissors."

        # Handle move
        if self.rps_active and text in ["rock", "paper", "scissors"]:
            user = text
            bot = random.choice(["rock", "paper", "scissors"])
            result = ""

            if user == bot:
                result = "It's a tie!"
            elif (user == "rock" and bot == "scissors") or \
                (user == "paper" and bot == "rock") or \
                (user == "scissors" and bot == "paper"):
                result = "You win!"
            else:
                result = "I win!"

            self.rps_active = False
            return f"You chose {user}, I chose {bot}. {result}"

        # Set name
        if text.startswith("my name is "):
            name = text.replace("my name is ", "").strip().capitalize()
            self.user_profile["name"] = name
            return f"Nice to meet you, {name}!"

        # Set favorite color
        if text.startswith("my favorite color is "):
            color = text.replace("my favorite color is ", "").strip()
            self.user_profile["favorite_color"] = color
            return f"I'll remember that your favorite color is {color}."

        # Set birthday
        if "my birthday is" in text:
            date = text.split("my birthday is")[-1].strip()
            self.user_profile["birthday"] = date
            return f"Got it! Your birthday is on {date}."

        # Recall known info
        if "what do you know about me" in text:
            profile = self.user_profile
            return (
                f"Your name is {profile['name'] or 'unknown'}, "
                f"your favorite color is {profile['favorite_color'] or 'unknown'}, "
                f"and your birthday is {profile['birthday'] or 'unknown'}."
            )

        if "recent files" in text or "show recent" in text:
            # You can change this to your Documents folder or any path
            target_dir = os.path.expanduser("~/Documents")
            
            if not os.path.exists(target_dir):
                return "I couldn't find the Documents folder."

            # Get all files with last modified time
            recent_files = []
            for root, _, files in os.walk(target_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    try:
                        mtime = os.path.getmtime(full_path)
                        recent_files.append((full_path, mtime))
                    except Exception:
                        continue

            # Sort by modified time (descending)
            recent_files.sort(key=lambda x: x[1], reverse=True)
            top_files = recent_files[:3]

            if not top_files:
                return "No recent files found."

            formatted = "\n".join(
                f"{os.path.basename(path)} ({datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')})"
                for path, mtime in top_files
            )
            return f"Here are your 3 most recent files in Documents:\n{formatted}"

        if "wifi info" in text or "wifi status" in text or "network info" in text:
            try:
                # Run Windows command to show wifi info
                result = subprocess.check_output(["netsh", "wlan", "show", "interfaces"], encoding='utf-8')
                
                ssid = None
                signal = None
                state = None

                for line in result.splitlines():
                    line = line.strip()
                    if line.startswith("SSID"):
                        ssid = line.split(":", 1)[1].strip()
                    elif line.startswith("Signal"):
                        signal = line.split(":", 1)[1].strip()
                    elif line.startswith("State"):
                        state = line.split(":", 1)[1].strip()

                if ssid and signal and state:
                    return f"Wi-Fi '{ssid}' is {state} with signal strength {signal}."
                else:
                    return "Could not retrieve complete Wi-Fi information."

            except Exception as e:
                return "Sorry, I couldn't get the Wi-Fi information."

        if text in ["hi", "hello"]:
            return "Hello."

        if text in ["mute audio", "mute sound"]:
            try:
                toggle_mute(True)
                return "Audio muted."
            except Exception as e:
                return "Sorry, I couldn't mute the audio."

        if text in ["unmute audio", "unmute sound"]:
            try:
                toggle_mute(False)
                return "Audio unmuted."
            except Exception as e:
                return "Sorry, I couldn't unmute the audio."

        return "Sorry, I don't understand. Try asking something else."

    def show_bubble(self, message):
        self.bubble.show_animated_text(message)

    def read_clipboard(self):
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        if text:
            return f"Clipboard says:\n{text[:300]}"  # You can adjust the char limit
        return "Clipboard is empty or not text."


    def launch_app(self, app_name):
        app_paths = {
            "chrome": r"C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
            "edge": r"C:/Program Files/Internet Explorer/iexplore.exe",
        }

        app_name = app_name.lower()
        if app_name in app_paths:
            try:
                os.startfile(app_paths[app_name])
                return f"Launching {app_name.capitalize()}..."
            except Exception as e:
                return f"Couldn't launch {app_name}: {e}"
        return f"I don't know how to open {app_name}."
    
    
    def hide_bubble(self):
        self.bubble.setVisible(False)
        self.adjustSize()  # Resize down after bubble hides



    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_pos = event.globalPos() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, event):
        if self.drag_pos and event.buttons() & Qt.LeftButton:
            self.move(event.globalPos() - self.drag_pos)

    def mouseReleaseEvent(self, event):
        self.drag_pos = None


if __name__ == "__main__":
    idle_img = "C:/Users/Dell/Documents/deskteto/tet00.png"
    mouth_open_img = "C:/Users/Dell/Documents/deskteto/tet01.png"
    blink_img = "C:/Users/Dell/Documents/deskteto/tet02.png"

    for img_path in [idle_img, mouth_open_img, blink_img]:
        if not os.path.exists(img_path):
            print(f"Missing image: {img_path}")
            sys.exit(1)

    app = QApplication(sys.argv)
    assistant = DesktopAssistant(idle_img, mouth_open_img, blink_img)
    assistant.show()
    sys.exit(app.exec_())
