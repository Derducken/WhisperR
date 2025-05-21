# WhisperR

(Pronounced "Whisperer", not like an angry pirate's "whisperARRRrr")

A vibe-coded solution enabling dictation in ANY other app. Powered by Whisper.

![WhisperR interface](https://github.com/Derducken/WhisperR/blob/master/WhisperR.jpg)

WhisperR offers ~~two~~ three modes of operation:

- A typical continuous recording mode that can record your voice to a single file, where you start and stop the recording manually. That's optimal for dictating long strings of text, for example when crafting the first draft of an article. In this mode, transcription begins only AFTER you've stopped recording, and processes all your speech in one go. This also gives the Whisper "brains" a larger frame of reference that improves results (more speech = more data to analyze - improved accuracy).
- An auto-pause/command mode, where WhisperR detects when you speak, and only records those snippets of audio as separate audio files. It then processes them with Whisper one-by-one, returning the transcribed results whenever each audio file has been processed. In this mode, the app can also execute commands. You can "bind" specific voice commands to actions-you'd-perform-in-a-terminal through WhisperR's configuration. For example, try adding a command like "Whisperer, run Firefox" and bind the action "firefox" (or the full path to firefox's executable), and whenever WhisperR detects that command in a transcribed snippet of text, it will try to execute the respective action.
- Due to popular demand (...that I've noticed in other apps, since, for now, I'm the only one using this :-D ), Push-To-Talk mode. Note that the Push-To-Talk mode merely enables either the continuous or the VAD mode for as long as the PTT hotkey is pressed. So, its behavior also depends on if you're using VAD mode or not. If you aren't, PTT will keep recording for as long as you keep its hotkey pressed, and start transcribing when you depress it. If you are (using VAD mode), PTT will record and transcribe individual speech snippets as you talk (based on how you've configured VAD in the WhisperR's options).

## Features

*   Can work in ANY app that supports pasting text (uses typical Clipboard & CTRL+V functionality for "sending text" to the active app).
*   Three recording modes: Continuous, Voice-Activated Dictation, and Push-To-Talk.
*   Support for individual lines with timestamps or continuous "live" transcription mode.
*   GUI for configuration (language, model, transcription/translation mode, command mode, and more).
*   Support for custom Whisper initial prompt (can significantly increase transcription accuracy).
*   Custom Whisper prompt Import & Export. 
*   Multiple audiovisual indicators of app's state (user-customizable colored bar, icons, beeps).
*   Hotkeys for toggling recording, push-to-talk, and showing/hiding the app's window.
*   Basic user-customizable command execution based on voice input.
*   Voice command management (add, edit, delete).
*   Convenient built-in transcription scratchpad.
*   Automatic Audio Levels Calibration.


## Use Cases

*   "Type" much faster in any app by dictating instead of banging on a keyboard.
*   Hands-free computer control. If you put in the effort.
*   I'd love to think it may even help with accessibility for users with disabilities (and love it even more if someone took this base idea and ran with it to make something better for the purpose).

Note that this was vibe-coded and tested with Faster Whisper's executable on Windows 11. I'm sharing it here because others might find it useful, too. All similar solutions I personally found on Windows either were tailored more towards transcribing subtitles, or didn't come with an easy-to-use-GUI. Don't blame me if you don't like it, it doesn't work for you, or it makes your PC gain consciousness and decide to kill your cat.

## Installation
1. Download the latest version of the app from Releases (on the right of this wall of text).
2. Run WhisperR.
3. Settings > Configuration.
4. On the **General & Hotkeys** tab, click **Download/Update Engine**.
5. If you'd like, change the **Global Hotkeys** at the end of the same page.
6. Move to the **Audio & VAD** tab and select the correct **Audio Input Device** ("the one where your microphone is connected").
7. Under **Auto-Pause (VAD) Settings**, click **Calibrate** and follow the instructions on that window.
8. Move to **Notifications & Output** and enable **Auto-Paste After Transcription** if you want WhisperR to automatically paste the transcribed text into any active app (where "active app" = Notepad, Google Docs in your favorite browser, etc.)
9. Click **Save Configuration**.
10. Select your spoken **Language** from WhisperR's main window, and the transcription model you want to use (**medium** and **large** should be the best for most users).

For more information and detailed instructions on how to install, configure, and use WhisperR, check [this tutorial](https://medium.com/@Ducklord/cf2ef6763841).