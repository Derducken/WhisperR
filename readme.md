# WhisperR

A vibe-coded solution enabling dictation in ANY other app. Powered by Whisper and FFDshow.

WhisperR offers two modes of operation:

- A typical continuous recording mode that can record your voice ((using FFDshow)) to a single file, where you start and stop the recording manually. That's optimal for dictating long strings of text, for example when crafting the first draft of an article. In this mode, transcription begins only AFTER you've stopped recording, and processes all your speech in one go. This also gives the Whisper "brains" a larger frame of reference that improves results (more speech = more data to analyze - improved accuracy).
- An auto-pause/command mode, where WhisperR detects when you speak, and only records those snippets of audio as separate audio files. It then processes them with Whisper one-by-one, returning the transcribed results whenever each audio file has been processed. In this mode, the app can also execute commands. You can "bind" specific voice commands to actions-you'd-perform-in-a-terminal through WhisperR's configuration. For example, try adding a command like "Whisperer, run Firefox" and bind the action "firefox" (or the full path to firefox's executable), and whenever WhisperR detects that command in a transcribed snippet of text, it will try to execute the respective action.

## Features

*   GUI for configuration (language, model, transcription/translation mode, command mode)
*   Hotkey for toggling recording
*   Command execution based on voice input
*   Audio recording using FFDShow
*   Configuration options for Whisper executable, audio input device, and export folder
*   Voice command management (add, edit, delete)

## Use Cases

*   "Type" much faster in any app by dictating instead of banging on a keyboard.
*   Hands-free computer control. If you put in the effort.
*   I'd love to think it may even help with accessibility for users with disabilities (and love it even better if someone took this base idea and ran with it to make something better for the purpose).

Note that this was vibe-coded and tested with Faster Whisper's executable on Windows 11. I'm sharing it here because others might find it useful, too - I personally hadn't found such an app on Windows that wasn't tailored more towards transcribing subtitles or offering an easy-to-use-GUI. Don't blame me if you don't like it, it doesn't work for you, or it makes your PC gain consciousness and decide to kill your cat.