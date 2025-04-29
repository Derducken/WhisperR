# TODO

## Core Functionality

*   [ ] Implement voice recording
    *   [ ] Choose audio recording library/tool (FFDShow, VOX, or Python library)
    *   [ ] Implement audio recording functionality
    *   [ ] Implement auto-pause mode (silence detection)
*   [ ] Implement Whisper integration
    *   [ ] Configure Whisper language and model options
    *   [ ] Implement transcription mode
    *   [ ] Implement translation mode
    *   [ ] Implement Whisper execution
    *   [ ] Implement file output (timestamped .md files)
    *   [ ] Copy Whisper output to clipboard
*   [ ] Implement command execution
    *   [ ] Implement voice command recognition
    *   [ ] Implement command execution logic (wildcard replacement)

## GUI

*   [ ] Implement main window
    *   [ ] Implement settings storage
    *   [ ] Implement hotkey registration
    *   [ ] Add GUI elements
        *   [ ] Language selection dropdown
        *   [ ] Model selection dropdown
        *   [ ] Transcription/translation toggle
        *   [ ] Command mode toggle
        *   [ ] Prompt text area
        *   [ ] Import prompt button
        *   [ ] Export prompt button
        *   [ ] OK button
        *   [ ] Shortcut information text
        *   [ ] Configuration button
*   [ ] Implement configuration window
    *   [ ] Whisper executable selection
    *   [ ] Audio input device selection
    *   [ ] Export folder selection
*   [ ] Implement command configuration window
    *   [ ] Add, edit, and delete voice commands
    *   [ ] Implement command input fields
    *   [ ] Implement action input fields
    *   [ ] Implement remove command button
    *   [ ] Implement add command button

## System Integration

*   [ ] Implement hotkey for toggling recording (CTRL + Alt + Space)
*   [ ] Implement hotkey for transcoding last recording (CTRL + Shift + Winkey + Space)
*   [ ] Implement hotkey for returning to main window (CTRL + Alt + Shift + Space)
*   [ ] Implement tray icon
    *   [ ] Change icon to reflect recording state
*   [ ] Implement top screen green line during recording

## Error Handling

*   [ ] Handle Whisper errors
*   [ ] Handle audio recording errors
*   [ ] Handle command execution errors

## Other

*   [ ] Implement file-based versioning
*   [ ] Test the project to ensure it doesn't contain errors
*   [ ] Don't create placeholder code unless planning to expand on it later.
*   [ ] Code from A to Z rather than just small parts that don't fulfill the user's needs.
*   [ ] Keep project files between 300-500 lines where possible.
*   [ ] Don't duplicate code; build upon existing implementations.
*   [ ] Before writing or modifying any code, always reference fixes.md, to ensure you don't repeat the same mistake twice.
*   [ ] Whenever you're ready to write or tweak some code, ask yourself one more time: is there a better way to produce an even more optimized, performant, and secure result? Go over your code and consider ways to improve it further before actually including it in a project.
*   [ ] Create fixes.md
