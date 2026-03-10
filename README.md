# Voice Assistant & Virtual Mouse

## Description

This is the second stage of my Virtual Mouse project, extending it with an integrated **voice assistant**. The application listens for spoken commands and responds by performing system actions such as opening File Explorer, closing windows, minimizing applications, navigating back, refreshing the screen, and displaying the current time. It also provides two hands-free control modes triggered by voice:

- **Control Mouse** – uses your webcam and hand-landmark detection to move the cursor and trigger single/double clicks purely through finger gestures.
- **Control Sound** – uses hand gestures captured by the webcam to raise or lower system volume in real time.

- Previous stage: https://github.com/nngeek195/Virtual-mouse

## Skills Learned

| Skill | Details |
|---|---|
| **Python Programming** | Structuring a multi-feature CLI application, functions, loops, and control flow |
| **Speech Recognition** | Capturing microphone audio, calibrating for ambient noise, and transcribing speech to text with `speech_recognition` + Google Speech API |
| **Computer Vision** | Reading live camera frames, colour-space conversion, and rendering overlays with **OpenCV** (`cv2`) |
| **Hand Landmark Detection** | Detecting and tracking 21 hand keypoints in real time using **MediaPipe Hands** |
| **Gesture Recognition** | Calculating Euclidean distances between finger landmarks (`math.hypot`) to detect pinch gestures for clicking and volume control |
| **Mouse / Keyboard Automation** | Simulating mouse movement, single-click, double-click, and keyboard shortcuts with **PyAutoGUI** |
| **Smooth Motion Algorithms** | Applying a weighted moving-average (exponential smoothing) to reduce jitter in cursor movement |
| **Debouncing** | Preventing unintended repeated actions by enforcing minimum time intervals between events |
| **System Audio Control** | Reading and setting the master volume scalar through the Windows Core Audio API via **pycaw** and **comtypes** |
| **Real-time Processing** | Maintaining a responsive event loop that processes camera frames, measures FPS, and handles voice input concurrently |
| **Windows API Integration** | Activating COM interfaces (`CLSCTX_ALL`) and casting pointers with `ctypes`/`comtypes` to control hardware-level audio endpoints |

## Implemented Features

- Voice-activated system commands: open This PC, close window, go back, hide (minimize), refresh, show time
- Gesture-controlled cursor with smooth movement
- Gesture-triggered single click, double click, and exit gesture
- Gesture-controlled system volume (increase / decrease)

## How to Use

1. Clone the repo
2. Install the requirements
    ```bash
    pip install -r requirements.txt
    ```
3. Run the assistant
    ```bash
    python control.py
    ```
4. Speak a command (e.g. *"control mouse"*, *"close"*, *"control sound"*) and enjoy!

## Video Tutorial
https://www.linkedin.com/posts/niranga-nayanajith-548a0a302_i-am-going-to-introduce-my-first-voice-assistant-activity-7218724627139756032-vQa9?utm_source=share&utm_medium=member_desktop

<img width="400" alt="image" src="https://github.com/user-attachments/assets/5f06d0e8-f84f-4c20-b6d0-74ddc4f78f98" />

 




