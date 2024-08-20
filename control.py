import speech_recognition as sr
import pyautogui
import cv2
import mediapipe as mp
import math
import time
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


def recognize_speech_from_mic():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Please wait. Calibrating microphone...")
        recognizer.adjust_for_ambient_noise(source, duration=1)
        print("Microphone calibrated. Start speaking...")
        audio = recognizer.listen(source)
        try:
            print("Recognizing...")
            text = recognizer.recognize_google(audio)
            print("You said: " + text)
            return text
        except sr.UnknownValueError:
            print("Could not understand the audio")
        except sr.RequestError as e:
            print(f"Request error: {e}")
        return None


#mouse cursor controller    
def handle_command(command):
    if "this pc" in command.lower():
        open_this_pc()
    if "close" in command.lower():
        Close_active_window() 
    if "go back" in command.lower():
        go_back()
    if "hide" in command.lower():
        Minimize()    
    if "refresh" in command.lower():
        refresh()  
    if "time" in command.lower():
        Time()  
    if 'control mouse' in command.lower():
        control_mouse()
    if 'control sound' in command.lower():
        control_sound()  
              

def open_this_pc():
    # Open "This PC" on Windows
    pyautogui.hotkey('win', 'e')
    time.sleep(1)
    #pyautogui.write('This PC')
    #pyautogui.press('enter')

def  Minimize():
    pyautogui.hotkey('win', 'down')
    pyautogui.hotkey('win', 'down')
    time.sleep(1)

def  Close_active_window():
    pyautogui.hotkey('alt', 'f4')
    time.sleep(1)
    #pyautogui.write('Close active window')
    #pyautogui.press('enter') 

def  go_back():
    pyautogui.hotkey('alt', 'left')
    time.sleep(1)
    #pyautogui.write('go back')Refresh
    #pyautogui.press('enter')  

def  refresh():
    pyautogui.hotkey('ctrl', 'f5')
    time.sleep(1)
    #pyautogui.write('Refresh')
    #pyautogui.press('enter')   

def  Time():
    pyautogui.hotkey('win', 'alt' , 'd')
    time.sleep(1)
    #pyautogui.press('enter') 

def control_mouse():
    mp_hands = mp.solutions.hands
    hands = mp_hands.Hands(static_image_mode=False, max_num_hands=1, min_detection_confidence=0.5, min_tracking_confidence=0.5)
    mp_drawing = mp.solutions.drawing_utils

    # Open the default camera
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("Error: Could not open video capture device.")
        exit()

    # Distance threshold for click detection
    CLICK_THRESHOLD = 15  # Adjust this value based on your needs
    # Debounce time in seconds
    DEBOUNCE_TIME = 0.2

    last_click_time = 0

    # Variables to store the previous mouse position for smoothing
    prev_mouse_x, prev_mouse_y = pyautogui.position()
    SMOOTHING_FACTOR = 0.5  # Adjust this value for more or less smoothing

    def draw_axes(frame, landmarks, w, h):
        index_finger_tip = landmarks[8]
        middle_finger_tip = landmarks[12]
        little_finger_tip = landmarks[6]
        cx1, cy1 = int(index_finger_tip.x * w), int(index_finger_tip.y * h)
        cx2, cy2 = int(middle_finger_tip.x * w), int(middle_finger_tip.y * h)
        cx3, cy3 = int(little_finger_tip.x * w), int(little_finger_tip.y * h)

        # Draw axes
        cv2.line(frame, (0, cy1), (w, cy1), (0, 255, 0), 2)  # Horizontal line for index finger
        cv2.line(frame, (cx1, 0), (cx1, h), (0, 255, 0), 2)  # Vertical line for index finger
        cv2.line(frame, (0, cy2), (w, cy2), (255, 0, 0), 2)  # Horizontal line for middle finger
        cv2.line(frame, (cx2, 0), (cx2, h), (255, 0, 0), 2)  # Vertical line for middle finger

    while cap.isOpened():
        start_time = time.time()
        ret, frame = cap.read()
        if not ret:
            print("Error: Could not read frame.")
            break

        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        results = hands.process(frame_rgb)

        if results.multi_hand_landmarks:
            for hand_landmarks in results.multi_hand_landmarks:
                mp_drawing.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                landmarks = hand_landmarks.landmark
                h, w, _ = frame.shape
                draw_axes(frame, landmarks, w, h)

                # Use pyautogui to move mouse with smoothing
                index_finger_tip = landmarks[8]
                screen_width, screen_height = pyautogui.size()
                screen_x = screen_width / w * int(index_finger_tip.x * w)
                screen_y = screen_height / h * int(index_finger_tip.y * h)
                
                # Smooth the mouse movement
                smooth_x = prev_mouse_x * (1 - SMOOTHING_FACTOR) + screen_x * SMOOTHING_FACTOR
                smooth_y = prev_mouse_y * (1 - SMOOTHING_FACTOR) + screen_y * SMOOTHING_FACTOR
                pyautogui.moveTo(smooth_x, smooth_y)
                prev_mouse_x, prev_mouse_y = smooth_x, smooth_y

                # Check distance between thumb tip and index finger tip for single click
                thumb_tip = landmarks[4]
                index_x, index_y = int(index_finger_tip.x * w), int(index_finger_tip.y * h)
                thumb_x, thumb_y = int(thumb_tip.x * w), int(thumb_tip.y * h)
                distance_thumb_index = math.hypot(index_x - thumb_x, index_y - thumb_y)

                little_finger_tip = landmarks[6]
                little_x, little_y = int(little_finger_tip.x * w), int(little_finger_tip.y * h)
                distance_thumb_little = math.hypot(little_x - thumb_x, little_y - thumb_y)  

                # Check distance between index finger tip and middle finger tip for double click
                middle_finger_tip = landmarks[12]
                middle_x, middle_y = int(middle_finger_tip.x * w), int(middle_finger_tip.y * h)
                distance_index_middle = math.hypot(index_x - middle_x, index_y - middle_y)
                distance_index_little = math.hypot(index_x - little_x, index_y - little_y)
                distance_thumb_middle = math.hypot(middle_x - thumb_x, middle_y - thumb_y)

                # Perform single click if the distance is below the threshold and debounce time has passed
                current_time = time.time()
                if distance_thumb_index < CLICK_THRESHOLD and (current_time - last_click_time) > DEBOUNCE_TIME:
                    pyautogui.doubleClick()
                    last_click_time = current_time

                if distance_thumb_middle < CLICK_THRESHOLD and (current_time - last_click_time) > DEBOUNCE_TIME:
                    pyautogui.click()
                    last_click_time = current_time

                if distance_index_little < CLICK_THRESHOLD and (current_time - last_click_time) > DEBOUNCE_TIME:
                    print(f"Exiting: distance_thumb_little={distance_thumb_little}, threshold={CLICK_THRESHOLD}")
                    cap.release()
                    cv2.destroyAllWindows()
                    return  # Exit the function properly
                    last_click_time = current_time

        cv2.imshow('Hand Tracking', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

    # Calculate and print the FPS (Frames Per Second)
    end_time = time.time()
    fps = 1 / (end_time - start_time)

def control_sound():       
        def get_volume_interface():
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            return volume

        def increase_volume(volume, increment=0.01):
            current_volume = volume.GetMasterVolumeLevelScalar()
            new_volume = min(current_volume + increment, 1.0)
            volume.SetMasterVolumeLevelScalar(new_volume, None)

        def decrease_volume(volume, decrement=0.01):
            current_volume = volume.GetMasterVolumeLevelScalar()
            new_volume = max(current_volume - decrement, 0.0)
            volume.SetMasterVolumeLevelScalar(new_volume, None)

        volume = get_volume_interface()

        # Initialize MediaPipe Hands
        mp_hands = mp.solutions.hands
        hands = mp_hands.Hands(static_image_mode=False, max_num_hands=1, min_detection_confidence=0.5, min_tracking_confidence=0.5)
        mp_drawing = mp.solutions.drawing_utils

        # Open the default camera
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("Error: Could not open video capture device.")
            exit()

        # Distance threshold for click detection
        CLICK_THRESHOLD = 30  # Adjust this value based on your needs

        # Variables to store the previous mouse position for smoothing
        prev_mouse_x, prev_mouse_y = pyautogui.position()
        SMOOTHING_FACTOR = 0.5  # Adjust this value for more or less smoothing

        def draw_axes(frame, landmarks, w, h):
            index_finger_tip = landmarks[8]
            middle_finger_tip = landmarks[12]
            cx1, cy1 = int(index_finger_tip.x * w), int(index_finger_tip.y * h)
            cx2, cy2 = int(middle_finger_tip.x * w), int(middle_finger_tip.y * h)

            # Draw axes
            cv2.line(frame, (0, cy1), (w, cy1), (0, 255, 0), 2)  # Horizontal line for index finger
            cv2.line(frame, (cx1, 0), (cx1, h), (0, 255, 0), 2)  # Vertical line for index finger
            cv2.line(frame, (0, cy2), (w, cy2), (255, 0, 0), 2)  # Horizontal line for middle finger
            cv2.line(frame, (cx2, 0), (cx2, h), (255, 0, 0), 2)  # Vertical line for middle finger

        # Variables for debouncing volume control
        last_volume_adjustment_time = 0
        VOLUME_ADJUSTMENT_INTERVAL = 0.5  # Minimum interval between volume adjustments in seconds

        while cap.isOpened():
            start_time = time.time()  # Track frame start time for FPS calculation

            ret, frame = cap.read()
            if not ret:
                print("Error: Could not read frame.")
                break

            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            results = hands.process(frame_rgb)

            if results.multi_hand_landmarks:
                for hand_landmarks in results.multi_hand_landmarks:
                    mp_drawing.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                    landmarks = hand_landmarks.landmark
                    h, w, _ = frame.shape
                    draw_axes(frame, landmarks, w, h)

                    index_finger_tip = landmarks[8]
                    thumb_tip = landmarks[4]
                    index_x, index_y = int(index_finger_tip.x * w), int(index_finger_tip.y * h)
                    thumb_x, thumb_y = int(thumb_tip.x * w), int(thumb_tip.y * h)
                    distance_thumb_index = math.hypot(index_x - thumb_x, index_y - thumb_y)

                    current_time = time.time()
                    if current_time - last_volume_adjustment_time > VOLUME_ADJUSTMENT_INTERVAL:
                        if distance_thumb_index > CLICK_THRESHOLD:
                            increase_volume(volume)
                        else:
                            decrease_volume(volume)
                        last_volume_adjustment_time = current_time

            cv2.imshow('Frame', frame)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

            # Calculate and print the FPS (Frames Per Second)
            end_time = time.time()
            fps = 1 / (end_time - start_time)



if __name__ == "__main__":
    while True:
        command = recognize_speech_from_mic()
        if command:
            handle_command(command)
