import os
import pyautogui
from pathlib import Path
from pynput import keyboard
from fpdf import FPDF
import win32gui, win32ui, win32api
from win32api import GetSystemMetrics
from PIL import Image

# init drawing
dc = win32gui.GetDC(0)
dcObj = win32ui.CreateDCFromHandle(dc)
hwnd = win32gui.WindowFromPoint((0,0))
monitor = (0, 0, GetSystemMetrics(0), GetSystemMetrics(1))

# init screenshot region
start_x = 0
start_y = 0
end_x = 0
end_y = 0


screenshot_path = "./pages"
Path(screenshot_path).mkdir(parents=True, exist_ok=True)

pdf = FPDF(unit='mm')
_keyboard = keyboard.Controller()


def screenshot_and_scroll():
    print('create screenshot')
    print(f"{start_x} - {start_y} : {abs(end_x-start_x)} - {abs(end_y-start_y)}")
    # myScreenshot = pyautogui.screenshot()
    myScreenshot = pyautogui.screenshot(region=(start_x,start_y,abs(end_x-start_x),abs(end_y-start_y)))
    myScreenshot.save(f'{screenshot_path}/page_{screenshot_and_scroll.counter}.png')
    screenshot_and_scroll.counter += 1 #increment page counter

    _keyboard.press(keyboard.Key.page_down)
    _keyboard.release(keyboard.Key.page_down)
    
screenshot_and_scroll.counter = 1 #init page counter
    
def create_pdf():
    for image in os.listdir(screenshot_path):
        cover = Image.open(f'{screenshot_path}/{image}')
        width, height = cover.size
        width, height = float(width * 0.264583), float(height * 0.264583)
        # pdf.add_page(format=(width, height))
        pdf.add_page()
        pdf.image(f'{screenshot_path}/{image}',x=0,y=0, w= width,h=height)
        # pdf.image(f'{screenshot_path}/{image}')
    pdf.output("./mybook.pdf", "F")



def define_screenshot_section():
    global start_x
    global start_y
    global end_x
    global end_y
    mouse_pos = win32gui.GetCursorPos()
    dcObj.Rectangle((mouse_pos[0], mouse_pos[1], mouse_pos[0]-15, mouse_pos[1]-15))
    if(define_screenshot_section.counter % 2):
        start_x = mouse_pos[0]
        start_y = mouse_pos[1]
    else:
        end_x = mouse_pos[0]
        end_y = mouse_pos[1]
        
    print(f"{start_x} - {start_y} : {end_x} - {end_y}")
    win32gui.InvalidateRect(hwnd, monitor, True) # Refresh the entire monitor
    define_screenshot_section.counter += 1
    
define_screenshot_section.counter = 1
    
    
def program_exit():
    print("exit")
    exit(0)

with keyboard.GlobalHotKeys({
        'n': screenshot_and_scroll,
        'b': define_screenshot_section,
        'e': program_exit,
        'p': create_pdf}) as h:
    h.join()