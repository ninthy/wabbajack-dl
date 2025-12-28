import win32gui
import ctypes
import win32ui
import win32con
from win32con import WM_LBUTTONDOWN, WM_LBUTTONUP, MK_LBUTTON, WM_MOUSEWHEEL, WM_KEYDOWN, WM_KEYUP, VK_DOWN, VK_NEXT, WM_VSCROLL, SB_LINEDOWN, SB_PAGEDOWN
from win32api import MAKELONG
from PIL import Image
import cv2
import numpy as np
from rich import print
import time
from rich.prompt import IntPrompt
from rich.progress import (
    Progress,
    TextColumn,
    BarColumn,
    TaskProgressColumn,
    TimeRemainingColumn,
    TimeElapsedColumn,
)
from datetime import timedelta

def find_window():
    hWnd = win32gui.FindWindow(None, "Wabbajack")
    return hWnd

def get_screenshot(hWnd):
    left, top, right, bot = win32gui.GetWindowRect(hWnd)
    w = right - left
    h = bot - top

    hwnd_dc = win32gui.GetWindowDC(hWnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()

    save_bitmap = win32ui.CreateBitmap()
    save_bitmap.CreateCompatibleBitmap(mfc_dc, w, h)

    save_dc.SelectObject(save_bitmap)
    
    result = ctypes.windll.user32.PrintWindow(hWnd, save_dc.GetSafeHdc(), 2)
    if result != 1:
        save_dc.BitBlt((0, 0), (w, h), mfc_dc, (0, 0), win32con.SRCCOPY)

    bmpinfo = save_bitmap.GetInfo()
    bmpstr = save_bitmap.GetBitmapBits(True)

    im = Image.frombuffer(
        "RGB",
        (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
        bmpstr,
        "raw",
        "BGRX",
        0,
        1,
    )

    win32gui.DeleteObject(save_bitmap.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hWnd, hwnd_dc)

    return im

def find_image(hwnd, image: Image.Image, confidence=0.8):
    screenshot = get_screenshot(hwnd)
    
    img_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    template_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    
    result = cv2.matchTemplate(img_cv, template_cv, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    
    if max_val < confidence:
        return None
        
    match_x, match_y = max_loc
    h, w = template_cv.shape[:2]
    center_x = match_x + w // 2
    center_y = match_y + h // 2
    
    rect = win32gui.GetWindowRect(hwnd)
    screen_x = rect[0] + center_x
    screen_y = rect[1] + center_y
    
    client_point = win32gui.ScreenToClient(hwnd, (screen_x, screen_y))
    return client_point

def post_click(hWnd, x, y):
    def callback(child_hwnd, result_list):
        if win32gui.GetClassName(child_hwnd) == "Chrome_RenderWidgetHostHWND":
            result_list.append(child_hwnd)
    
    child_hwnds = []
    try:
        win32gui.EnumChildWindows(hWnd, callback, child_hwnds)
    except Exception:
         pass

    target_hwnd = child_hwnds[0] if child_hwnds else hWnd

    screen_point = win32gui.ClientToScreen(hWnd, (x, y))
    child_point = win32gui.ScreenToClient(target_hwnd, screen_point)
    
    lParam = MAKELONG(child_point[0], child_point[1])

    win32gui.PostMessage(target_hwnd, WM_LBUTTONDOWN, MK_LBUTTON, lParam)
    win32gui.PostMessage(target_hwnd, WM_LBUTTONUP, MK_LBUTTON, lParam)

def post_scroll(hWnd, x, y, scroll_path):
    scroll_button = find_image(hWnd, Image.open(scroll_path))
    if not scroll_button:
        return
    post_click(hWnd, *scroll_button)

def main():
    hWnd = find_window()
    folder = "imgs"
    image_path = f"{folder}/button.png"
    scroll_path = f"{folder}/scroll.png" 
    if not hWnd:
        print("[red]Failed to find Wabbajack window, exiting")
        return


    rect = win32gui.GetWindowRect(hWnd)

    center = (rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2
    scroll_count = 0
    scroll_threshold = 20
    total_mods = IntPrompt.ask("How many mods do you want to install?")
    current_index = IntPrompt.ask("How many mods are already downloaded (if the current page says 100, enter 99)?", default=0)


    start_time = time.time() 
    with Progress(
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        TextColumn("{task.completed}/{task.total} mods"),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
    ) as progress:
        task = progress.add_task("[cyan]Installing mods...", total=total_mods, completed=current_index)
        
        while not progress.finished:
            if scroll_count > scroll_threshold:
                break
            
            coords = find_image(hWnd, Image.open(image_path))
            
            if not coords:
                progress.update(task, description="[yellow]Scrolling...[/yellow]")
                post_scroll(hWnd, *center, scroll_path)
                scroll_count += 1
                time.sleep(.5)
                continue
            
            progress.update(task, description="[green]Installing mod...[/green]")
            time.sleep(.3)
            post_click(hWnd, coords[0], coords[1])
            progress.advance(task)
            scroll_count = 0
            time.sleep(5)

    if scroll_count < scroll_threshold:
        print(f"You saved [yellow]{timedelta(seconds=int(time.time() - start_time))}[/yellow] of manual labor!")
    else:
        print(f"[red]Failed to find [bold]{image_path}[/bold] after multiple scrolls, exiting")

if __name__ == "__main__":
    main()
    input("Press enter to exit")