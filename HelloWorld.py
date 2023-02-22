import pyautogui
import time

p_list = pyautogui.locateAllOnScreen("C:\\Python\\source.png")
p_list = list(p_list)
print(p_list)
p_center = pyautogui.center(p_list[0])
pyautogui.click(p_center)

