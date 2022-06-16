from webbrowser import get

import pyautogui
import time
import keyboard
import cv2 as cv

import numpy as np
import random
import win32api
import win32con
import re


# Press the green button in the gutter to run the script.

def upload():
    try:
        x, y = pyautogui.locateCenterOnScreen('d_arrow.png', grayscale=True, confidence=0.75)
        pyautogui.click(x, y)
        time.sleep(1)
        x, y = pyautogui.locateCenterOnScreen('json.png', grayscale=True, confidence=0.75)
        pyautogui.click(x, y)
        x, y = pyautogui.locateCenterOnScreen('pficon.png', grayscale=True, confidence=0.8)
        pyautogui.click(x, y)
        time.sleep(2)

        #pyautogui.click('File.png')#Click file button
        keyboard.press_and_release('alt')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('enter')

        #pyautogui.click('importJSON.png')#click JSON import
        time.sleep(1)
        #pyautogui.click('import.png') #click import
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('enter')
        time.sleep(.75)

        pyautogui.keyDown("shift")
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        pyautogui.keyUp("shift")
        pyautogui.typewrite("downloads")
        keyboard.press_and_release('enter')
        time.sleep(.7)

        #x, y = pyautogui.locateCenterOnScreen('Downloads.png', confidence=0.75)
        #pyautogui.click(x, y)
        keyboard.press_and_release('tab')
        keyboard.press_and_release('space')
        keyboard.press_and_release('enter')


        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')#Pending

        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('enter') #build structural model\
        time.sleep(.5)
        if pyautogui.locateCenterOnScreen('Save Changes.png', grayscale=True, confidence=0.75) is not None:
            pyautogui.click('Save Changes.png')
            keyboard.press_and_release('tab')
            keyboard.press_and_release('enter')
            keyboard.press_and_release('tab')
            keyboard.press_and_release('enter')


        keyboard.press_and_release('enter')
        time.sleep(.5)

        if pyautogui.locateCenterOnScreen('Save Changes.png', grayscale=True, confidence=0.75) is not None:
            pyautogui.click('Save Changes.png')
            keyboard.press_and_release('tab')
            keyboard.press_and_release('enter')
            keyboard.press_and_release('tab')
            keyboard.press_and_release('enter')

        if pyautogui.locateCenterOnScreen('missing.png', grayscale=True, confidence=0.75) is not None:
            return 0
        if pyautogui.locateCenterOnScreen('objecterror.png', grayscale=True, confidence=0.75) is not None:
            return 0

        time.sleep(.60)
        keyboard.press_and_release('alt')
        keyboard.press_and_release('right')
        keyboard.press_and_release('right')
        keyboard.press_and_release('enter') #Tools
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('down')
        keyboard.press_and_release('enter') #Options
        time.sleep(.40)


        pyautogui.click('Pole Deflection.png')
        #pyautogui.click('no limit.png')

        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('down')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('enter')
        return 0


    except:
        print(" error occurred")
        return 0



def cvTest():
    points = findClickPositions('d_arrowHD.png', 'haystack_arrows.png', debug_mode='points')
    print(points)
    print('Done.')
    return 0

def saveIR():
    try:
        x, y = pyautogui.locateCenterOnScreen('d_arrow.png', grayscale=True, confidence=0.75)
        pyautogui.click(x, y)
        time.sleep(1)
        x, y = pyautogui.locateCenterOnScreen('IKE_R.png', grayscale=True, confidence=0.8)
        pyautogui.click(x, y)
        time.sleep(2)
        pyautogui.tripleClick(1941, 185)
        pyautogui.hotkey('ctrl', 'c')
        pyautogui.click(3779, 120)
        time.sleep(1)
        keyboard.press_and_release('enter')
        time.sleep(3)
        pyautogui.typewrite("P.")
        pyautogui.hotkey('ctrl', 'v')
        keyboard.press_and_release("enter")
        pyautogui.click(2386, 16)
        return 0
        #print(pyautogui.position())
    except:
        print(" Save IR error occurred")
        return 0


def saveSC():
    try:
        time.sleep(1)
        pyautogui.tripleClick(3568, 152)
        pyautogui.hotkey('ctrl', 'c')
        keyboard.press_and_release("print screen")
        time.sleep(.5)
        pyautogui.moveTo(2555, 142)
        time.sleep(1)
        pyautogui.dragTo(3239, 1043, duration=.80)
        pyautogui.rightClick()
        pyautogui.click(3304, 971)
        time.sleep(.25)
        pyautogui.typewrite("P.")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.typewrite("_Measurements")
        keyboard.press_and_release("enter")
        return 0

    except:
        print("Save SC error occurred")
        return 0

def saveSC_CloseUp():
    try:
        time.sleep(1)
        pyautogui.tripleClick(3568, 152)
        pyautogui.hotkey('ctrl', 'c')
        keyboard.press_and_release("print screen")
        time.sleep(.5)
        pyautogui.moveTo(2555, 142)
        time.sleep(1)
        pyautogui.dragTo(3239, 1043, duration=.80)
        pyautogui.rightClick()
        pyautogui.click(3304, 971)
        time.sleep(.25)
        pyautogui.typewrite("P.")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.typewrite("_Close-Up Measurements")
        keyboard.press_and_release("enter")
        return 0

    except:
        print("Save SC error occurred")
        return 0

def getCoords():
    try:
        time.sleep(1)
        print(pyautogui.position())
        return 0
    except:
        print("minor error")
        return 0

def findClickPositions(needle_img_path, haystack_img_path, threshold=0.5, debug_mode=None):
    # https://docs.opencv.org/4.2.0/d4/da8/group__imgcodecs.html
    haystack_img = cv.imread(haystack_img_path, cv.IMREAD_UNCHANGED)
    needle_img = cv.imread(needle_img_path, cv.IMREAD_UNCHANGED)
    # Save the dimensions of the needle image
    needle_w = needle_img.shape[1]
    needle_h = needle_img.shape[0]

    # There are 6 methods to choose from:
    # TM_CCOEFF, TM_CCOEFF_NORMED, TM_CCORR, TM_CCORR_NORMED, TM_SQDIFF, TM_SQDIFF_NORMED
    method = cv.TM_CCOEFF_NORMED
    result = cv.matchTemplate(haystack_img, needle_img, method)

    # Get the all the positions from the match result that exceed our threshold
    locations = np.where(result >= threshold)
    locations = list(zip(*locations[::-1]))
    # print(locations)

    # You'll notice a lot of overlapping rectangles get drawn. We can eliminate those redundant
    # locations by using groupRectangles().
    # First we need to create the list of [x, y, w, h] rectangles
    rectangles = []
    for loc in locations:
        rect = [int(loc[0]), int(loc[1]), needle_w, needle_h]
        # Add every box to the list twice in order to retain single (non-overlapping) boxes
        rectangles.append(rect)
        rectangles.append(rect)
    # Apply group rectangles.
    # The groupThreshold parameter should usually be 1. If you put it at 0 then no grouping is
    # done. If you put it at 2 then an object needs at least 3 overlapping rectangles to appear
    # in the result. I've set eps to 0.5, which is:
    # "Relative difference between sides of the rectangles to merge them into a group."
    rectangles, weights = cv.groupRectangles(rectangles, groupThreshold=2, eps=0.5)
    # print(rectangles)

    points = []
    if len(rectangles):
        # print('Found needle.')

        line_color = (0, 255, 0)
        line_type = cv.LINE_4
        marker_color = (255, 0, 255)
        marker_type = cv.MARKER_CROSS

        # Loop over all the rectangles
        for (x, y, w, h) in rectangles:

            # Determine the center position
            center_x = x + int(w / 2)
            center_y = y + int(h / 2)
            # Save the points
            points.append((center_x, center_y))

            if debug_mode == 'rectangles':
                # Determine the box position
                top_left = (x, y)
                bottom_right = (x + w, y + h)
                # Draw the box
                cv.rectangle(haystack_img, top_left, bottom_right, color=line_color,
                             lineType=line_type, thickness=2)
            elif debug_mode == 'points':
                # Draw the center point
                cv.drawMarker(haystack_img, (center_x, center_y),
                              color=marker_color, markerType=marker_type,
                              markerSize=40, thickness=2)

        if debug_mode:
            cv.imshow('Matches', haystack_img)
            cv.waitKey()
            # cv.imwrite('result_click_point.jpg', haystack_img)

    return points


if __name__ == '__main__':

    while True:
        if keyboard.is_pressed("Shift+Home"):
            print("uploading")
            upload()

        if keyboard.is_pressed("Ctrl+Home"):
            cvTest()

        if keyboard.is_pressed("Ctrl+delete"):
            saveIR()

        if keyboard.is_pressed("Ctrl + *"):
            saveSC()

        if keyboard.is_pressed("Ctrl + space"):
            saveSC_CloseUp()

        if keyboard.is_pressed("Ctrl + -"):
            getCoords()


















