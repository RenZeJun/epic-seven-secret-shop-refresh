import pyautogui
import time
import xlrd
import random
from numpy import random


def mouseClick(clickTimes, lOrR, img, reTry):
    time.sleep(random.uniform(0.3,1.2))
    i=0
    while i < 3:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            break
        print("can't not find picture", img)
        time.sleep(0.9)
        i+=1


def mouseClickBookMark(clickTimes, lOrR, img, reTry, purchaseComfirmBm):
    time.sleep(random.uniform(0.3,1.2))
    x=0
    if reTry == 1:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.95)
        if location is not None:
            pyautogui.click(location.x+779, location.y+29, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            x+=1
            print("find bookmarks")
            while reTry < 4:
                location2 = pyautogui.locateCenterOnScreen(purchaseComfirmBm, confidence=0.95)
                if location2 is not None:
                    print("locating purchasebutton")
                    pyautogui.click(location2.x, location2.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                    break
                reTry += 1
    if x == 0:
        print("can't not find picture(bm)")
    time.sleep(0.9)





def mainWork(img):
    i = 1
    containBM = 1
    while i < 6:
        # check for bookmarks
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            # read image
            purchaseComfirm = sheet1.row(1)[3].value
            img = sheet1.row(i)[1].value
            img2 = sheet1.row(1)[2].value
            retry = 1
            print(i)
            if i == 1:
                mouseClickBookMark(random.randint(1,3), "left", img, retry, purchaseComfirm)
                print("Left click", img)
                mouseClickBookMark(random.randint(1,3),"left",img2, retry, purchaseComfirm)
                print("Left click",img2)
                i += 1
                continue
            elif i == 3:
                mouseClickBookMark(random.randint(1,3), "left", img, retry, purchaseComfirm)
                print("Left click", img)
                mouseClickBookMark(random.randint(1,3),"left",img2, retry, purchaseComfirm)
                print("Left click",img2)
                i += 1
                continue
            mouseClick(random.randint(1,3), "left", img, retry)
            print("Left click", img)
        elif cmdType.value == 2.0:
            print(i)
            time.sleep(1)
            pyautogui.scroll(-3)
            print("scrolling 3")
            time.sleep(random.uniform(1,2.5))
        i += 1


if __name__ == '__main__':
    file = 'cmd.xls'

    wb = xlrd.open_workbook(filename=file)

    sheet1 = wb.sheet_by_index(0)
    print('good luck on bookmarks')

    key = input('press 1 to do it once and press 2 to loop to death! XD')
    if key == '1':
        mainWork(sheet1)
    elif key == '2':
        while True:
            mainWork(sheet1)
            time.sleep(0.1)
            print("waiting for 0.1 sec")
else:
    print('input error')
