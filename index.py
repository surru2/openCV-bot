from functions import *
import cv2
import win32gui, win32con, win32api
import numpy as np

CALL_NUMBER = '10645'

while True:
    try:
        inCall = detect_updating_screen('templates/bIncoming.PNG', "Cisco IP Communicator")
        if inCall != 0:
            print("Incomming call")
            bCall = get_object_coord('templates/bAnswer.png', "Cisco IP Communicator")
            if bCall != 0:
                click(bCall[0], bCall[1])
                time.sleep(0.5)
                bTransferAgent = detect_updating_screen('templates/bTransferAgent.PNG', "Cisco Agent Desktop")
                if bTransferAgent != 0:
                    pyautogui.click(bTransferAgent[0], bTransferAgent[1])
                    time.sleep(0.5)
                    bDigits = getNumbers("Передать вызов")
                    for i in CALL_NUMBER:
                        click(bDigits[int(i)][0], bDigits[int(i)][1])
                    time.sleep(0.5)
                    baTransfer = get_object_coord('templates/baTransfer.PNG', "Передать вызов")
                    if baTransfer != 0:
                        click(baTransfer[0], baTransfer[1])
                        time.sleep(0.5)
                        bTransfer = detect_updating_screen('templates/bTransfer.PNG', "Cisco IP Communicator")
                        if bTransfer != 0:
                            click(bTransfer[0], bTransfer[1])
                            print("Call was transfered")
        else:
            time.sleep(1)
    except Exception as e:
        print("error: ", e)

# cv2.imshow('output', d)
# cv2.waitKey(0)
# cv2.destroyAllWindows()

