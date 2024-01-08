import pyautogui
import time
import win32com.client
import openpyxl
import os

"""
            This file is not a part of the project
"""


def click_coordinates(x, y):
    print("check====")
    pyautogui.click(x, y)

def save_workbook(workbook, file_path):
    workbook.save(file_path)

def automate_excel_label_change():
    file_path = r"raw_dump.xlsx"

    workbook = openpyxl.load_workbook(file_path)
    # save_workbook(workbook, file_path)
    worksheet = workbook['RAW']

    time.sleep(5)

    # label_option_coordinates = (1304, 238)
    # public_option_coordinates = (1315, 328)
    # public_option_coordinates = (1213, 279)
    public_option_coordinates = (1102, 281)

    # Open Excel and navigate to the "Label" option
    # pyautogui.hotkey('alt', 'h')  # open "Home" tab
    time.sleep(1)  # Wait for the menu to appear
    # pyautogui.press('r')  # activate "Review" tab
    time.sleep(1)  # Wait for the menu to appear

    print("clicking..")
    # click_coordinates(*label_option_coordinates)

    # Wait for the "Label" options to appear
    time.sleep(1)

    print("finish..")

    # Click the "Public" option
    click_coordinates(*public_option_coordinates)

    print("final..")

    pyautogui.hotkey('ctrl', 's')


if __name__ == "__main__":
    automate_excel_label_change()