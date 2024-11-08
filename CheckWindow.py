# import pygetwindow as gw
# import time



# def income_tax_activate():
#     # Get the currently active window
#     active_window = gw.getActiveWindow()

#     active_window.title
#     if 'Income Tax Department' in active_window.title:
#         print('income tax ix active...')
#         return True
#     else:
#         return False

import os

download_location = 'C:\\Users\\gurus\\Downloads'

file_locatio = os.path.join(download_location, 'Form 15CB_Filed Form.pdf')

if os.path.exists(file_locatio):
    os.mkdir("C:\\Users\\gurus\\Downloads\\slips")
    os.rename(file_locatio, f"{download_location}\\slips\\guru.pdf")

    print('bna di...')
else:
    os.mkdir("C:\\Users\\gurus\\Downloads\\slips")
    print('Ko ni.....')

