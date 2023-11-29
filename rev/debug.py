import win32api
e_msg = win32api.FormatMessage(-2147024893)
# print (e_msg)

import win32com.client as win32
win32c = win32.constants
for k, v in win32c.__dicts__[0].items(): 
    print("{:<45}: {}".format(k, v))

print(win32c.__dicts__)