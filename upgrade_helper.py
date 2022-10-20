import win32event, win32process
from os import getcwd, remove, rename, getenv
from time import sleep

sleep(2)

src = getcwd() + "/upgrade.exe"
profile_path = getenv('TEMP')[:len(getenv('TEMP')) - 4] + "zju_ahr.txt"
with open(profile_path, 'r') as f:
    data = eval(f.read())
    dst = data['exe_path']
print(dst)
remove(dst)
rename(src, dst)

handle = win32process.CreateProcess(dst, '', None, None, 0, win32process.CREATE_NO_WINDOW, None, None, win32process.STARTUPINFO())
win32event.WaitForSingleObject(handle[0], 2)



