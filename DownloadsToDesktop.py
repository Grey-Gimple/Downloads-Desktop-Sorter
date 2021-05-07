# required modules
import os
import shutil
import time

def DownloadsToDesktop(startDir, endDir):
    # get list of files in startDir
    files = os.listdir(startDir)

    # check if the list of files is not empty
    if len(files) != 0:
        # iterate code block for every file in files
        for f in files:
            # seperate the file name and file extension into seperate variable
            file_name, file_ext = os.path.splitext(f)

            if file_ext != '.tmp' and file_ext != '.crdownload':
                # get the start path of the file
                fileStartPath = startDir + '\\' + file_name + file_ext
                
                # get the default end path of the file
                fileEndPath = endDir + '\\' + file_name + file_ext
                
                # increament counter
                i = 0
                # check if the file exists already
                while os.path.exists(fileEndPath):
                    # increate the iteration of duplicate file
                    i+=1
                    # create new file end path with the duplicate number
                    fileEndPath = endDir + '\\' + file_name + '_' + str(i) + file_ext

                # move file from its start path to its end path
                shutil.move(fileStartPath, fileEndPath)
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                si.wShowWindow = subprocess.SW_HIDE # default
                subprocess.call(r'C:\\Users\\Owner\\Documents\\RefreshDesktop\\refresh.bat', startupinfo=si)

# run continuously
while True:
    # try and execpt for if the process crashes or is stopped and will break the loop
    try:
        # wait one second before running the functions
        time.sleep(1)
        # calling function
        DownloadsToDesktop(r'C:\\Users\\Owner\Downloads',
                           r'C:\\Users\\Owner\Desktop')

    except KeyboardInterrupt:
        break
