import os

import subprocess

appdataPath = os.getenv('APPDATA')
authInfoFile = "xlsx2pdf.txt"
def getPcUUID():
    command = "wmic csproduct get uuid"
    output = subprocess.check_output(command, shell=True).decode("utf-8")
    return output.strip().split("\n")[-1]

def saveAuthInfo(content):
    file_path = os.path.join(appdataPath, authInfoFile)
    with open(file_path, 'w') as file:
        file.write(content)

def readAuthInfo():
    file_path = os.path.join(appdataPath, authInfoFile)
    # 读取文件内容
    with open(file_path, 'r') as file:
        content = file.read()
    return content

def isAuthInfo():
    file_path = os.path.join(appdataPath, authInfoFile)
    return os.path.exists(file_path)

