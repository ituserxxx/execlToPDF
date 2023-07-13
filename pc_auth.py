import json
import os

import subprocess

# appdataPath = os.getenv('APPDATA')
appdataPath = os.getcwd()
authInfoFile = "xlsx2pdf.auth"


def getPcUUID():
    command = "wmic csproduct get uuid"
    output = subprocess.check_output(command, shell=True).decode("utf-8")
    return output.strip().split("\n")[-1]


def initData():
    return {
        "is_permanent": False,
        "number_of_times": 2,  # second
        "pc_id": getPcUUID(),
        "auth_code":""
    }


def saveAuthInfo(content):
    file_path = os.path.join(appdataPath, authInfoFile)
    with open(file_path, 'w') as file:
        json.dump(content, file)


def readAuthInfo():
    if isExistsAuthInfo() is False:
        saveAuthInfo(
            initData()
        )
        return initData()
    file_path = os.path.join(appdataPath, authInfoFile)
    # 读取文件内容
    with open(file_path, 'r') as file:
        content = json.load(file)
    return content


def isExistsAuthInfo():
    file_path = os.path.join(appdataPath, authInfoFile)
    return os.path.exists(file_path)


def descNumberOfTimes(n):
    data = readAuthInfo()
    new_n = data["number_of_times"] - n
    if new_n < 0:
        new_n = 0
    data["number_of_times"] = new_n
    saveAuthInfo(data)
    return data

def incrNumberOfTimes(n):
    data = readAuthInfo()
    new_n = data["number_of_times"] + n
    if new_n < 0:
        new_n = 0
    data["number_of_times"] = new_n
    saveAuthInfo(data)
    return data


def getAuth(code):
    data = readAuthInfo()
    data["is_permanent"] = True
    data["auth_code"] = code
    saveAuthInfo(data)
    return data


