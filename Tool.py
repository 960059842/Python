import os
import zipfile
import xlrd
from collections import OrderedDict
from pyexcel_xls import get_data
from pyexcel_xls import save_data
import pymssql
import decimal
import time
import requests
import hashlib
import sys
import datetime
import shutil
import ctypes
#加载C#dll需要
import clr

# 发布为True
isdos=False

def getPath():
    return os.getcwd() if isdos else os.path.dirname(os.path.realpath(__file__))

libPath= getPath()

#加载dll
clr.AddReference(os.path.join(libPath,"Lib", "ETL.Utils.dll"))
clr.AddReference(os.path.join(libPath,"Lib", "PythonDll.dll"))

#将命名空间导入
from ETLUtils import *
from PythonDll import *

def Decrypt3DES(value):
    return SecurityHelper.Decrypt3DES(value)

def DecAnswer(filepath):
    return BIMUtil.DecAnswer(filepath)

def EncAnswer(answerFilePath,fileName):
    BIMUtil.EncAnswer(answerFilePath,fileName)

"""
读取Excel
for colValues in data:
    print(len(colValues))
    for value in colValues:
        print(value)
"""

def readExcel(excelpath,sheetIndex=0):
    data = xlrd.open_workbook(excelpath)
    table = data.sheets()[sheetIndex]    #用索引取第一个sheet


    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    data=[]

    for i in range(1, nrows):
        rowValues = table.row_values(i)  # 某一行数据
        data.append(rowValues);

    return data

#data=readExcel(r"C:\Users\FBC\Desktop\线上考核试题0606.xls")

"""
ms = MSSQL(host="192.168.1.1",user="sa",pwd="sa",db="testdb")
reslist = ms.ExecQuery("select * from webuser")
for i in reslist:
    print i

newsql="update webuser set name='%s' where id=1"%u'测试'
print newsql
ms.ExecNonQuery(newsql.encode('utf-8'))
"""
class MSSQL:
    def __init__(self,host,user,pwd,db):
        self.host = host
        self.user = user
        self.pwd = pwd
        self.db = db

    def __GetConnect(self):
        if not self.db:
            raise(NameError,"没有设置数据库信息")
        self.conn = pymssql.connect(host=self.host,user=self.user,password=self.pwd,database=self.db,charset="utf8")
        cur = self.conn.cursor()
        if not cur:
            raise(NameError,"连接数据库失败")
        else:
            return cur

    def ExecQuery(self,sql):
        cur = self.__GetConnect()
        cur.execute(sql)
        resList = cur.fetchall()

        #查询完毕后必须关闭连接
        self.conn.close()
        return resList

    def ExecNonQuery(self,sql):
        cur = self.__GetConnect()
        cur.execute(sql)
        self.conn.commit()
        self.conn.close()

    def ExecuteMany(self,sql,data):
        cur = self.__GetConnect()
        cur.executemany(
            sql,
            data
        )
        self.conn.commit()
        self.conn.close()

"""
七牛Helper
"""
from qiniu import build_batch_stat, Auth, BucketManager, put_file, etag, urlsafe_base64_encode
class qiniuHelper:
    def __init__(self,access_key="HNkFJ3P5_5UpP5BOaXDFWTjZ6ZK9NIAEhafchRMh",secret_key="uwxfBCn-4Rc4M6B1awZx-084Kt-bMXYJbufZNKLk"):
        self.access_key = access_key
        self.secret_key = secret_key

    def __GetBucket(self):
        self.q = Auth(self.access_key, self.secret_key)
        bucket = BucketManager(self.q)
        if not bucket:
            raise(NameError,"连接七牛失败")
        else:
            return bucket

    """
    获取七牛文件列表
    """
    def getListFile(self, bucket_name="bimuptest", prefix=None):
        # 列举条目
        limit = 100
        # 列举出除'/'的所有文件以及以'/'为分隔的所有前缀
        delimiter = None
        # 标记
        marker = None

        bucket = self.__GetBucket()

        listfile = []
        while True:
            ret, eof, info = bucket.list(bucket_name, prefix, marker, limit, delimiter)

            for item in ret["items"]:
                listfile.append(item)

            if "marker" in ret.keys():
                marker = ret["marker"]
            else:
                return listfile

    """
    修改文件存储类型 1表示低频存储，0是标准存储
    """
    def changeSaveType(self,bucket_name="bimuptest",key="",saveType=1):
        bucket = self.__GetBucket()
        ret, info = bucket.change_type(bucket_name, key, saveType)  # 1表示低频存储，0是标准存储
        print(info)
        return info

    """
    获取七牛下载文件地址
    """
    def getDownUrl(self,url, expires=3600):
        self.q = Auth(self.access_key, self.secret_key)
        downUrl=self.q.private_download_url(url, expires)
        print(downUrl)
        return downUrl

    """
    上传文件
    """
    def upLoadFile(self,bucket_name="bimuptest",keyName="ExamCode/SafeCode/QuestionID.anst",filePath=""):
        self.q = Auth(self.access_key, self.secret_key)

        # 生成上传 Token，可以指定过期时间等
        token = self.q.upload_token(bucket_name, keyName, 3600)

        ret, info = put_file(token, keyName, filePath)
        print(info)
        assert ret['key'] == keyName
        assert ret['hash'] == etag(filePath)


"""
print('This is a \033[1;35m test \033[0m!')
print('This is a \033[1;32;43m test \033[0m!')
print('\033[1;33;44mThis is a test !\033[0m')

显示颜色格式：\033[显示方式;字体色;背景色m......[\033[0m]
-------------------------------------------
字体色     |       背景色     |      颜色描述
-------------------------------------------
30        |        40       |       黑色
31        |        41       |       红色
32        |        42       |       绿色
33        |        43       |       黃色
34        |        44       |       蓝色
35        |        45       |       紫红色
36        |        46       |       青蓝色
37        |        47       |       白色
-------------------------------------------
-------------------------------
显示方式     |      效果
-------------------------------
0           |     终端默认设置
1           |     高亮显示
4           |     使用下划线
5           |     闪烁
7           |     反白显示
8           |     不可见
-------------------------------
"""

def output(value,delimiter="*",number=25):
    print((delimiter * number), value, (delimiter * number))

def printColor(value):
    if isdos:
        printPink(value+'\n')
    else:
        print("\033[1;35m%s\033[0m" % (value))

def timeStampToTime(stamp):
    timeArray = time.localtime(stamp)
    otherStyleTime = time.strftime("%Y-%m-%d %H:%M:%S", timeArray)
    return otherStyleTime

#获取目录下的所有文件
def getFileList(dir,fileList):
    newDir = dir
    #是否是文件
    if os.path.isfile(dir):
        fileList.append(dir)
    #是否是文件夹
    elif os.path.isdir(dir):
        for s in os.listdir(dir):
            #如果需要忽略某些文件夹，使用以下代码
            #if s == "xxx":
                #continue
            newDir=os.path.join(dir,s)
            getFileList(newDir, fileList)
    return fileList

def GetFileList(dir):
    fileList=[]
    if os.path.isdir(dir):
        for name in os.listdir(dir):
            filepath= os.path.join(dir, name)
            if os.path.isfile(filepath):
                fileList.append(filepath)

    return fileList

#获取文件扩展名
def getExtension(filepath):
    return os.path.splitext(filepath)[1]

#返回不具有扩展名的指定路径字符串的文件名
def getFileNameWithoutExtension(filepath):
    filename= os.path.split(filepath)[1]
    return os.path.splitext(filename)[0]

#返回指定路径字符串的目录信息
def getFileName(filepath):
    return os.path.split(filepath)[1]

#返回指定路径字符串的目录信息
def getDicName(filepath):
    return os.path.split(filepath)[0]

#获取文件大小,KB
def getFileSize(filePath):
    fsize = os.path.getsize(filePath)
    fsize = fsize/float(1024)
    return round(fsize) #保留两位小数

def unZip(fileName):
    """unzip zip file"""
    zipFile = zipfile.ZipFile(fileName)
    zipPath=os.path.join(getDicName(fileName),"UnZip",getFileNameWithoutExtension(fileName))
    if os.path.isdir(zipPath):
        shutil.rmtree(zipPath)

    os.makedirs(zipPath)

    for name in zipFile.namelist():
        #解决zip解压乱码问题
        filename=name
        try:
            filename = name.encode("cp437").decode('gbk')
        except:
            filename = name
        filename=filename.replace('/', os.path.sep)
        path = os.path.join(zipPath, filename)
        #print(name,path)
        if name[-1]=="/":
            if os.path.exists(path):
                pass
            else:
                os.makedirs(path)
            continue
        else:
            dir=getDicName(path)
            if os.path.exists(dir):
                pass
            else:
                os.makedirs(dir)

        data = zipFile.read(name)
        with open(path, "wb") as f:
            f.write(data)

    zipFile.close()

    return zipPath

def downFile(url,downFilePath):
    # 下载
    request = requests.get(url)
    # 判断文件是否存在不存在删除
    if os.path.isfile(downFilePath):
        os.remove(downFilePath)

    # 判断路径是否存在不存在则删除
    downPath = getDicName(downFilePath)
    if os.path.exists(downPath):
        pass
    else:
        os.makedirs(downPath)

    with open(downFilePath, "wb") as code:
        code.write(request.content)


def getFileMd5(filename):
    if not os.path.isfile(filename):
        return ""
    myhash = hashlib.md5()
    f = open(filename,'rb')
    while True:
        b = f.read(8096)
        if not b :
            break
        myhash.update(b)
    f.close()
    return myhash.hexdigest()

"""
DOS 下输入颜色
"""
STD_INPUT_HANDLE = -10
STD_OUTPUT_HANDLE = -11
STD_ERROR_HANDLE = -12

# 字体颜色定义 ,关键在于颜色编码，由2位十六进制组成，分别取0~f，前一位指的是背景色，后一位指的是字体色
#由于该函数的限制，应该是只有这16种，可以前景色与背景色组合。也可以几种颜色通过或运算组合，组合后还是在这16种颜色中

# Windows CMD命令行 字体颜色定义 text colors
FOREGROUND_BLACK = 0x00 # black.
FOREGROUND_DARKBLUE = 0x01 # dark blue.
FOREGROUND_DARKGREEN = 0x02 # dark green.
FOREGROUND_DARKSKYBLUE = 0x03 # dark skyblue.
FOREGROUND_DARKRED = 0x04 # dark red.
FOREGROUND_DARKPINK = 0x05 # dark pink.
FOREGROUND_DARKYELLOW = 0x06 # dark yellow.
FOREGROUND_DARKWHITE = 0x07 # dark white.
FOREGROUND_DARKGRAY = 0x08 # dark gray.
FOREGROUND_BLUE = 0x09 # blue.
FOREGROUND_GREEN = 0x0a # green.
FOREGROUND_SKYBLUE = 0x0b # skyblue.
FOREGROUND_RED = 0x0c # red.
FOREGROUND_PINK = 0x0d # pink.
FOREGROUND_YELLOW = 0x0e # yellow.
FOREGROUND_WHITE = 0x0f # white.


# Windows CMD命令行 背景颜色定义 background colors
BACKGROUND_BLUE = 0x10 # dark blue.
BACKGROUND_GREEN = 0x20 # dark green.
BACKGROUND_DARKSKYBLUE = 0x30 # dark skyblue.
BACKGROUND_DARKRED = 0x40 # dark red.
BACKGROUND_DARKPINK = 0x50 # dark pink.
BACKGROUND_DARKYELLOW = 0x60 # dark yellow.
BACKGROUND_DARKWHITE = 0x70 # dark white.
BACKGROUND_DARKGRAY = 0x80 # dark gray.
BACKGROUND_BLUE = 0x90 # blue.
BACKGROUND_GREEN = 0xa0 # green.
BACKGROUND_SKYBLUE = 0xb0 # skyblue.
BACKGROUND_RED = 0xc0 # red.
BACKGROUND_PINK = 0xd0 # pink.
BACKGROUND_YELLOW = 0xe0 # yellow.
BACKGROUND_WHITE = 0xf0 # white.

# get handle
std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)

def set_cmd_text_color(color, handle=std_out_handle):
    Bool = ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)
    return Bool

#reset white
def resetColor():
    set_cmd_text_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE)

###############################################################

#暗蓝色
#dark blue
def printDarkBlue(mess):
    set_cmd_text_color(FOREGROUND_DARKBLUE)
    sys.stdout.write(mess)
    resetColor()

#暗绿色
#dark green
def printDarkGreen(mess):
    set_cmd_text_color(FOREGROUND_DARKGREEN)
    sys.stdout.write(mess)
    resetColor()

#暗天蓝色
#dark sky blue
def printDarkSkyBlue(mess):
    set_cmd_text_color(FOREGROUND_DARKSKYBLUE)
    sys.stdout.write(mess)
    resetColor()

#暗红色
#dark red
def printDarkRed(mess):
    set_cmd_text_color(FOREGROUND_DARKRED)
    sys.stdout.write(mess)
    resetColor()

#暗粉红色
#dark pink
def printDarkPink(mess):
    set_cmd_text_color(FOREGROUND_DARKPINK)
    sys.stdout.write(mess)
    resetColor()

#暗黄色
#dark yellow
def printDarkYellow(mess):
    set_cmd_text_color(FOREGROUND_DARKYELLOW)
    sys.stdout.write(mess)
    resetColor()

#暗白色
#dark white
def printDarkWhite(mess):
    set_cmd_text_color(FOREGROUND_DARKWHITE)
    sys.stdout.write(mess)
    resetColor()

#暗灰色
#dark gray
def printDarkGray(mess):
    set_cmd_text_color(FOREGROUND_DARKGRAY)
    sys.stdout.write(mess)
    resetColor()

#蓝色
#blue
def printBlue(mess):
    set_cmd_text_color(FOREGROUND_BLUE)
    sys.stdout.write(mess)
    resetColor()

#绿色
#green
def printGreen(mess):
    set_cmd_text_color(FOREGROUND_GREEN)
    sys.stdout.write(mess)
    resetColor()

#天蓝色
#sky blue
def printSkyBlue(mess):
    set_cmd_text_color(FOREGROUND_SKYBLUE)
    sys.stdout.write(mess)
    resetColor()

#红色
#red
def printRed(mess):
    set_cmd_text_color(FOREGROUND_RED)
    sys.stdout.write(mess)
    resetColor()

#粉红色
#pink
def printPink(mess):
    set_cmd_text_color(FOREGROUND_PINK)
    sys.stdout.write(mess)
    resetColor()

#黄色
#yellow
def printYellow(mess):
    set_cmd_text_color(FOREGROUND_YELLOW)
    sys.stdout.write(mess)
    resetColor()

#白色
#white
def printWhite(mess):
    set_cmd_text_color(FOREGROUND_WHITE)
    sys.stdout.write(mess)
    resetColor()

##################################################

#白底黑字
#white bkground and black text
def printWhiteBlack(mess):
    set_cmd_text_color(FOREGROUND_BLACK | BACKGROUND_WHITE)
    sys.stdout.write(mess)
    resetColor()

#白底黑字
#white bkground and black text
def printWhiteBlack_2(mess):
    set_cmd_text_color(0xf0)
    sys.stdout.write(mess)
    resetColor()


#黄底蓝字
#white bkground and black text
def printYellowRed(mess):
    set_cmd_text_color(BACKGROUND_YELLOW | FOREGROUND_RED)
    sys.stdout.write(mess)
    resetColor()


##############################################################

"""
printDarkBlue(u'printDarkBlue:暗蓝色文字\n')
printDarkGreen(u'printDarkGreen:暗绿色文字\n')
printDarkSkyBlue(u'printDarkSkyBlue:暗天蓝色文字\n')
printDarkRed(u'printDarkRed:暗红色文字\n')
printDarkPink(u'printDarkPink:暗粉红色文字\n')
printDarkYellow(u'printDarkYellow:暗黄色文字\n')
printDarkWhite(u'printDarkWhite:暗白色文字\n')
printDarkGray(u'printDarkGray:暗灰色文字\n')
printBlue(u'printBlue:蓝色文字\n')
printGreen(u'printGreen:绿色文字\n')
printSkyBlue(u'printSkyBlue:天蓝色文字\n')
printRed(u'printRed:红色文字\n')
printPink(u'printPink:粉红色文字\n')
printYellow(u'printYellow:黄色文字\n')
printWhite(u'printWhite:白色文字\n')
printWhiteBlack(u'printWhiteBlack:白底黑字输出\n')
printWhiteBlack_2(u'printWhiteBlack_2:白底黑字输出\n')
printYellowRed(u'printYellowRed:黄底红字输出\n')
"""

 # 输出文件的md5值以及记录运行时间
"""
starttime = datetime.datetime.now()
filepath=r"E:\test.rar.zip"
print(getFileMd5(filepath))
endtime = datetime.datetime.now()
print('运行时间：%ds'%((endtime-starttime).seconds))
"""

""" 路径
print("__file__=%s" % __file__)
print("os.path.realpath(__file__)=%s" % os.path.realpath(__file__))
print("os.path.dirname(os.path.realpath(__file__))=%s" % os.path.dirname(os.path.realpath(__file__)))
print("os.path.split(os.path.realpath(__file__))=%s" % os.path.split(os.path.realpath(__file__))[0])
print("os.path.abspath(__file__)=%s" % os.path.abspath(__file__))
print("os.getcwd()=%s" % os.getcwd())
print("sys.path[0]=%s" % sys.path[0])
print("sys.argv[0]=%s" % sys.argv[0])
"""
