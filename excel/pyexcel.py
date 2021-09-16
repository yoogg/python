#python操作excel功能（建议使用python3.9）
#---需要安装的包---#
#pip install Pillow
#pip install pywin32
import os
import time
from PIL import ImageGrab
try:
    from win32com.client import Dispatch
except:
    pass

class py_excel:
    def __init__(self,wkbPath='',shtNo=0,kill=True):
        if kill:
            os.system('taskkill /F /IM EXCEL.EXE')   #关闭excel进程
        self._app = Dispatch('excel.application')    #打开excel程序
        self._app.Visible = 1
        if wkbPath == '':
            self._wkb = self._app.Workbooks.Add()    #如果没有工作簿，则增加工作簿
        else:
            self._wkb = self._app.Workbooks.Open(wkbPath)   #打开工作簿
        self._sht = self._wkb.Worksheets[shtNo]             #第一个工作表
    
    @property
    def workbook(self):
        return self._wkb 

    @property 
    def worksheet(self):
        return self._sht

    def close(self,savePath=''):
        if savePath == '':
            self._wkb.Close(0)
            os.system('taskkill /F /IM EXCEL.EXE') 
        elif savePath == 'cat':
            self._wkb.Save()
            self._wkb.Close(1)
        else:
            saveDir = '\\'.join(savePath.split('\\')[:-1])  #用\分隔路径，返回列表，去除最后一个元素
            if not os.path.exists(saveDir):                 #检查文件夹是否存在，如果不存在，就创建
                os.makedirs(saveDir)
            # self._wkb.Save()
            self._wkb.SaveAs(savePath)                      #self._wkb保存在此路径下
            self._wkb.Close(1)
            # self._app.DisplayAlerts = False
            # self._app.Quit()
            # self._app.DisplayAlerts = True

    def refresh(self,sleepTime=10,saveName='',RunMacro='',StopClose=False):
        if RunMacro != '':
            self._app.Run(RunMacro)
        else:
            self._wkb.RefreshAll()     #刷新
        time.sleep(sleepTime)

        if saveName == '':
            self._wkb.Save()
            if not StopClose:
                self._wkb.Close(1)
            newFilePath = None
        else:
            saveDir = 'C:\\mailps\\%s' % int(time.time()*1000000)
            newFilePath = '%s\\%s' % (saveDir,saveName)
            if not os.path.exists(saveDir):
                os.makedirs(saveDir)
            self._wkb.Save()
            self._wkb.SaveAs(newFilePath)
            if not StopClose:
                self._app.Quit()
        return newFilePath

    def macro(self,macroName='宏1',sleepTime=10):
        self._app.run(macroName)                   #运行指定宏
        time.sleep(sleepTime)

    def getCell(self,Row,Col,shtNo=0):
        if shtNo!=0:
            self._sht = self._wkb.Worksheets[shtNo]       #第一个工作表
        return self._sht.cells(Row,Col).value
    
    def getAll(self,shtNo=0):
        if shtNo!=0:
            self._sht = self._wkb.Worksheets[shtNo]       #第一个工作表
        return self._sht.UsedRange()

    def getTxt(self,fanwei,sht=''):
        if sht!='':
            self._sht = self._wkb.Worksheets[sht]
        return self._sht.Range(fanwei).value

    def autoScreenRange(self,sheetNames='',fmt='jpg'):
        '''自动截取现有数据区域的图片
        如果取默认值则遍历所有sheet
        '''
        sheetnameList = []
        if sheetNames == '':
            for sht in self._wkb.Sheets:
                sheetnameList.append(sht.Name)
        else:
            sheetnameList = sheetNames

        imgpaths = []
        for sheetname in sheetnameList:
            img = self._autoScreenRange(sheetname,fmt)
            imgpaths.extend(img)
        return imgpaths

    def _autoScreenRange(self,sheetName,fmt='jpg'):
        imglist = []
        self._sht = self._wkb.Sheets[sheetName]
        # 删除原有的
        for shp in self._sht.Shapes:
            shp.Delete()
        # 截图 保存到本地
        self._sht.UsedRange.CopyPicture()
        time.sleep(2)
        self._sht.Paste()
        for shp in self._sht.Shapes:
            shp.Copy()
            time.sleep(2)
            img = ImageGrab.grabclipboard()
            imgpath = os.path.join(os.path.dirname(__file__),'img\\{imgname}.{fmt}'.format(imgname=time.strftime('%Y%m%d%H%M%S'),fmt=fmt))
            imgsaveDir = '\\'.join(imgpath.split('\\')[:-1])  #用\分隔路径，返回列表，去除最后一个元素
            if not os.path.exists(imgsaveDir):                 #检查文件夹是否存在，如果不存在，就创建
                os.makedirs(imgsaveDir)
            try:
                if fmt == 'jpg':
                    img=img.convert('RGB')
                    img.save(imgpath,'jpeg')
                else:
                    img.save(imgpath)
            except Exception as e:
                print(e)
            imglist.append(imgpath)
        return imglist

    def savePic(self,PicNames=[],picFormat='jpg',shtNo=0):
        self._sht = self._wkb.Worksheets[shtNo]             #第一个工作表
        from sys import argv
        dirList = argv[0].split('\\')[:-1]
        newDir = '\\'.join(dirList)
        self._picPaths = []
        for pic in PicNames:
            # print(self._sht.name)
            self._sht.Shapes(pic).Copy()
            self._picPath = '%s\\%s.%s' % (newDir,pic,picFormat)
            print('图片路径:%s' % self._picPath)
            time.sleep(3)
            img = ImageGrab.grabclipboard()  #拍摄剪贴板图像的快照
            if picFormat == 'jpg':
                img = img.convert('RGB')     #大多数模型都只支持RGB格式的图片，转换为RGB格式的图片
            img.save(self._picPath)
            self._picPaths.append(self._picPath)
            time.sleep(1)
        print('图片地址集:%s' % self._picPaths)
        return self._picPaths

    def savePicIndex(self,PicIndex=1,afterFix='jpg',shtNo=0):
        self._sht = self._wkb.Worksheets[shtNo]             #第一个工作表
        from sys import argv
        dirList = argv[0].split('\\')[:-1]
        newDir = '\\'.join(dirList)
        self._picPaths = []
        # print(self._sht.name)
        if isinstance(PicIndex,int):       #isinstance() 函数来判断一个对象是否是一个已知的类型，如判断PicIndex是否是整数型
            self._sht.Shapes(PicIndex).Copy()
            self._picPath = '%s\\%s.%s' % (newDir,PicIndex,afterFix)
            print(self._picPath)
            time.sleep(3)
            img = ImageGrab.grabclipboard()
            img.save(self._picPath)
            time.sleep(1)
            print('图片地址:%s' % self._picPath)
            self._picPaths.append(self._picPath)
        else:
            self._picPaths = None
        return self._picPaths

    def savePicIndexPath(self,PicIndex=1,afterFix='jpg',Folder=r'c:\pyoutput',shtNo=0):
        self._sht = self._wkb.Worksheets[shtNo]             #第一个工作表
        from sys import argv
        from PIL import ImageGrab
        import os 
        if not os.path.exists(Folder):
            os.makedirs(Folder)
        # print(self._sht.name)
        if isinstance(PicIndex,int):
            self._sht.Shapes(PicIndex).Copy()
            self._picPath = '%s\\%s.%s' % (Folder,int(time.time()),afterFix)
            time.sleep(1)
            img = ImageGrab.grabclipboard()
            img.save(self._picPath)
            time.sleep(1)
            print('图片地址:%s' % self._picPath)
        else:
            self._picPaths = None
        return self._picPath

    def tuple2sht(self,ttuple,Title=[]):
        # 嵌套的tuple写入到sht
        #1 写标题
        startRow = 1
        startCol = 1
        for x in Title:
            self._sht.cells(startRow,startCol).value = str(x)
            startCol += 1

        #2 写内容
        startRow = 2
        for x in ttuple:
            startCol = 1
            # print(x)
            for y in x:
                try:
                    self._sht.cells(startRow,startCol).value = str(y)
                except Exception:
                    self._sht.cells(startRow,startCol).value = ''
                startCol += 1
            startRow += 1
        print(':' * 20,'Wrote',':' * 20)