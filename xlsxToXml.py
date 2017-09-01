import xlrd
import tkinter.messagebox
from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
# import xml.dom.minidom

class XlsxToXML():
    def __init__(self):
        self.app = tkinter.Tk()
        self.app.title("转换工具")
        self.app.geometry("400x200")
        self.createUI()
        self.app.mainloop()
        
        
    def createUI(self):
        # 导入文件label和输入框
        self.importFileLabel = Label(self.app, borderwidth=1, justify="left", text="文件地址：")
        self.importFileLabel.grid(row=0, column=0, padx=10, pady=40)
        self.filePath = StringVar()
        self.filePathImport = Entry(self.app, textvariable=self.filePath, width=40)
        self.filePathImport.grid(row=0, column=1, columnspan=4)
        # 导入文件按钮
        self.button_importFile = Button(self.app, text="导入文件", width=10, command=self.selectPath)
        self.button_importFile.grid(row=1,column=1)
        # 执行转换按钮
        self.button_makeFile = Button(self.app, text="生成文件", width=10, command=self.start)
        self.button_makeFile.grid(row=1,column=2)
        # 菜单栏
        # self.app.option_add('*tearOff', FALSE)
        # menubar = Menu(self.app)
        # self.app['menu'] = menubar
        # menu_file = Menu(menubar)
        # menu_about = Menu(menubar)
        # menubar.add_cascade(menu=menu_file, label='文件')
        # menubar.add_cascade(menu=menu_about, label='关于')
        # menu_file.add_command(label='添加文件', command=self.selectPath)
        # menu_file.add_command(label='退出', command=quit)
        # menu_about.add_command(label='关于', command=self.aboutMenuInfo)  

    # def aboutMenuInfo(self):
        # messagebox.showinfo(title='说明',message='1.支持xls，xlsx格式转换为XML格式。\n2.生成的XML文件在excel文件同目录下\n3.有问题请联系：liangyuyang@orbbec.com')        
        
    def selectPath(self):
        # path_ = askdirectory()  # 文件夹路径
        path_ = askopenfilename()  # 文件路径
        self.filePath.set(path_)   
     
    '''
    def makeXML(self):
        #从list里面遍历dict，生成XML文件
        def makeTestCase(dom, name, elementName = 'testcase'):
            testcase = dom.createElement(elementName)
            testcase.setAttribute('name', name)
            return testcase
            
        def makeElement(dom, elementName, text=None,):
            mainEle = dom.createElement(elementName)
            if text != None:
                txt = dom.createTextNode("<![CDATA[" + text + "]]>")
                mainEle.appendChild(txt)
            return mainEle
            
        # 构建DOM树
        impl = xml.dom.minidom.getDOMImplementation()  
        dom = impl.createDocument(None, 'testsuite', None) 
        root = dom.documentElement
        mainElement = makeTestCase(dom, self.outPutName, elementName = 'testsuite')
        root.appendChild(mainElement)
        
        # 每个sheet去构建DOM
        for  innerList in self.outerInfoList:
            testSuite1 = makeTestCase(dom, innerList[0])
            for dic in innerList[1:]:                                       # 内层数组，开头为name，后面为dict
                if  dic['isTestsuite']:
                    testSuite2 = makeTestCase(dom, dic['caseName'])         # excel中为合并空行的地方，新建testsuite
                    
                else:
                    summary = makeElement(dom, 'summary', dic['summary'])
                    preconditions = makeElement(dom, 'preconditions', dic['preconditions'])
                    steps = makeElement(dom, 'steps')
                    step = makeElement(dom, )
        
            
            mainElement.appendChild(testSuite1)
        
        # 生成文件
        f = open("name.xml", 'w', encoding="utf-8")
        dom.writexml(f, addindent='  ', newl='\n',encoding='utf-8')
        f.close()

    '''    
        
    def makeXML(self):    
    
        def makeTestCaseInfoFromDic(dic):
            f.write('<testcase name="' + dic['caseName'] + '">\n')
            f.write('   <summary><![CDATA[' + dic['summary'].replace('\n', '<br />') + ']]></summary>\n')
            f.write('   <preconditions><![CDATA[' + dic['preconditions'].replace('\n', '<br />') + ']]></preconditions>\n')
            f.write('   <importance><![CDATA[' + str(dic['importance']) + ']]></importance>\n')
            f.write('   <steps><step><step_number><![CDATA[1]]></step_number>\n')
            f.write('   <actions><![CDATA[' + str(dic['steps']).replace('\n', '<br />') + ']]></actions>\n')
            f.write('   <expectedresults><![CDATA['+ dic['expectedresults'].replace('\n', '<br />') +']]></expectedresults>\n')
            f.write('   <execution_type><![CDATA[1]]></execution_type>\n')
            f.write('</step></steps></testcase>\n')
            
        f = open(self.filePath.get() + '.xml', 'w', encoding='utf-8')
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write('<testsuite name="' + self.outPutName + '">\n')
        
        for innerInfoList in self.outerInfoList:
            f.write('<testsuite name="' + innerInfoList[0] + '" >\n')
            testsuiteNumber = 1
            for dic in innerInfoList[1:]:
                if dic['isTestsuite']:
                    f.write('<testsuite name="' + dic['caseName'] + '" >\n')
                    testsuiteNumber += 1
                else:
                    makeTestCaseInfoFromDic(dic)
            for i in range(testsuiteNumber):
                f.write('</testsuite>\n')
            
            
            
        f.write('</testsuite>\n')
        
        f.close()
        
    def readXlsx(self):
        '''
        从给定的地址，读取表格内容，每条测试用例存放一个dict，用一个list存放所有dict
        '''
        self.outerInfoList = []
        xlsxData = xlrd.open_workbook(self.filePath.get())          # 打开xlsx文件
        sheetsList = xlsxData.sheets()                              # 获取所有表格
        for sheet in sheetsList:                                    # 每个sheet
            # print('列数', sheet.ncols)
            # print('行数', sheet.nrows)
            if sheet.ncols < 9:                                     # 不符合格式的跳过
                continue
            innerInfoList = []                                      # innerInfoList存放每个sheet的测试用例，outerInfoList存放每个innerInfoList
            innerInfoList.append(sheet.name)                        # innerInfoList开头为testsuite的name
            for i in range(1,sheet.nrows):                          # 按行获取信息，每行是一个测试用例，跳过开始行
                dic = {}
                rowDataList = sheet.row_values(i)
                dic["caseName"] = rowDataList[0]
                dic["importance"] = rowDataList[1]
                dic["execution"] = rowDataList[2]
                dic["keywords"] = rowDataList[3]
                dic["summary"] = rowDataList[4]
                dic["preconditions"] = rowDataList[5]
                dic["steps"] = rowDataList[6]
                dic["expectedresults"] = rowDataList[7]
                dic["ExecutionType"] = rowDataList[8]
                dic["RequirementID"] = rowDataList[9]
                dic["isTestsuite"] = False
                
                # 判断是否新加testsuite
                if dic["summary"] == "":
                    dic["isTestsuite"] = True
                if dic["caseName"] != '':
                    innerInfoList.append(dic)
            
            self.outerInfoList.append(innerInfoList)
            
        
    
    def start(self):
        infoList = []
        index = self.filePath.get().rfind('/')
        self.outPutName = self.filePath.get()[index+1:].replace("xlsx","xml")  # 切片，并转换文件扩展名
        # print(self.outPutName)
        try:
            self.readXlsx()
            # 生成XML文件
            self.makeXML()
            messagebox.showinfo(title='生成成功',message='生成的XML文件在excel文件同目录下')
        except:
            messagebox.showinfo(title='生成失败',message='原因不详。\n请联系：liangyuyang@orbbec.com')
        
        
        
if __name__ == "__main__":
    xtx = XlsxToXML()