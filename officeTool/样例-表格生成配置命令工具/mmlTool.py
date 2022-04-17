# -!- coding: utf-8 -!-
"""
create UDS disaster recovery related mml config.
version:v1.0
version information:
    python 3.9.5
    xlrd3 1.1.0
Created by LC on 2021.5.18
"""

"""
默认情况读取脚本同级目录下的"容灾配置信息采集表.xlsx"文件。
直接输入文件名（不带后缀）时，默认搜索脚本所在目录。
可以带路径+文件名，文件名无需加后缀，比如：”D:\AA\BB\CC.xlsx“，需要直接输入“D:\AA\BB\CC“。
生成名为"容灾配置"的目录,生成在输入文件名同级目录中。
"""

import xlrd3
import os
import copy

class SiteInfo(object):
    """UDS disaster recovery related information"""
    nfid = 0
    # 0-not main site, 1-main site
    mainSiteFlag = 0
    siteId = 0
    udsSyncIp = 0
    clusterNum = 0
    smscId = 0
    scNo = 0
    # The key is ClusterId, value is "clusterInfo".
    # clusterInfo:the first element(clusterInfo[0]) is the number of nodes,the rest is node ID.
    clusterCollection = {}

    def _init_(self):
        self.nfid = 0
        self.mainSiteFlag = 0
        self.siteId = 0
        self.udsSyncIp = 0
        self.clusterNum = 0
        self.clusterCollection = {}

    def cleanDate(self):
        self.nfid = 0
        self.mainSiteFlag = 0
        self.siteId = 0
        self.udsSyncIp = 0
        self.clusterNum = 0
        self.smscId = 0
        self.scNo = 0
        self.clusterCollection = {}

    def setNfid(self, inNfId):
        self.nfid = inNfId

    def getNfId(self):
        return self.nfid

    def setSmscid(self, inSmscid):
        self.smscId = inSmscid

    def getSmscid(self):
        return self.smscId

    def setScNo(self, inScNo):
        self.scNo = inScNo

    def getScNo(self):
        return self.scNo

    def setMainSiteFlag(self, inMainSiteFlag):
        self.mainSiteFlag = inMainSiteFlag

    def getMainSiteFlag(self):
        return self.mainSiteFlag

    def setSiteId(self, inSiteId):
        self.siteId = inSiteId

    def getSiteId(self):
        return self.siteId

    def setUdsSyncIp(self, inUdsSyncIp):
        self.udsSyncIp = inUdsSyncIp

    def getUdsSyncIp(self):
        return self.udsSyncIp

    def setCluster(self, inClusterId, inNodeId):
        if(inClusterId == "" or inNodeId == ""):
            return False
        if(inClusterId in self.clusterCollection):
            self.clusterCollection[inClusterId].append(inNodeId)
            self.clusterCollection[inClusterId][0] += 1
        elif(self.clusterNum == 0):
            self.clusterCollection = {inClusterId:[1, inNodeId]}
            self.clusterNum += 1
        else:
            self.clusterCollection[inClusterId] = [1, inNodeId]
            self.clusterNum += 1
        return True

    def getClusterNum(self):
        return self.clusterNum

    def getAllClusterId(self):
        return self.clusterCollection.keys()

    def getClusterAllNodeId(self, inClusterId):
        if(inClusterId in self.clusterCollection):
            return self.clusterCollection[inClusterId]
        else:
            return 0

def generateMML(inSiteInfo, mmlPath):
    """Generate mml commend for uds."""
    try:
        #mmlPath = os.getcwd()
        mmlPath = mmlPath + "容灾配置"

        siteIndex = 0
        mainSiteNum = 0
        tmpSiteInfo = SiteInfo()
        siteNum = len(inSiteInfo)
        while(siteIndex < siteNum):
            if("是" == inSiteInfo[siteIndex].getMainSiteFlag()):
                if(1 == mainSiteNum):
                    print("错误：主局数量不能大于一")
                    return False

                tmpSiteInfo = inSiteInfo[0]
                inSiteInfo[0] = inSiteInfo[siteIndex]
                inSiteInfo[siteIndex] = tmpSiteInfo
                mainSiteNum = 1
            siteIndex += 1

        if(2 != siteIndex):
            print("错误：搭建容灾只能有2个局")
            return False

        if(0 == mainSiteNum):
            print("错误：主局数量不能为0")
            return False

        if(inSiteInfo[0].getClusterNum() != inSiteInfo[1].getClusterNum()):
            print("错误：主备局集群数量不一致")
            return False

        for cid in inSiteInfo[0].getAllClusterId():
            if(0 == inSiteInfo[1].getClusterAllNodeId(cid)):
                print("错误：主备局集群ID不一致")
                return False

        #生成UDS局间通信配置
        mmlInitUdsPath = mmlPath + "\\" + "UDS局间通信配置"
        mmlfileMasterInit = mmlInitUdsPath + "\\" + "1_主局容灾配置.txt"
        mmlfileSlaverInit = mmlInitUdsPath + "\\" + "2_备局容灾配置.txt"

        if os.path.exists(mmlfileMasterInit):
            os.remove(mmlfileMasterInit)
        if os.path.exists(mmlfileSlaverInit):
            os.remove(mmlfileSlaverInit)
        if not os.path.exists(mmlInitUdsPath):
            os.makedirs(mmlInitUdsPath)

        mmlFile = open(mmlfileMasterInit, 'a')
        mmlFile.write("SET UDSTENANT:TENANTID=1,TENANT_NAME=tenant_1,NFID={0:.0f},NFID2={1:.0f},SCHEMANAME=nef\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[1].getNfId()))
        mmlFile.write("SET UDSFUNC:TENANTID=1,AUTOFAILBACK=\"YES\"\n")
        mmlFile.write("ADD NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=self_nf,COMMTYPE=NULL,NFSTATE=NORMAL\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[0].getNfId()))
        mmlFile.write("ADD NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=diff_tmsp_nf,COMMTYPE=NULL,NFSTATE=NORMAL\n"\
                      .format(inSiteInfo[1].getNfId(), inSiteInfo[1].getNfId()))
        mmlFile.write("ADD NFNET:ID=1,IPTYPE=\"IPV4\",IP=\"{0}\",VPNID=0\n".format(inSiteInfo[0].getUdsSyncIp()))
        mmlFile.write("ADD NFSEED:ID=1,NFID={0:.0f},TYPE=\"CUDR_DMCC\",IPTYPE=\"IPV4\",IP=\"{1}\",PORT=60001\n"\
                      .format(inSiteInfo[1].getNfId(), inSiteInfo[1].getUdsSyncIp()))
        mmlFile.write("SET DISASTERTOLERANT:DTTYPE=\"MAIN_DT\",DTSITEID={0:.0f}\n"\
                      .format(inSiteInfo[1].getSiteId()))
        mmlFile.write("SET SMSCINFO:SMSC_ID={0:.0f},SC_NO={1:.0f},STATUS=\"ACTIVE\"\n"\
                      .format(inSiteInfo[0].getSiteId(), inSiteInfo[0].getScNo()))
        mmlFile.write("SYNA:STYPE=\"ALL\",INCREF=\"NO\"\n")
        mmlFile.close()

        mmlFile = open(mmlfileSlaverInit, 'a')
        mmlFile.write("SET UDSTENANT:TENANTID=1,TENANT_NAME=tenant_1,NFID={0:.0f},NFID2={1:.0f},SCHEMANAME=nef\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[1].getNfId()))
        mmlFile.write("SET UDSFUNC:TENANTID=1,AUTOFAILBACK=\"YES\"\n")
        mmlFile.write("ADD NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=self_nf,COMMTYPE=NULL,NFSTATE=NORMAL\n"\
                      .format(inSiteInfo[1].getNfId(), inSiteInfo[1].getNfId()))
        mmlFile.write("ADD NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=diff_tmsp_nf,COMMTYPE=NULL,NFSTATE=NORMAL\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[0].getNfId()))
        mmlFile.write("ADD NFNET:ID=1,IPTYPE=\"IPV4\",IP=\"{0}\",VPNID=0\n".format(inSiteInfo[1].getUdsSyncIp()))
        mmlFile.write("ADD NFSEED:ID=1,NFID={0:.0f},TYPE=\"CUDR_DMCC\",IPTYPE=\"IPV4\",IP=\"{1}\",PORT=60001\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[0].getUdsSyncIp()))
        mmlFile.write("SET DISASTERTOLERANT:DTTYPE=\"STANDBY_DT\",DTSITEID={0:.0f}\n"\
                      .format(inSiteInfo[0].getSiteId()))
        mmlFile.write("SET SMSCINFO:SMSC_ID={0:.0f},SC_NO={1:.0f},STATUS=\"INACTIVE\"\n"\
                      .format(inSiteInfo[1].getSiteId(), inSiteInfo[1].getScNo()))
        mmlFile.write("SYNA:STYPE=\"ALL\",INCREF=\"NO\"\n")
        mmlFile.close()

        #生成回归独立局
        mmltoSinglePath = mmlPath + "\\" + "回归独立局"
        mmlfileMasterToSingle = mmltoSinglePath + "\\" + "1_主局回归独立配置.txt"
        mmlfileSlaverToSingle = mmltoSinglePath + "\\" + "2_备局回归独立配置.txt"

        if os.path.exists(mmlfileMasterToSingle):
            os.remove(mmlfileMasterToSingle)
        if os.path.exists(mmlfileSlaverToSingle):
            os.remove(mmlfileSlaverToSingle)
        if not os.path.exists(mmltoSinglePath):
            os.makedirs(mmltoSinglePath)

        mmlFile = open(mmlfileMasterToSingle, 'a')
        mmlFile.write("SET NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=self_nf,COMMTYPE=NULL,NFSTATE=BLOCK\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[0].getNfId()))
        mmlFile.write("SET NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=diff_tmsp_nf,COMMTYPE=NULL,NFSTATE=BLOCK\n"\
                      .format(inSiteInfo[1].getNfId(), inSiteInfo[1].getNfId()))
        mmlFile.write("SET DISASTERTOLERANT:DTTYPE=\"NONE_DT\"\n")
        mmlFile.write("SET SMSCINFO:SMSC_ID={0:.0f},SC_NO={1:.0f},STATUS=\"ACTIVE\"\n"\
                      .format(inSiteInfo[0].getSiteId(), inSiteInfo[0].getScNo()))
        mmlFile.write("SYNA:STYPE=\"ALL\",INCREF=\"NO\"\n")
        mmlFile.close()

        mmlFile = open(mmlfileSlaverToSingle, 'a')
        mmlFile.write("SET NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=self_nf,COMMTYPE=NULL,NFSTATE=BLOCK\n"\
                      .format(inSiteInfo[1].getNfId(), inSiteInfo[1].getNfId()))
        mmlFile.write("SET NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=diff_tmsp_nf,COMMTYPE=NULL,NFSTATE=BLOCK\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[0].getNfId()))
        mmlFile.write("SET DISASTERTOLERANT:DTTYPE=\"NONE_DT\"\n")
        mmlFile.write("SET SMSCINFO:SMSC_ID={0:.0f},SC_NO={1:.0f},STATUS=\"ACTIVE\"\n"\
                      .format(inSiteInfo[1].getSiteId(), inSiteInfo[1].getScNo()))
        mmlFile.write("SYNA:STYPE=\"ALL\",INCREF=\"NO\"\n")
        mmlFile.close()

        #生成无需清库时从双单局恢复双局同步
        mmlSingle2DdoublePath = mmlPath + "\\" + "无需清库时从双单局恢复双局同步"
        mmlfileMasterToDouble = mmlSingle2DdoublePath + "\\" + "1_主局恢复双局配置.txt"
        mmlfileSlaverToDouble = mmlSingle2DdoublePath + "\\" + "2_备局恢复双局配置.txt"

        if os.path.exists(mmlfileMasterToDouble):
            os.remove(mmlfileMasterToDouble)
        if os.path.exists(mmlfileSlaverToDouble):
            os.remove(mmlfileSlaverToDouble)
        if not os.path.exists(mmlSingle2DdoublePath):
            os.makedirs(mmlSingle2DdoublePath)

        mmlFile = open(mmlfileMasterToDouble, 'a')
        mmlFile.write("SET UDSTENANT:TENANTID=1,TENANT_NAME=tenant_1,NFID={0:.0f},NFID2={1:.0f},SCHEMANAME=nef\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[1].getNfId()))
        mmlFile.write("SET UDSFUNC:TENANTID=1,AUTOFAILBACK=\"YES\"\n")
        mmlFile.write("SET NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=self_nf,COMMTYPE=NULL,NFSTATE=NORMAL\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[0].getNfId()))
        mmlFile.write("SET NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=diff_tmsp_nf,COMMTYPE=NULL,NFSTATE=NORMAL\n"\
                      .format(inSiteInfo[1].getNfId(), inSiteInfo[1].getNfId()))
        mmlFile.write("SET DISASTERTOLERANT:DTTYPE=\"MAIN_DT\",DTSITEID={0:.0f}\n".format(inSiteInfo[1].getSiteId()))
        mmlFile.write("SET SMSCINFO:SMSC_ID={0:.0f},SC_NO={1:.0f},STATUS=\"ACTIVE\"\n"\
                      .format(inSiteInfo[0].getSiteId(), inSiteInfo[0].getScNo()))
        mmlFile.write("SYNA:STYPE=\"ALL\",INCREF=\"NO\"\n")
        mmlFile.close()

        mmlFile = open(mmlfileSlaverToDouble, 'a')
        mmlFile.write("SET UDSTENANT:TENANTID=1,TENANT_NAME=tenant_1,NFID={0:.0f},NFID2={1:.0f},SCHEMANAME=nef\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[1].getNfId()))
        mmlFile.write("SET UDSFUNC:TENANTID=1,AUTOFAILBACK=\"YES\"\n")
        mmlFile.write("SET NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=self_nf,COMMTYPE=NULL,NFSTATE=NORMAL\n"\
                      .format(inSiteInfo[1].getNfId(), inSiteInfo[1].getNfId()))
        mmlFile.write("SET NFCFG:NFID={0:.0f},NFNAME=UDS_{1:.0f},NFTYPE=diff_tmsp_nf,COMMTYPE=NULL,NFSTATE=NORMAL\n"\
                      .format(inSiteInfo[0].getNfId(), inSiteInfo[0].getNfId()))
        mmlFile.write("SET DISASTERTOLERANT:DTTYPE=\"STANDBY_DT\",DTSITEID={0:.0f}\n"\
                      .format(inSiteInfo[0].getSiteId()))
        mmlFile.write("SET SMSCINFO:SMSC_ID={0:.0f},SC_NO={1:.0f},STATUS=\"INACTIVE\"\n"\
                      .format(inSiteInfo[1].getSiteId(), inSiteInfo[1].getScNo()))
        mmlFile.write("SYNA:STYPE=\"ALL\",INCREF=\"NO\"\n")
        mmlFile.close()

        #生成主备切换
        mmlSwitchMasterPath = mmlPath + "\\" + "主备切换"
        mmlfileMaster2Slaver = mmlSwitchMasterPath + "\\" + "1_主局切换为备局.txt"
        mmlfileSlaver2Master = mmlSwitchMasterPath + "\\" + "2_备局切换为主局.txt"

        if os.path.exists(mmlfileMaster2Slaver):
            os.remove(mmlfileMaster2Slaver)
        if os.path.exists(mmlfileSlaver2Master):
            os.remove(mmlfileSlaver2Master)
        if not os.path.exists(mmlSwitchMasterPath):
            os.makedirs(mmlSwitchMasterPath)

        mmlFile = open(mmlfileMaster2Slaver, 'a')
        mmlFile.write("SET UDSTENANT:TENANTID=1,TENANT_NAME=tenant_1,NFID={0:.0f},NFID2={1:.0f},SCHEMANAME=nef\n"\
                      .format(inSiteInfo[1].getNfId(), inSiteInfo[0].getNfId()))
        mmlFile.write("SET DISASTERTOLERANT:DTTYPE=\"STANDBY_DT\",DTSITEID={0:.0f}\n"\
                      .format(inSiteInfo[1].getSiteId()))
        mmlFile.write("SET SMSCINFO:SMSC_ID={0:.0f},SC_NO={1:.0f},STATUS=\"INACTIVE\"\n"\
                      .format(inSiteInfo[0].getSiteId(), inSiteInfo[0].getScNo()))
        mmlFile.write("SYNA:STYPE=\"ALL\",INCREF=\"NO\"\n")
        mmlFile.close()

        mmlFile = open(mmlfileSlaver2Master, 'a')
        mmlFile.write("SET UDSTENANT:TENANTID=1,TENANT_NAME=tenant_1,NFID={0:.0f},NFID2={1:.0f},SCHEMANAME=nef\n"\
                      .format(inSiteInfo[1].getNfId(), inSiteInfo[0].getNfId()))
        mmlFile.write("SET DISASTERTOLERANT:DTTYPE=\"MAIN_DT\",DTSITEID={0:.0f}\n"\
                      .format(inSiteInfo[0].getSiteId()))
        mmlFile.write("SET SMSCINFO:SMSC_ID={0:.0f},SC_NO={1:.0f},STATUS=\"ACTIVE\"\n"\
                      .format(inSiteInfo[1].getSiteId(), inSiteInfo[1].getScNo()))
        mmlFile.write("SYNA:STYPE=\"ALL\",INCREF=\"NO\"\n")
        mmlFile.close()

        print("生成完毕。")
        return True
    except BaseException as err:
        print("{0}".format(err))
        return False

def exelMainFunc(filePathAndName = '容灾配置信息采集表'):
    """Reade xlsx file and get date"""
    try:
        if("\\" in filePathAndName):
            charIndex = filePathAndName.rfind("\\")
            charIndex += 1
            path = filePathAndName[0:charIndex]
            fileName = filePathAndName + '.xlsx'
        else:
            path = os.getcwd() + "\\"
            fileName = path + filePathAndName + '.xlsx'

        book = xlrd3.open_workbook(fileName)
        sheetsNum = book.nsheets
        sheetIndex = 0
        siteInfo = []
        siteNum = 0
        currentSite = SiteInfo()
        clusterColIndex = 0
        nodeColIndex = 0
        # sheet loop
        while(sheetIndex < sheetsNum):
            currentSheet = book.sheet_by_index(sheetIndex)
            rowNum = currentSheet.nrows
            colNum = currentSheet.ncols
            rowIndex = 0
            colIndex = 0
            colNameList = ["集群ID","节点ID","NFID","站点标识","UDS同步网路平面IP","是否主局","网元ID","绑定模块号"]
            # column loop
            while(colIndex < colNum):
                currentColName = currentSheet.cell_value(rowx=rowIndex, colx=colIndex)
                if(currentColName == '集群ID'):
                    try:
                        colNameList.remove("集群ID")
                        clusterColIndex = colIndex
                    except BaseException as exelMainFuncErr:
                        print("{0}--忽略此列：集群ID".format(exelMainFuncErr))
                    colIndex += 1
                    continue
                elif(currentColName == '节点ID'):
                    try:
                        colNameList.remove("节点ID")
                        nodeColIndex = colIndex
                    except BaseException as exelMainFuncErr:
                        print("{0}--忽略此列：节点ID".format(exelMainFuncErr))
                    colIndex += 1
                    continue
                elif(currentColName == 'vnf_instance_id'):
                    colIndex += 1
                    continue
                else:
                    rowIndex += 1
                    # row loop
                    while(rowIndex < rowNum):
                        currentValue = currentSheet.cell_value(rowx=rowIndex, colx=colIndex)
                        if(currentValue == ""):
                            rowIndex += 1
                            continue

                        if(currentColName == 'NFID'):
                            try:
                                colNameList.remove("NFID")
                                currentSite.setNfid(currentValue)
                            except BaseException as exelMainFuncErr:
                                print("{0}--忽略此列：NFID".format(exelMainFuncErr))
                            break
                        elif(currentColName == '站点标识'):
                            try:
                                colNameList.remove("站点标识")
                                currentSite.setSiteId(currentValue)
                            except BaseException as exelMainFuncErr:
                                print("{0}--忽略此列：站点标识".format(exelMainFuncErr))
                            break
                        elif(currentColName == 'UDS同步网路平面IP'):
                            try:
                                colNameList.remove("UDS同步网路平面IP")
                                currentSite.setUdsSyncIp(currentValue)
                            except BaseException as exelMainFuncErr:
                                print("{0}--忽略此列：UDS同步网路平面IP".format(exelMainFuncErr))
                            break
                        elif(currentColName == '是否主局'):
                            try:
                                colNameList.remove("是否主局")
                                currentSite.setMainSiteFlag(currentValue)
                            except BaseException as exelMainFuncErr:
                                print("{0}--忽略此列：是否主局".format(exelMainFuncErr))
                            break
                        elif(currentColName == '网元ID'):
                            try:
                                colNameList.remove("网元ID")
                                currentSite.setSmscid(currentValue)
                            except BaseException as exelMainFuncErr:
                                print("{0}--忽略此列：网元ID".format(exelMainFuncErr))
                            break
                        elif(currentColName == '绑定模块号'):
                            try:
                                colNameList.remove("绑定模块号")
                                currentSite.setScNo(currentValue)
                            except BaseException as exelMainFuncErr:
                                print("{0}--忽略此列：绑定模块号".format(exelMainFuncErr))
                            break
                        else:
                            print("未知的列名:'{0}'".format(currentColName))
                            break
                        rowIndex += 1

                    if(rowIndex == rowNum and currentValue == ""):
                        if(currentColName == 'NFID'):
                            print("错误：NFID未填写!")
                            return
                        elif(currentColName == '站点标识'):
                            print("错误：站点标识未填写!")
                            return
                        elif(currentColName == 'UDS同步网路平面IP'):
                            print("错误：UDS同步网路平面未填写!")
                            return
                        elif(currentColName == '是否主局'):
                            print("错误：是否主局未填写!")
                            return
                        elif(currentColName == '网元ID'):
                            print("错误：网元ID未填写!")
                            return
                        elif(currentColName == '绑定模块号'):
                            print("错误：绑定模块号未填写!")
                            return
                        else:
                            print("告警：未知的列名:'{0}'".format(currentColName))

                    rowIndex = 0
                colIndex += 1
            colIndex = 0

            sheetIndex += 1

            for colName in colNameList:
                print("错误：{0}号表格列名\"{1}\"没有出现".format(sheetIndex, colName))
                return

            # cluster info fill in
            rowIndex += 1
            currentClusterId = 0
            currentNodeId = 0
            while(rowIndex < rowNum):
                currentValue = currentSheet.cell_value(rowx=rowIndex, colx=clusterColIndex)
                if(currentValue != ""):
                    currentClusterId = currentValue
                elif(0 == currentClusterId):
                    continue

                currentValue = currentSheet.cell_value(rowx=rowIndex, colx=nodeColIndex)
                if(currentValue != ""):
                    currentNodeId = currentValue
                    if(False == currentSite.setCluster(currentClusterId, currentNodeId)):
                        print('错误：读取集群信息失败，集群ID或者节点ID为空。')
                        return

                rowIndex += 1
            rowIndex = 0

            if(rowNum != 0 and colNum != 0):
                siteInfo.append(copy.deepcopy(currentSite))
            currentSite.cleanDate()

        sheetIndex = 0

        if(False == generateMML(siteInfo, path)):
            print("错误：生成MML命令错误。")

        return

    except FileNotFoundError as err:
        print("错误：找不到文件 %s" %(fileName)) #old string formatting
        return
    except BaseException as err:
        print("{0}".format(err))
        return

try:
    while True:
        userInpute = input("\n输入 'Q' or 'q' 退出, 输入 's' or 'S' 用默认文件名开始执行。\n\
直接输入要处理的文件名或路径加文件名。文件名为.xlsx格式，输入文件名无需带后缀。\n\
结果输出到输入文件同目录\n")

        if(userInpute == "q" or userInpute == "Q"):
            break
        elif(userInpute == "s" or userInpute == "S"):
            exelMainFunc()
        else:
            exelMainFunc(userInpute)
except  BaseException as err:
    print("{0}".format(err))
