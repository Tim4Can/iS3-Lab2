import os
import docx
import config


inputFilePath = '/Users/budi/Desktop/未命名文件夹/源数据'
outputFilePath = '/Users/budi/Desktop/未命名文件夹'

def Process(inputFilePath, outputFilePath):
    # 读取文件夹路径下所有文件名
    fileList = os.listdir(inputFilePath)
    print("The total folder include:")
    print(fileList)
    # 遍历文件夹选择合适的处理方法
    for _file in fileList:
        if "GPR" in _file or "掌子面" in _file or "TSP" in _file or "进度" in _file or "变更" in _file:
            print("Find target folder.The folder include:")
            newFilePath = inputFilePath + '/' + _file
            newFileList = os.listdir(newFilePath)
            print(newFileList)
            for subfolder in newFileList:
                if "S1" in subfolder or "S2" in subfolder:
                    if "GPR" in _file:
                        # 调用 GPR-S1S2 的处理方法
                        subFilePath = newFilePath + '/' + subfolder
                        subFileList = os.listdir(subFilePath)
                        print(subFileList)
                        for subfile in subFileList:
                            if ".doc" in subfile:
                                print(subFilePath + '/' + subfile)
                                # GPRS1S2 = config.FileProcess_GPR_S1S2()
                                # GPRS1S2.findKeywords()
                                break
                #     if "掌子面" in _file:
                #         # 调用 掌子面-S1S2 的处理方法
                #         print(".")
                #     if "TSP" in _file:
                #         # 调用 TSP-S1S2 的处理方法
                #         print(".")
                #     if "进度" in _file:
                #         # 调用 进度-S1S2 的处理方法
                #         print(".")
                #     if "变更" in _file:
                #         # 调用 变更-S1S2 的处理方法
                #         print(".")
                # if "S3" in subfolder or "S4" in subfolder:
                #     if "GPR" in _file:
                #         # 调用 GPR-S3S4 的处理方法
                #         print(".")
                #     if "掌子面" in _file:
                #         # 调用 掌子面-S3S4 的处理方法
                #         print(".")
                #     if "TSP" in _file:
                #         # 调用 TSP-S3S4 的处理方法
                #         print(".")
                #     if "进度" in _file:
                #         # 调用 进度-S3S4 的处理方法
                #         print(".")
                #     if "变更" in _file:
                #         # 调用 GPR-S3S4 的处理方法
                #         print(".")
    
    # 处理完毕文件的保存
    # document.save(outputFilePath)

Process(inputFilePath, outputFilePath)