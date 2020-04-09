# coding=utf-8
import abc
import sys

class FileProcessBasic(abc.ABC):
    '所有文件处理类的基类'
    
    # 查找关键字
    @abc.abstractmethod
    def findKeywords(self):
        pass
    
    # 寻找上下文
    @abc.abstractmethod
    def findContent(self):
        pass

    # 后续的处理
    @abc.abstractmethod
    def subsequentProcess(self):
        pass

    @abc.abstractmethod
    def run(self,ouputpath,inputpath):
       pass

class FileProcess_GPR_S1S2(FileProcessBasic):
    'GPR_S1S2文件处理类'

    # 查找关键字
    def findKeywords(self):
        pass
    
    # 寻找上下文
    def findContent(self):
        pass

    # 后续的处理
    def subsequentProcess(self):
        pass

    def run(self,inputpath,ouputpath):
        print("hello world \t"+inputpath+"\t"+ouputpath)

class FileProcess_GPR_S3S4(FileProcessBasic):
    'GPR_S3S4文件处理类'

    # 查找关键字
    def findKeywords(self):
        pass
    
    # 寻找上下文
    def findContent(self):
        pass

    # 后续的处理
    def subsequentProcess(self):
        pass

    def run(self,inputpath,ouputpath):
        print("hello world \t"+inputpath+"\t"+ouputpath)

class FileProcess_ZZM_S1S2(FileProcessBasic):
    '掌子面_S1S2文件处理类'

    # 查找关键字
    @abc.abstractmethod
    def findKeywords(self):
        pass
    
    # 寻找上下文
    @abc.abstractmethod
    def findContent(self):
        pass

    # 后续的处理
    @abc.abstractmethod
    def subsequentProcess(self):
        pass

class FileProcess_ZZM_S3S4(FileProcessBasic):
    '掌子面_S3S4文件处理类'

    # 查找关键字
    @abc.abstractmethod
    def findKeywords(self):
        pass
    
    # 寻找上下文
    @abc.abstractmethod
    def findContent(self):
        pass

    # 后续的处理
    @abc.abstractmethod
    def subsequentProcess(self):
        pass

class FileProcess_TSP_S1S2(FileProcessBasic):
    'TSP_S1S2文件处理类'

    # 查找关键字
    @abc.abstractmethod
    def findKeywords(self):
        pass
    
    # 寻找上下文
    @abc.abstractmethod
    def findContent(self):
        pass

    # 后续的处理
    @abc.abstractmethod
    def subsequentProcess(self):
        pass

class FileProcess_TSP_S3S4(FileProcessBasic):
    'TSP_S3S4文件处理类'

    # 查找关键字
    @abc.abstractmethod
    def findKeywords(self):
        pass
    
    # 寻找上下文
    @abc.abstractmethod
    def findContent(self):
        pass

    # 后续的处理
    @abc.abstractmethod
    def subsequentProcess(self):
        pass

class FileProcess_Progress_S1S2(FileProcessBasic):
    '进度_S1S2文件处理类'

    # 查找关键字
    @abc.abstractmethod
    def findKeywords(self):
        pass
    
    # 寻找上下文
    @abc.abstractmethod
    def findContent(self):
        pass

    # 后续的处理
    @abc.abstractmethod
    def subsequentProcess(self):
        pass

class FileProcess_Progress_S3S4(FileProcessBasic):
    '进度_S3S4文件处理类'

    # 查找关键字
    @abc.abstractmethod
    def findKeywords(self):
        pass
    
    # 寻找上下文
    @abc.abstractmethod
    def findContent(self):
        pass

    # 后续的处理
    @abc.abstractmethod
    def subsequentProcess(self):
        pass

class FileProcess_Change_S1S2(FileProcessBasic):
    '变更_S1S2文件处理类'

    # 查找关键字
    @abc.abstractmethod
    def findKeywords(self):
        pass
    
    # 寻找上下文
    @abc.abstractmethod
    def findContent(self):
        pass

    # 后续的处理
    @abc.abstractmethod
    def subsequentProcess(self):
        pass

class FileProcess_Change_S3S4(FileProcessBasic):
    '变更_S3S4文件处理类'

    # 查找关键字
    @abc.abstractmethod
    def findKeywords(self):
        pass
    
    # 寻找上下文
    @abc.abstractmethod
    def findContent(self):
        pass

    # 后续的处理
    @abc.abstractmethod
    def subsequentProcess(self):
        pass