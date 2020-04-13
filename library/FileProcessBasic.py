import abc
import sys

class FileProcessBasic(abc.ABC):
    '所有文件处理类的基类'

    # 查找关键字
    # @abc.abstractmethod
    # def findKeywords(self):
    #     raise NotImplementedError
    #
    # # 寻找上下文
    # @abc.abstractmethod
    # def findContent(self):
    #     raise NotImplementedError
    #
    # # 后续的处理
    # @abc.abstractmethod
    # def subsequentProcess(self):
    #     raise NotImplementedError

    @abc.abstractmethod
    def save(self, output):
        raise NotImplementedError

    @abc.abstractmethod
    def run(self, ouputpath, inputpath):
        raise NotImplementedError
