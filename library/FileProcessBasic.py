import abc

class FileProcessBasic(abc.ABC):
    """所有文件处理类的基类"""

    @abc.abstractmethod
    def save(self, output, record):
        raise NotImplementedError

    @abc.abstractmethod
    def run(self, ouputpath, inputpath):
        raise NotImplementedError
