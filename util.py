from win32com.client import Dispatch

def doc2docx(file):
    try:
        word = Dispatch('Word.Application')
        doc = word.Documents.Open(file)
        doc.SaveAs("{}x".format(file), 12)
        doc.Close()
        word.Quit()
    except IOError:
        print("读取文件异常：" + file)