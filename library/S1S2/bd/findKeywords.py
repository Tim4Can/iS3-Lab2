import docx
import os

def locateParagraph():
    doc = docx.Document('/Users/budi/Desktop/未命名文件夹/源数据/GPR地质预报/S1-S2标数据/001期-老营特长隧道 右幅进口K1+483-508.docx')
    # 定位6.2段落标签flag(此处写死，之后要改)
    flag = 0
    for p in doc.paragraphs:
        if "6.2" in p.text and flag == 0:
            flag = 1
            continue
        if flag == 0:
            continue
        if flag == 1:
            return p.text

para = locateParagraph()

def findKeywords():
    # 掌子面桩号
    zzmzh = ''
    for i in range (len(para)):
        if (para[i] == '+'):
            j = i
            while (para[j] != '（'):
                j = j - 1
            while (para[j+1] != '～'):
                zzmzh = zzmzh + (para[j+1])
                j = j + 1
            break
    print ("掌子面桩号：" + zzmzh)
    # 桩号区间
    zhqj = ''
    for i in range (len(para)):
        if (para[i] == '+'):
            j = i
            while (para[j] != '（'):
                j = j - 1
            while (para[j+1] != '）'):
                zhqj = zhqj + (para[j+1])
                j = j + 1
            break
    print ("桩号区间：" + zhqj)
findKeywords()