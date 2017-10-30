# coding=utf-8
# -*-coding : UTF-8-*-
import win32com
import os
from win32com.client import Dispatch, constants

# 判断文件是否存在
def get_file_content(file_path):
    if os.path.exists(file_path) is False:
        os.mkdir(file_path)

class DOC_Set:

    def __init__(self):
        return

    def Creat_DOC(self):
        global wordApp
        wordApp = win32com.client.Dispatch('Word.Application')

        # 后台运行，显示，不警告
        wordApp.Visible = True
        # wordApp.visible = False
        wordApp.DisplayAlerts = 0

        # 创建新的文档
        global doc
        doc= wordApp.Documents.Add()

        doc.PageSetup.PaperSize = 7  # A4
        global select
        select = wordApp.Selection

        pic = select.InlineShapes.AddPicture('D:\\11.jpg')

        # select.ParagraphFormat.Alignment = 0     # 段落对齐,0=左对齐,1=居中,2=右对齐
        # select.Font.Name="黑体"
        # select.Font.Name="宋体"
        # select.Font.Size = 24    #字号
        # select.Font.Bold = True    #粗体
        # doc.Paragraphs.Last.Range.Text = uart_data[0]  # 保存文件
        # doc.Paragraphs.Last.Range.Text = '  ' # 保存文件
        # doc.Paragraphs.Last.Range.Text = '  ' # 保存文件
        # select.ParagraphFormat.Alignment = 1

    def Set_Format(self, font, size, underline, bold, alignment):
        # 文档内容
        select.Font.Name = font # "微软雅黑"
        select.Font.Size = size # 30
        select.Font.Underline = underline # False
        select.Font.Bold = bold # True
        select.ParagraphFormat.Alignment = alignment # 1

    def Set_Text(self, enter_front, text):
        for num in range(0, enter_front):
            select.TypeText("\n")

        select.TypeText(text) # "测 试 报 告"

    def Read_Content(self, DOCPath):
        uart_fd = open(DOCPath, 'r')  # 'FreeBug.ini'
        global uart_data
        uart_data = uart_fd.readlines()
        uart_fd.close()

class Doc_Content:
    def __init__(self):
        return

    # 第一页
    def Page_1(self):
        doc_set.Set_Format ("微软雅黑", 26, False, True, 1)
        doc_set.Set_Text (2, '用电研发中心产品审查室')

        doc_set.Set_Format ("微软雅黑", 36, False, True, 1)
        doc_set.Set_Text (1, '测试报告')

        doc_set.Set_Format ("宋体", 16, False, False, 1)
        doc_set.Set_Text (1, '报告编号：')
        doc_set.Set_Text (0, uart_data[0][0:14])

        doc_set.Set_Format ("宋体", 16, False, False, 1)
        doc_set.Set_Text (5, '产品名称：')
        doc_set.Set_Text (0, uart_data[1][0:9])

        doc_set.Set_Format ("宋体", 16, False, False, 0)
        doc_set.Set_Text (1, '产品型号：')
        doc_set.Set_Text (0, uart_data[2][0:13])

        doc_set.Set_Format ("宋体", 16, False, False, 0)
        doc_set.Set_Text (1, '委托部门：')
        doc_set.Set_Text (0, uart_data[3][0:4])

        doc_set.Set_Format ("宋体", 16, False, False, 0)
        doc_set.Set_Text (1, '测试人：')
        doc_set.Set_Text (0, uart_data[4][0:2])

        doc_set.Set_Format ("宋体", 16, False, False, 0)
        doc_set.Set_Text (1, '测试日期：')
        doc_set.Set_Text (0, uart_data[5][0:10])

    # 第二页
    def Page_2(self):
        doc_set.Set_Format ("黑体", 22, False, False, 0)
        doc_set.Set_Text (1, '注意事项')
        doc_set.Set_Format ("宋体", 14, False, False, 1)
        doc_set.Set_Text (1, '')
        doc_set.Set_Format ("宋体", 14, False, False, 0)
        doc_set.Set_Text (1, '1.测试报告无检测人员、审核人员的签字无效。')
        doc_set.Set_Text (1, '2.测试报告涂改无效。')
        doc_set.Set_Text (1, '3.对测试报告若有异议，应于收到报告之日起15天内提出，逾期不接受异议。')
        doc_set.Set_Text (1, '4.测试报告只适用于被测试样表。')
        doc_set.Set_Text (1, '5.测试报告部分复制无效。')

    # 第三页
    def Page_3(self):
        doc_set.Set_Format ("黑体", 16, False, True, 0)
        doc_set.Set_Text (0, '产品审查室测试报告')
        doc_set.Set_Format ("宋体", 12, False, False, 1)
        Table = doc.add.add_table(rows=10,cols=5)



if __name__ == "__main__":

    get_file_content('report\\')

    # 文档设置
    doc_set =DOC_Set()
    doc_set.Creat_DOC()
    doc_set.Read_Content('FreeBug.ini')

    # 文档内容编写
    doc_content = Doc_Content()
    # 第一页
    doc_content.Page_1()
    # 第二页
    doc_content.Page_2 ()
    # 增加分页
    doc_set.Set_Text (15, '')
    # 第三页
    doc_content.Page_3 ()

    #文档保存、退出
    FileName = uart_data[0][0:14] + uart_data[1][0:9] + uart_data[2][0:13] + "测试报告.doc"
    FilePath = os.getcwd() + "\\report\\" + FileName
    doc.SaveAs(FilePath)
    # wordApp.Quit()
