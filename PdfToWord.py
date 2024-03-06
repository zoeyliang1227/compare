import PySimpleGUI as sg

from pdf2docx import Converter

def pdf2word(file_path):
    file_name = file_path.split('.')[0]
    doc_file = f'{file_name}.docx'
    p2w = Converter(file_path)
    p2w.convert(doc_file, start=0, end=None)
    p2w.close()
    return doc_file

def main():
    # 選擇主題
     sg.theme('LightBlue5')
     # 設定視窗
     layout = [
         [sg.Text('pdfToword', font=('微軟雅黑', 12)),
          sg.Text('', key='filename', size=(50, 1), font=('微軟雅黑', 10), text_color='blue')],
         [sg.Output(size=(80, 10), font=('微軟雅黑', 10))],
         [sg.FilesBrowse('選擇檔案', key='file', target='filename'), sg.Button('開始轉換'), sg.Button('退出')]]
     # 建立視窗
     window = sg.Window("Python與資料分析_青青", layout, font=("微軟雅黑", 15), default_element_size=(50, 1))
     # 事件循環
     while True:
         # 視窗的讀取，有兩個回傳值（1.事件；2.值）
         event, values = window.read()
         print(event, values)

         if event == "開始轉換":
             # 單一文件
             if values['file'] and values['file'].split('.')[1] == 'pdf':
                 filename = pdf2word(values['file'])
                 print('檔案個數 ：1')
                 print('\n' + '轉換成功！' + '\n')
                 print('檔案保存位置：', filename)
             # 多個文件
             elif values['file'] and values['file'].split(';')[0].split('.')[1] == 'pdf':
                 print('檔個數 ：{}'.format(len(values['file'].split(';'))))
                 for f in values['file'].split(';'):
                     filename = pdf2word(f)
                     print('\n' + '轉換成功！' + '\n')
                     print('檔案保存位置：', filename)
             else:
                 print('請選擇pdf格式的檔案哦!')
         if event in (None, '退出'):
             break

     window.close()
     
if __name__ == "__main__":
    main()