'''
因为word文件自带的PDF/A是第一代协议，不支持png的透明度，所以png图片需要额外把背景改成白色，不然默认PDF/A转换成黑色。
首先复制docx到tmp文件夹，然后后缀改成zip。改成zip是为了方便提取docx文件中的png图片(在word/media/文件夹下)。
找到zip文件夹中的png图片，把alpha channel <= 127 （规定阈值）的像素设成白色。
然后再把zip文件改回成docx。
之后用标准的转pdf流程转换。
'''
import tkinter as tk
from tkinter.filedialog import askopenfilename, askopenfilenames, askdirectory
from tkinter.messagebox import showinfo,showerror

from win32com.client import constants, gencache #Dispatch,DispatchEx
import zipfile
import cv2
import numpy as np
from ruamel.std.zipfile import delete_from_zip_file
#python自有库
from pathlib import Path
from functools import partial
import time
import threading
import os, shutil

def transparence2white(img):
    '''只考虑alpha channel为0的像素'''
    sp=img.shape  # 获取图片维度
    width=sp[0]  # 宽度
    height=sp[1]  # 高度
    for yh in range(height):
        for xw in range(width):
            color_d=img[xw,yh]  # 遍历图像每一个点，获取到每个点4通道的颜色数据
            if color_d.size != 4: #如果图片只有三个通道，也是可以正常处理
            	continue
            if color_d[3] == 0:  # 最后一个通道为透明度，如果其值为0，即图像是透明
                img[xw,yh]=[255,255,255,255]  # 则将当前点的颜色设置为白色，且图像设置为不透明
    return img

def transparence2white_v2(img):
    '''通过设定alpha channel的阈值来确定哪些像素为白色背景。阈值越高字体越细（白色部分越多），阈值越低字体周围会有黑框。'''
    if img.shape[2] == 3: #预防只有3通道的情况
        return img
    else:
        alpha_channel = img[:,:,3]
        if np.sum(alpha_channel != 255) == 0: #说明不透明
            return img
        else:
            _,mask = cv2.threshold(alpha_channel, 127, 255, cv2.THRESH_BINARY) #binarize mask, 高于阈值的都变成255
            color = img[:,:,:3]
            new_img = cv2.bitwise_not(cv2.bitwise_not(color, mask=mask))
            return new_img

def img2byte(img):
    success,encoded_image = cv2.imencode(".png",img)
    byte_data = encoded_image.tobytes()

    return byte_data

def transparentPNG2WhiteBackgroundinZip(tmp_path):
    '''
    把透明PNG改成白色背景，主函数。
    '''
    z = zipfile.ZipFile(tmp_path,'r')
    if not zipfile.Path(z,at='word/media/').exists(): #如果不存在说明没有图片
        z.close()
    else:
        zip_file_paths = []
        png_figures = []
        #不加下面一行的话会报错，是包的bug？
        z = zipfile.ZipFile(tmp_path,'r')
        for zip_file_path in z.namelist():
            if zip_file_path[-4:] == '.png':
                zip_file_paths.append(zip_file_path)
                png_figures.append(z.read(zip_file_path)) #bytes
        z.close()

        #如果有PNG才做
        if zip_file_paths:
            #避免文件重复，先要删除png
            delete_from_zip_file(tmp_path,file_names = zip_file_paths)

            z = zipfile.ZipFile(tmp_path,'a')
            for i, png_figure in enumerate(png_figures): 

                    img = cv2.imdecode(np.frombuffer(png_figure,np.uint8), cv2.IMREAD_UNCHANGED) #UNCHANGED会保留png的四通道
                    #DEBUG: 显示图片
            #         cv2.imshow("image",img)
            #         cv2.waitKey(0)
                    img = transparence2white_v2(img)
            #         cv2.imshow("image",img)
            #         cv2.waitKey(0)
                    byte_img = img2byte(img)
                    z.writestr(zip_file_paths[i], byte_img,compress_type=zipfile.ZIP_DEFLATED)

            z.close()

def selectPath():
    path_ = askopenfilenames()
    path_ = [p_.replace("/","\\\\") for p_ in path_]
    path.set(path_)

def selectDir():
    path_ = askdirectory()
    path_ = path_.replace("/","\\\\")
    savedir.set(path_)

def thread_it(func, *args):
    # 创建
    t = threading.Thread(target=func, args=args) 
    # 守护 !!!
    t.setDaemon(True) 
    # 启动
    t.start()
    # 阻塞--卡死界面！
    # t.join()

def delete_directory(path):
    for root, dirs, files in os.walk(path,topdown=False):                
        for filename in files:  
            abspath = os.path.join(root,filename)
            os.remove(abspath)
    os.rmdir(path)
    

def main(path,savedir):
    
    #防止button按几次
    button_file['state'] = tk.DISABLED
    button_dir['state'] = tk.DISABLED
    button_pdf['state'] = tk.DISABLED
    checkbutton_png['state'] = tk.DISABLED
    
    word_paths = eval(path.get())
    savedir_ = savedir.get()
    
    if allow_png_check.get():
    
        #判断文件格式是否正确
        doc_paths = []
        docx_paths = []
        for word_path in word_paths:
            filename = word_path.split('\\\\')[-1]
            if '.docx' == filename[-5:]:
                docx_paths.append(word_path)
            elif '.doc' == filename[-4:]:
                doc_paths.append(word_path)
            else:
                button_file['state'] = tk.NORMAL
                button_dir['state'] = tk.NORMAL
                button_pdf['state'] = tk.NORMAL
                checkbutton_png['state'] = tk.NORMAL
                progress.set('')
                root.update()
                showerror(title = "错误",
                  message = f"{filename}不是word文件！")
                return None

        #在当前路径下建tmp文件夹
        cwdpath = Path.cwd()
        tmp_dir = 'tmp'
        while Path(cwdpath/tmp_dir).exists():
            tmp_dir += 'p'
        Path(cwdpath/tmp_dir).mkdir(parents=True, exist_ok=False)

        #如果有doc文件，先要另存为docx，再执行png操作
        if doc_paths:
            #set progress bar
            progress.set('preparing doc...')
            root.update()

            gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
            wd = gencache.EnsureDispatch('Word.Application')
            wd.Visible = 0
            wd.DisplayAlerts = 0
            for doc_path in doc_paths:

                doc = wd.Documents.Open(doc_path)
                doc_filename = doc_path.split('\\\\')[-1][:-4] + '.docx'
                doc_sav_filepath = str(cwdpath/tmp_dir/doc_filename)
                doc.SaveAs(doc_sav_filepath,constants.wdFormatDocumentDefault)
                doc.Close(SaveChanges=constants.wdDoNotSaveChanges)
                #这里append到docx path list后面
                docx_paths.append(doc_sav_filepath)
            wd.Quit(constants.wdDoNotSaveChanges)


        tmp_docx_paths = []
        for doc in docx_paths:
            fpath, fname = os.path.split(doc)

            #set progress bar
            progress.set(f'checking PNG transparency for {fname}...')
            root.update()

            fzipname = fname[:-5] + '.zip'
            #把文件copy进tmp文件夹，后缀改成.zip。tmp_path为临时的文件路径。
            tmp_path = cwdpath/tmp_dir/fzipname
            shutil.copy(doc, tmp_path)

            transparentPNG2WhiteBackgroundinZip(tmp_path)

            #zip文件后缀改为docx
            fdocname = cwdpath/tmp_dir/fname
            shutil.move(tmp_path,fdocname)
            tmp_docx_paths.append(fdocname)
    else:
        tmp_docx_paths = [Path(wp) for wp in word_paths]
        

    gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
    wd = gencache.EnsureDispatch('Word.Application')
    wd.Visible = 0
    wd.DisplayAlerts = 0
    
    
    for docx_path in tmp_docx_paths:
        filename = docx_path.name
        #set progress bar
        progress.set('converting ' + filename + '...')
        #print(progress.get())
        root.update()

        if '.docx' == filename[-5:]:
            pdf_name = filename.lower()[:-5] + '.pdf'
        elif '.doc' == filename[-4:]:
            pdf_name = filename.lower()[:-4] + '.pdf'
        else:
            print('error!!')

        pdf_path = savedir_ + '\\\\'+ pdf_name

        doc = wd.Documents.Open(str(docx_path))
        s = wd.Selection
        s.WholeStory()
        #全选字体变黑色
        wd.Selection.Font.Color = 0

        for f in doc.Fields:
            #如果是Reference 域，字体变蓝色
            if f.Type == 3 or f.Type == 88: #wdFieldRef
                f.Select()
                wd.Selection.Font.Color = 16711680  #13209 brown #16711680 #wdColorBlue

        doc.ExportAsFixedFormat(pdf_path, constants.wdExportFormatPDF, Item=constants.wdExportDocumentWithMarkup,
                                CreateBookmarks=constants.wdExportCreateHeadingBookmarks, UseISO19005_1 = True)
        doc.Close(SaveChanges=constants.wdDoNotSaveChanges)
        
    wd.Quit(constants.wdDoNotSaveChanges)
    
    if allow_png_check.get():
        delete_directory(cwdpath/tmp_dir)
    
    progress.set('')
    button_file['state'] = tk.NORMAL
    button_dir['state'] = tk.NORMAL
    button_pdf['state'] = tk.NORMAL
    checkbutton_png['state'] = tk.NORMAL
    showinfo(title = "提示",
              message = f"已完成!")
    return None    



root = tk.Tk()
root.title("word转pdf工具v1.3   Author:Haonan Zhang")

path = tk.StringVar()
savedir = tk.StringVar()
progress = tk.StringVar()
allow_png_check = tk.IntVar(value = 1)

text1 = tk.Label(root,text='功能：\n   1.word批量转换pdf（保留有效的书签、超链接、目录导航、PDF/A格式）\n   2.批量文件重命名，所有大写\
字母转换成小写字母\n   3.文档中除了超链接之外的所有字体设置成黑色\n   4.把透明PNG的背景变成白色',
         justify='left')
text1.grid(row=0, column=1)
text_file = tk.Label(root, text="请指定文件：")
text_file.grid(row=1,column=0)
entry_file = tk.Entry(root,textvariable=path,width=80)
entry_file.grid(row=1,column=1)
button_file = tk.Button(root, text="选择",command=selectPath)
button_file.grid(row=1,column=2)
text_dir = tk.Label(root, text="请指定保存目录：")
text_dir.grid(row=2,column=0)
entry_dir = tk.Entry(root,textvariable=savedir,width=80)
entry_dir.grid(row=2,column=1)
button_dir = tk.Button(root, text="选择",command=selectDir)
button_dir.grid(row=2,column=2)
checkbutton_png = tk.Checkbutton(root, text="是否转换透明PNG图片的背景为白色（取消会快一些）", variable = allow_png_check, \
                                onvalue = 1, offvalue = 0, width = 80)
checkbutton_png.grid(row=3,column=1)
button_pdf = tk.Button(root, text="执行pdf转换",command=lambda: thread_it(main,path,savedir))
button_pdf.grid(row=4,column=1)
text_progress = tk.Label(root,textvariable=progress)
text_progress.grid(row=5,column=1)


root.mainloop()