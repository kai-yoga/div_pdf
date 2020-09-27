from pdfminer.high_level import extract_text
import os
from PyPDF2 import PdfFileReader,PdfFileWriter
import openpyxl as opl
import time
def get_dic(file_name):
    '''识别拆分文件页码范围'''
    text=extract_text(file_name,password='Gzdx230!@#$')
    text=repr(text).replace('\n','').split('广州大学')
    L=[]
    dic={}
    for i in range(len(text)):
        if len(text[i])>5:
            # print(i,text[i])
            xh=text[i].split('学 号')[-1].split('姓 名')[0].replace(' ','').replace('\\n','')
            xm=text[i].split('姓 名')[-1].split('年 级')[0].replace(' ','').replace('\\n','')
            nj=text[i].split('年 级')[-1].split('班 级')[0].replace(' ','').replace('\\n','')
            bj=text[i].split('班 级')[-1].split('\\n')[2].replace(' ','').replace('\\n','')
            lsh=text[i].split('打印流水号:')[-1].split('#')[0]
            if bj=='性质':
                print(text[i])
                # print(xh,xm,nj,bj,lsh)
            L.append(xh+'#'+xm+'#'+nj+'#'+bj+'#'+lsh)
    for index in range(len(L)):
        if L[index] not in dic.keys():
            dic[L[index]]=str(L.index(L[index]))+'-'+str(index+L.count(L[index])-1)
    return dic
#'1401100011#王燕萍#2016#地理161#06132022': '0-0',
# def mkdirs(savepath,dic={}):
#     '''创建文件保存目录'''
#     for k in dic.keys():
#         dir_keys=k.split('#')
#         dir=dir_keys[2]+'\\'+dir_keys[3]
#         if not os.path.exists(os.path.join(savepath,dir)):
#             os.makedirs(os.path.join(savepath,dir))
def create_file_save_path(savepath='',k=''):
    path=os.getcwd()
    if savepath and k:
        nj,bj=k.split('#')[2],k.split('#')[3]
        path=os.path.join(os.path.join(savepath,nj),bj)
        if not os.path.exists(path):
            os.makedirs(path)
    return path


def div_files(file_name,savepath,dic={}):
    '''按页码范围拆分文件'''
    pdf=PdfFileReader(open(file_name,'rb'))
    L=[('学号','姓名','年级','班级','文件位置')]
    if pdf.decrypt('Gzdx230!@#$'):
        # L=[]
        ########################拆分、加密文档
        for k,v in dic.items():
            password=k.split('#')[-1]
            # print(password)
            doc=PdfFileWriter()
            # if doc.decrypt('Gzdx230!@#$'):
            start_page,end_page=int(v.split('-')[0]),int((v.split('-')[-1]))
            for index in range(start_page,end_page+1):
                doc.addPage(pdf.getPage(index))
            doc.encrypt(user_pwd=password)
            ####保存拆分后的pdf文件
            xh,xm,nj,bj=k.split('#')[0],k.split('#')[1],k.split('#')[2],k.split('#')[3]
            save_file_name=xh+' '+xm+' '+nj+' '+bj+'.pdf'
            if os.path.exists(os.path.join(create_file_save_path(savepath,k),save_file_name)):
                os.remove(os.path.join(create_file_save_path(savepath,k),save_file_name))
            with open(os.path.join(create_file_save_path(savepath,k),save_file_name),'wb') as f:
                doc.write(f)
            f.close()
            L.append((xh,xm,nj,bj,
                      "=hyperlink('"+os.path.join(create_file_save_path(savepath,k),save_file_name)+"')")
                     )
            # print((xh,xm,nj,bj,
            #           '=hyperlink('+os.path.join(create_file_save_path(savepath,k),save_file_name)+')'))
            print('当前保存文件=',save_file_name,'完成!')
    return L

def create_xlsx_list(L,file_name=''):
    '''生成拆分的任务清单,保存在exe文件路径下'''
    xlsx_path=os.getcwd()
    if os.path.exists(os.path.join(xlsx_path,file_name.replace('.pdf','')+'拆分结果.xlsx')):
        os.remove(os.path.join(xlsx_path,file_name.replace('.pdf','')+'拆分结果.xlsx'))
    wb=opl.Workbook()
    ws=wb.active
    for row in L:
        ws.append(row)
    wb.save(os.path.join(xlsx_path,file_name.replace('.pdf','')+'拆分结果.xlsx'))
    return 1

def get_input():
    print('*'*30)
    print('请输入需要拆分的pdf文件夹路径\n如不输入，则需确保需要拆分的pdf文件与当前exe文件在同一文件夹下：\n输入后请按回车键：')
    path=input()
    # print('接收输入完成，请按回车键！')
    return path if len(path)>0 else os.getcwd()

def main():
    '''调用主函数'''
    savepath=os.path.join(os.getcwd(),'拆分后')
    sourpath=get_input()
    # sourpath=r"C:\Users\kai-y\Desktop\新建文件夹"
    print('程序已启动，请等待！')
    start=time.perf_counter()
    for file in os.listdir(sourpath):
        if file.lower().endswith('.pdf'):
            file_name=os.path.join(sourpath,file)
            dic=get_dic(file_name)
            L=div_files(file_name=file_name,savepath=savepath,dic=dic)
            if create_xlsx_list(L,file):
                print('处理完成！正在清理，请等待')
            time.sleep(10)
        else:
            print('文件=',file,'不是pdf文件，忽略处理!')
    end=time.perf_counter()
    print('处理完成，耗时=',str(end-start),'秒！')
if __name__=='__main__':
    main()