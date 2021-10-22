import win32api
import win32print
import os
import uuid
import platform
from flask import Flask, request, redirect, url_for, render_template
from werkzeug.utils import secure_filename

def PrintFile(file_path):
    allPrinter=[printer[2] for printer in win32print.EnumPrinters(2)]
    # 此命令可以看到你的所有打印机
    if "HP" not in allPrinter[4]:
        return False
    # PrintNum = int(input("选择打印机:\n"+"\n".join([f"{n} {p}" for n,p in enumerate(allPrinter)])))
    win32print.SetDefaultPrinter(allPrinter[4])
    # pdf_path="C:\\Users\Luke\Desktop\Doc1.docx"
    pdf_path=file_path
    win32api.ShellExecute(0,
                          "print",
                          pdf_path,
                        '/d:"%s"' % win32print.GetDefaultPrinter (),
                          ".",
                          0)
    return True

slash = '\\'
UPLOAD_FOLDER = 'upload'
ALLOW_EXTENSIONS = {'doc', 'docx', 'pdf'}
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
#判断文件夹是否存在，如果不存在则创建
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
else:
    pass
# 判断文件后缀是否在列表中
def allowed_file(filename):
    return '.' in filename and \
            filename.rsplit('.', 1)[1] in ALLOW_EXTENSIONS

@app.route('/',methods=['GET','POST'])
def upload_file():
    if request.method =='POST':
        #获取post过来的文件名称，从name=file参数中获取
        file = request.files['file']
        if file and allowed_file(file.filename):
            # secure_filename方法会去掉文件名中的中文
            filename = secure_filename(file.filename)
            #因为上次的文件可能有重名，因此使用uuid保存文件
            file_name = str(uuid.uuid4()) + '.' + filename.rsplit('.', 1)[1]
            file.save(os.path.join(app.config['UPLOAD_FOLDER'],file_name))
            base_path = os.getcwd()
            file_path = base_path + slash + app.config['UPLOAD_FOLDER'] + slash + file_name
            state=PrintFile(file_path)
            if state:
                msg="发送任务成功，等待打印\n"
            else:
                msg="未连接的HP打印机,请联系管理员\n"
            # return redirect(url_for('upload_file',filename = file_name))
            return render_template("index.html",msg=msg)
    else:
        msg=""
    return render_template('index.html',msg=msg)


if __name__ == "__main__":
    app.run(host='0.0.0.0',port=7001)#IP Port