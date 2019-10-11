import zipfile
from xml.dom import minidom
import xlwt
from bottle import *

HTML = """
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>定义input type="file" 的样式</title>
<style type="text/css">
body{
font-size:14px;
align-text:center;
}
input{ 
vertical-align:middle;
margin:0;
padding:0
}
.file-box{
position:relative;
width:340px;
margin:0px auto;
}
.txt{
height:22px;
border:1px solid #cdcdcd;
width:180px;
}
.btn{
background-color:#FFF;
border:1px solid #CDCDCD;
height:24px;
width:70px;
}
.file{
position:absolute;
top:0;
right:80px;
height:24px;
filter:alpha(opacity:0);
opacity: 0;
width:260px
}
</style>
</head>
<body>
<div class="file-box">
<form action="/upload" method="post" enctype="multipart/form-data">
<input type='text' name='textfield' id='textfield' class='txt' />  
<input type='button' class='btn' value='浏览...' />
<input type="file" name="fileField" class="file" id="fileField" size="28" onchange="document.getElementById('textfield').value=this.value" />
<input type="submit" name="submit" class="btn" value="上传" onclick=""/>
</form>
</div>
</body>
</html>
"""

DOWNHTML = """
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>定义input type="file" 的样式</title>
<style type="text/css">
body{
font-size:14px;
align-text:center;
}
input{ 
vertical-align:middle;
margin:0;
padding:0
}
.file-box{
position:relative;
width:340px;
margin:0px auto;
}
.txt{
height:22px;
border:1px solid #cdcdcd;
width:180px;
}
.btn{
background-color:#FFF;
border:1px solid #CDCDCD;
height:24px;
width:70px;
}
.file{
position:absolute;
top:0;
right:80px;
height:24px;
filter:alpha(opacity:0);
opacity: 0;
width:260px
}
</style>
</head>
<body>
<div class="file-box">
<form action="/download/result.xls" method="get" enctype="multipart/form-data"> 

<input type="submit" name="submit" class="btn" value="下载" onclick=""/>
</form>
</div>
</body>
</html>
"""

base_path = os.path.dirname(os.path.realpath(__file__))  # 获取脚本路径

upload_path = os.path.join(base_path, 'temp')  # 上传文件目录
if not os.path.exists(upload_path):
    os.makedirs(upload_path)


@route('/', method='GET')
@route('/upload', method='GET')
@route('/index.html', method='GET')
@route('/upload.html', method='GET')
def index():
    """显示上传页"""
    return HTML


@route('/upload', method='POST')
def do_upload():
    """处理上传文件"""
    filedata = request.files.get('fileField')

    if filedata.file:
        file_name = os.path.join(upload_path, filedata.filename)
        if os.path.exists(file_name):
            os.remove(file_name)
        try:
            filedata.save(file_name)  # 上传文件写入
        except IOError:
            return '上传文件失败'
        respath = base_path + '/res'
        if os.path.exists(respath):
            del_file(respath)
        unzip_file(upload_path + '/res.zip', base_path)
        xml2excel(respath)
        return DOWNHTML
        # return '上传文件成功, 文件名: {}'.format(file_name)
    else:
        return '上传文件失败'


@route('/download/<filename>')
def download(filename):
    return static_file(filename, root=base_path, download=filename)


@route('/favicon.ico', method='GET')
def server_static():
    """处理网站图标文件, 找个图标文件放在脚本目录里"""
    return static_file('favicon.ico', root=base_path)


@error(404)
def error404(error):
    """处理错误信息"""
    return '404 发生页面错误, 未找到内容'


def get_new_ip():
    cmd_file = os.popen('ifconfig')
    cmd_result = cmd_file.read()
    pattern = re.compile(r'inet[ ](\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3})')
    ip_list = re.findall(pattern, cmd_result)
    return ip_list[0]

def del_file(path):
    for i in os.listdir(path):
        path_file = os.path.join(path, i)
        if os.path.isfile(path_file):
            os.remove(path_file)
        else:
            del_file(path_file)


def unzip_file(zip_src, dst_dir):
    r = zipfile.is_zipfile(zip_src)
    if r:
        fz = zipfile.ZipFile(zip_src, 'r')
        for file in fz.namelist():
            fz.extract(file, dst_dir)
    else:
        print('This is not zip')


def init_xls(sheetname):
    font = xlwt.Font()
    font.name = 'SimSun'
    style = xlwt.XFStyle()
    style.font = font

    fxls = xlwt.Workbook(encoding='utf-8')
    table = fxls.add_sheet(sheetname, cell_overwrite_ok=True)

    return fxls, table


def read_string_xml(xmlname):
    sdict = {}
    doc = minidom.parse(xmlname)
    stringtaglen = len(doc.getElementsByTagName('string'))
    itemlen = len(doc.getElementsByTagName('item'))

    for idx in range(0, stringtaglen):
        nameattr = doc.getElementsByTagName('string')[idx].getAttribute('name')
        fc = doc.getElementsByTagName('string')[idx].firstChild
        if fc is None:
            content = ""
        else:
            content = fc.data
        sdict[nameattr] = content

    for idx in range(0, itemlen):
        nameattr = doc.getElementsByTagName('item')[idx].getAttribute('name')
        fc = doc.getElementsByTagName('item')[idx].firstChild
        if fc is None:
            content = ""
        else:
            content = fc.data
        sdict[nameattr] = content
    return sdict


def xml2excel(inputfilepath):
    xls, table = init_xls('Language')
    table.write(0, 0, 'ResourceName')
    dir_list = os.listdir(inputfilepath)
    list = []
    dict = {}
    i = 1
    for file in dir_list:
        if file.__contains__('values'):
            filename = inputfilepath + '/' + file + '/' + 'strings.xml'
            if not os.path.exists(filename):
                continue
            resourcedict = read_string_xml(filename)
            for resource in resourcedict:
                if resource not in list:
                    list.append(resource)
                    dict[resource] = i
                    table.write(i, 0, resource)
                    i = i + 1
    k = 1
    for file in dir_list:
        if file.__contains__('values'):
            filename = inputfilepath + '/' + file + '/' + 'strings.xml'
            if not os.path.exists(filename):
                continue
            resourcedict = read_string_xml(filename)
            colname = file[7:]
            if colname == "":
                colname = "default"
            table.write(0, k, colname)
            for resource in resourcedict:
                table.write(dict[resource], k, resourcedict[resource])
            k = k + 1
    outputfilepath = base_path + '/result.xls'
    if os.path.exists(outputfilepath):
        os.remove(outputfilepath)
    xls.save(outputfilepath)


run(port=8080, reloader=False)  # reloader设置为True可以在更新代码时自动重载
