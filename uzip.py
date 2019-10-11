import zipfile


def unzip_file(zip_src, dst_dir):
    r = zipfile.is_zipfile(zip_src)
    if r:
        fz = zipfile.ZipFile(zip_src, 'r')
        for file in fz.namelist():
            fz.extract(file, dst_dir)
    else:
        print('This is not zip')


if __name__ == '__main__':
    inputfilepath = r'C:\Users\siwei.yan\Desktop\res.zip'
    outputfilepath = r'C:\Users\siwei.yan\Desktop'
    unzip_file(inputfilepath, outputfilepath)

