import os


def fileName_replace(file_dir):
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if ".pyd" in file:
                new_file_name = str(file).split(".", 1)[0] + ".pyd"
                os.rename(file_dir + "\\%s" % file, file_dir + "\\%s" % new_file_name)


def delete_file(file_path):
    if (os.path.exists(file_path)):
        os.remove(file_path)


def folder_directory(file_dir):
    L = []
    for root, dirs, files in os.walk(file_dir):
        for file in dirs:
            if file != "__pycache__":
                L.append(file_dir + "\\" + file)
        return L


def file_name(file_dir):
    L = {}
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if file != "__init__.py":
                os.environ['file_path'] = file_dir + "\\" + file
                os.system('python "%s" build_ext --inplace' % "./setup.py")
                fileName_replace(file_dir)
                delete_file(file_dir + "\\" + file)
        return L


li = folder_directory(os.getcwd() + "\\Lbt")
# print(os.getcwd())
# li = ["C:\\Users\\wang984\\Desktop\\RPA\\rpa_client\\Lbt\\ACCESS"]
for item in li:
    file_name(item)
