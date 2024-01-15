from zipfile import ZIP_DEFLATED, ZipFile
import os


def recursive_files_list(dir):
    for path, _, files in os.walk(dir):
        for file in files:
            yield os.path.join(path, file)


def create_path(path: str):
    import os
    path = os.path.normpath(path)
    created_path = ''
    for part in path.split(os.sep):
        created_path = os.path.join(created_path, part)
        if not os.path.exists(created_path):
            os.mkdir(created_path)


def xlsx_to_folder(xlsx_filename, folder=None):
    if not folder:
        folder = xlsx_filename.rstrip('.xlsx')
    with ZipFile(xlsx_filename, 'r') as xlsx_file:
        for sub_filename in xlsx_file.namelist():
            path = os.path.join(
                folder,
                os.path.split(sub_filename)[0],
            )
            create_path(path)
            with open(os.path.join(folder, sub_filename), 'wb') as sub_file:
                sub_file.write(xlsx_file.read(sub_filename))


def folder_to_xlsx(folder: str, xlsx_filename: str = None):
    if not xlsx_filename:
        xlsx_filename = folder + '_output.xlsx'
    with ZipFile(xlsx_filename, 'w', compression=ZIP_DEFLATED) as output:
        for path in recursive_files_list(folder):
            with open(path, 'rb') as sub_file:
                xlsx_path = os.path.join(*path.split(os.sep)[1:])
                content = minimize_xml(sub_file.read())
                output.writestr(xlsx_path, content)


def minimize_xml(text: bytes):
    while True:
        if b'\n\t' in text:
            text = text.replace(b'\n\t', b'\n')
        else:
            break

    while True:
        if b'\n ' in text:
            text = text.replace(b'\n ', b'\n')
        else:
            break

    text = text.replace(b'\n<', b'<')
    text = text.replace(b'\n', b' ')

    return text


if __name__ == '__main__':
    xlsx_to_folder('example.xlsx')
    folder_to_xlsx('example')
