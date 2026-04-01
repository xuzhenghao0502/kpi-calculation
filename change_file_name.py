import os


def remove_string_from_filenames(directory, remove_string):
    """
    从指定目录下所有xlsx文件的文件名中去除特定字符串。

    :param directory: 文件夹路径
    :param remove_string: 需要去除的字符串
    """
    # 获取文件夹中的所有文件
    for filename in os.listdir(directory):
        if not filename.endswith('.xlsx'):
            continue
        # WPS：.~xxx.xlsx；Excel：~$xxx.xlsx
        if filename.startswith('.~') or filename.startswith('~$'):
            continue
        # 检查是否包含需要去除的字符串
        if remove_string in filename:
            # 构造新的文件名
            new_filename = filename.replace(remove_string, '')
            # 获取旧文件的完整路径和新文件的完整路径
            old_file = os.path.join(directory, filename)
            new_file = os.path.join(directory, new_filename)

            # 重命名文件
            os.rename(old_file, new_file)
            print(f"Renamed '{filename}' to '{new_filename}'")


# 使用示例
directory_path = '/home/zhenghaoxu@geometricalpal.com/PycharmProjects/pythonProject/2024Q3'  # 替换为你的目录路径
string_to_remove = '研发中心季度绩效考评表-'  # 替换为你想去除的字符串
remove_string_from_filenames(directory_path, string_to_remove)