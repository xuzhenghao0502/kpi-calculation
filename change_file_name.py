import os


def remove_string_from_filenames(directory, strings_to_remove):
    """
    从指定目录下所有 .xlsx 文件的文件名中，依次去除列表中出现的每个子串。

    :param directory: 文件夹路径
    :param strings_to_remove: 需要去除的字符串列表（按顺序逐个 replace；空串会被忽略）
    """
    if isinstance(strings_to_remove, str):
        strings_to_remove = [strings_to_remove]

    for filename in os.listdir(directory):
        old_path = os.path.join(directory, filename)
        if not os.path.isfile(old_path):
            continue
        if not filename.endswith(".xlsx"):
            continue
        # WPS：.~xxx.xlsx；Excel：~$xxx.xlsx
        if filename.startswith(".~") or filename.startswith("~$"):
            continue

        new_filename = filename
        for s in strings_to_remove:
            if not s:
                continue
            new_filename = new_filename.replace(s, "")

        if new_filename == filename or not new_filename:
            continue

        new_path = os.path.join(directory, new_filename)
        if os.path.exists(new_path):
            print(f"跳过（目标已存在）: '{filename}' -> '{new_filename}'")
            continue

        os.rename(old_path, new_path)
        print(f"Renamed '{filename}' to '{new_filename}'")


# 使用示例
directory_path = "/home/zhenghao/Program/kpi-calculation/data/2026Q1-staff/26Q1slam绩效自评/"
strings_to_remove = [
    "2026Q1绩效考评表-", 
    "研发中心季度绩效考评表-",
    "2026研发中心Q1季度绩效考评表",
    "附件三.研发中心季度绩效考评表-专业线-",
    "附件三.",
]
remove_string_from_filenames(directory_path, strings_to_remove)
