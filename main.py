import os
from typing import List

from prompt_toolkit.validation import Validator

from cli import promote_input_dialog, promote_file_selection_dialog, promote_success_dialog, promote_fail_dialog
from processor import process_file, save_file

TITLE = "东B大学考试排期处理程序"


def scan_directory(directory_path: str = '.') -> List[str]:
    result = []
    for _, _, FILES in os.walk(directory_path):
        for name in FILES:
            if name.endswith(".xlsx"):
                result.append(name)

    return result


def is_directory(path: str = "."):
    return os.path.isdir(path)


if __name__ == '__main__':
    input_directory = promote_input_dialog("请输入存放xlsx文件的目录,默认为.", default=".",
                                           validator=Validator.from_callable(is_directory, "非法输入"))
    files = scan_directory(input_directory)
    input_filename = promote_file_selection_dialog(files)
    if input_filename is None:
        quit(0)
    default_output_filename = f"巡考_{input_filename}"
    output_filename = promote_input_dialog(f"请输入保存的文件名,默认:{default_output_filename}", default=f"{default_output_filename}")

    try:
        full_input_path = os.path.join(input_directory, input_filename)
        full_output_path = os.path.join('.', output_filename)
        print("正在处理...")
        data = process_file(full_input_path)
        save_file(data, full_output_path)
        promote_success_dialog(output_filename)
    except Exception as e:
        promote_fail_dialog(str(e))
