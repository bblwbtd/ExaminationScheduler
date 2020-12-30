import os
from typing import List

from prompt_toolkit.shortcuts import input_dialog, radiolist_dialog, message_dialog
from prompt_toolkit.validation import Validator

TITLE = "东B大学考试排期处理程序"


def promote_input_dialog(text: str, default='', validator: Validator = None):
    try:
        text = input_dialog(title=TITLE, text=text, cancel_text="取消", ok_text="确定", validator=validator).run()
        print(text)
        if text == '' or text is None:
            text = default

        return text
    except:
        return default


def promote_file_selection_dialog(filenames: List[str]):
    return radiolist_dialog(title=TITLE, text="请选择被处理的xlsx文件", values=[(name, name) for name in filenames],
                            ok_text="确定",
                            cancel_text="取消").run()


def promote_success_dialog(output_path: str = ''):
    return message_dialog(title=TITLE, text=f"导出成功!\n文件路径{output_path}").run()


def promote_fail_dialog(error_message: str = ''):
    return message_dialog(title=TITLE, text=f"导出失败!\n错误输出{error_message}").run()
