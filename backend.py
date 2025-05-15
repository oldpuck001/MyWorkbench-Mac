# backend.py

import sys
import json

from file_tools import modifythefilename
from file_tools import character
from file_tools import image
from file_tools import export
from file_tools import sort
from file_tools import collect_file
from file_tools import copy_folder

from xlsx_tools import splice

from audit_tools import select_folder




def main():
    
    # 從標準輸入讀取數據
    input_data = sys.stdin.read()

    # 將數據轉換為 Python 對象
    request = json.loads(input_data)

    # 處理數據
    if request["command"] == "filename_modify":
        result = modifythefilename.modify(request)
    elif request["command"] == "filename_character":
        result = character.character(request)
    elif request["command"] == "filename_image":
        result = image.image(request)
    elif request["command"] == "filename_export":
        result = export.export(request)
    elif request["command"] == "filename_sort":
        result = sort.sort(request)
    elif request["command"] == "collect_file":
        result = collect_file.collect_file(request)
    elif request["command"] == "copy_folder":
        result = copy_folder.copy_folder(request)

    elif request["command"] == "splice_sheet_input":
        result = splice.input_sheet(request)
    elif request["command"] == "splice_sheet_output":
        result = splice.output_sheet(request)

    elif request["command"] == "select_folder_path":
        result = select_folder.select_folder_path(request)



    else:
        result = "Unknown command"

    # 返回結果
    response = {"result": result}
    print(json.dumps(response))
    sys.stdout.flush()

if __name__ == "__main__":
    main()