// data_cleaning.js

const { ipcRenderer } = require('electron');

let file_path = '';

export async function data_cleaning() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>数据清洗</h1>
        
        <div class="import">
            <div>
                <label>源数据文件</label>
                <input id="file_path" type="text">
            </div>

            <div>
                <label for="sheetDropdown">选择工作表</label>
                <select id="sheetDropdown" name="sheetDropdown"></select>
            </div>

            <div>
                <label for="columnDropdown">选择要进行清洗的列</label>
                <select id="columnDropdown" name="columnDropdown"></select>
            </div>

            <div>
                <label for="cleaning_modeDropdown">选择清理模式</label>
                <select id="cleaning_modeDropdown" name="cleaning_modeDropdown">
                    <option value="remove_duplicates">删除重复行</option>
                    <option value="fill_missing_zero">填充缺失值为：0</option>
                    <option value="fill_missing_blank">填充缺失值为：<空白></option>
                    <option value="fill_missing_previous">填充缺失值为：重复上一行</option>
                    <option value="standardize_text">标准化文本（去除首尾空格及转换为小写英文字母）</option>
                    <option value="convert_data_str">将数据类型转换为字符型</option>
                    <option value="convert_data_int">将数据类型转换为整数型</option>
                    <option value="convert_data_float">将数据类型转换为浮点数型</option>
                    <option value="convert_data_date">将数据类型转换为时间日期类型</option>
                </select>
            </div>

            <div>
                <button id="selectButton">选择文件</button>
                <button id="cleanButton">数据清洗</button>
                <button id="exportButton">导出SQLite数据库</button>
            </div>
        </div>

        <div class="export">
            <div>
                <label>操作反馈：</label>
            </div>
            <div>
                <textarea id="result_output" rows="18"></textarea>
            </div>
        </div>
    `;

    var file_path_input = document.getElementById('file_path');
    file_path_input.classList.add('readonly');
    file_path_input.readOnly = true;

    const sheetDropdown = document.getElementById('sheetDropdown');
    const columnDropdown = document.getElementById('columnDropdown');
    const cleaning_modeDropdown = document.getElementById('cleaning_modeDropdown');

    // 导入按钮js代码
    document.getElementById('selectButton').addEventListener('click', async () => {
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        file_path = await ipcRenderer.invoke('dialog:openFile', fileFilters);
        if (!file_path) { return }
 
        ipcRenderer.send('asynchronous-message', { command: 'data_cleaning_select', data: {'file_path': file_path }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('sheetDropdown').addEventListener('change', async (event) => {
        const selected_sheet = event.target.value;

        ipcRenderer.send('asynchronous-message', { command: 'data_cleaning_index', data: {'file_path': file_path, 'sheet_name': selected_sheet }});
    });

    // 工作表选择的更改事件
    const triggerChange = (element) => {
        var event = new Event('change');
        element.dispatchEvent(event);
    };

    document.getElementById('cleanButton').addEventListener('click', async () => {

        file_path = file_path_input.value;
        if (!file_path) {
            alert('请先导入文件！');
            return;
        }
        
        const data = {
            file_path: file_path,
            sheet_name: sheetDropdown.value,
            column_name: columnDropdown.value,
            cleaning_mode: cleaning_modeDropdown.value
        };

        ipcRenderer.send('asynchronous-message', { command: 'data_cleaning_clean', data: data });
    });

    document.getElementById('exportButton').addEventListener('click', async () => {

        file_path = file_path_input.value;
        if (!file_path) {
            alert('请先导入文件！');
            return;
        }

        const data = {
            file_path: file_path,
            sheet_name: sheetDropdown.value,
        };

        ipcRenderer.send('asynchronous-message', { command: 'data_cleaning_export', data: data});
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');

    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'data_cleaning_select') {

            file_path_input.value = file_path;

            sheetDropdown.innerHTML = '';

            result[1].forEach(item => {
                const option = document.createElement('option');
                option.value = item;
                option.text = item;
                sheetDropdown.appendChild(option);
            });

            triggerChange(sheetDropdown);       // 触发 change 事件，自动加载列数据

            document.getElementById(`result_output`).value += '选择文件成功！\n';

        } else if (result[0] === 'data_cleaning_index') {
    
            columnDropdown.innerHTML = '';
    
            result[1].forEach(item => {
                const columnOption = document.createElement('option');
                columnOption.value = item;
                columnOption.text = item;
                columnDropdown.appendChild(columnOption);
            });

        } else if (result[0] === 'data_cleaning_clean') {

            document.getElementById(`result_output`).value += result[1];

        } else if (result[0] === 'data_cleaning_export') {

            document.getElementById(`result_output`).value += result[1];

        };
    });
}