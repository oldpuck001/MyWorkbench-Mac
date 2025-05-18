// index.js

import { modifythefilenameFunction } from './file_tools/modifythefilename.js'
import { characterFunction } from './file_tools/character.js'
import { imageFunction } from './file_tools/image.js'
import { sortFunction } from './file_tools/sort.js'
import { exportFunction } from './file_tools/export.js'
import { collect_fileFunction } from './file_tools/collect_file.js'
import { copy_folderFunction } from './file_tools/copy_folder.js'

import { spliceFunction } from './xlsx_tools/splice.js';

import { select_folder } from './audit_tools/select_folder.js';
import { set_up } from './audit_tools/set_up.js';
import { import_account_balance_sheet } from './audit_tools/import_account_balance_sheet.js';
import { import_chronological_account } from './audit_tools/import_chronological_account.js';
import { import_balance_sheet } from './audit_tools/import_balance_sheet.js';
import { import_income_statement } from './audit_tools/import_income_statement.js';
import { import_cash_flow_statement } from './audit_tools/import_cash_flow_statement.js';


document.querySelectorAll('.sidebar > ul > li').forEach(item => {
    item.addEventListener('click', function (e) {
        // 检查是否有子菜单
        const sublist = this.querySelector('ul');
        if (sublist) {
            e.stopPropagation();
            sublist.classList.toggle('active');
        }
    });
});

document.querySelectorAll('.sidebar ul ul li').forEach(item => {
    item.addEventListener('click', function(e) {
        e.stopPropagation();
    });
});

document.addEventListener('DOMContentLoaded', () => {
    const sidebar = document.getElementById('sidebar');
    const toggleBtn = document.getElementById('toggleSidebarBtn');
    const content = document.getElementById('content');

    window.toggleSidebar = function () {
        if (sidebar.classList.contains('hidden')) {
            sidebar.classList.remove('hidden');
            sidebar.style.width = '250px';
            toggleBtn.style.left = '265px';
            content.style.marginLeft = '250px';
        } else {
            sidebar.classList.add('hidden');
            sidebar.style.width = '0';
            toggleBtn.style.left = '10px';
            content.style.marginLeft = '0';
        }
    };
});

// 页面加载后自动显示的页面
window.addEventListener('DOMContentLoaded', () => {
    const contentDiv = document.getElementById('content');
    contentDiv.innerHTML = `<h1>欢迎使用</h1>`;
});

window.file_tools_modifythefilenameFunction = function() {
    modifythefilenameFunction();
}

window.file_tools_characterFunction = function() {
    characterFunction();
}

window.file_tools_imageFunction = function() {
    imageFunction();
}

window.file_tools_exportFunction = function() {
    exportFunction();
}

window.file_tools_sortFunction = function() {
    sortFunction();
}

window.file_tools_collect_fileFunction = function() {
    collect_fileFunction();
}

window.file_tools_copy_folderFunction = function() {
    copy_folderFunction();
}

window.xlsx_tools_spliceFunction = function() {
    spliceFunction();
}

window.select_folder = function() {
    select_folder();
}

window.set_up = function() {
    set_up();
}

window.import_account_balance_sheet = function() {
    import_account_balance_sheet();
}

window.import_chronological_account = function() {
    import_chronological_account();
}

window.import_balance_sheet = function() {
    import_balance_sheet();
}

window.import_income_statement = function() {
    import_income_statement();
}

window.import_cash_flow_statement = function() {
    import_cash_flow_statement();
}