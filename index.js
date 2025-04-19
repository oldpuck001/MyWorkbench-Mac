// index.js

import { page_1 } from './page.js';

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
    page_1();
});

window.one = function() {
    page_1();
}

window.two = function() {
    const contentDiv = document.getElementById('content');
    contentDiv.innerHTML = `<h1>演示页面（index.js）</h1>`;
}