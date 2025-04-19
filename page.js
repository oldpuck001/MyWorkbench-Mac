// page.js

const { ipcRenderer } = require('electron');

let file_path = '';

export async function page_1() {
    const contentDiv = document.getElementById('content');
    contentDiv.innerHTML = `<h1>演示页面（page.js）</h1>`;
}