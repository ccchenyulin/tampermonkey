// ==UserScript==
// @name         豆包表格转CSV/Excel下载器（精准标题修复版）
// @namespace    https://github.com/yourname/
// @version      3.2
// @description  精准识别豆包markdown表格标题（跨容器查找），交互式弹窗导出CSV/Excel
// @author       Your Name
// @match        https://www.doubao.com/*
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @grant        GM_download
// @grant        GM_addStyle
// @license      MIT
// ==/UserScript==

(function() {
    'use strict';

    // ==================== 核心配置 ====================
    const TABLE_SELECTOR = 'table.markdown-table, table';
    const BUTTON_ID = 'doubao-table-exporter-btn';
    const MODAL_ID = 'doubao-table-exporter-modal';
    const TITLE_SELECTOR = 'h1, h2, h3, h4, h5, h6'; // 支持所有标题标签
    let tableDataCache = null;
    let throttleTimer = null;

    // ==================== 样式定义 ====================
    GM_addStyle(`
        #${BUTTON_ID} {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 9998;
            padding: 10px 15px;
            background: #409eff;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 14px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            opacity: 0.9;
            transition: opacity 0.3s;
        }
        #${BUTTON_ID}:hover {
            opacity: 1;
        }
        #${MODAL_ID} {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            z-index: 9999;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .modal-content {
            width: 450px;
            background: white;
            border-radius: 8px;
            padding: 25px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
        }
        .modal-title {
            font-size: 18px;
            font-weight: 600;
            margin: 0 0 20px 0;
            color: #333;
        }
        .table-select-group {
            margin-bottom: 20px;
        }
        .table-select-group label {
            display: block;
            padding: 10px 12px;
            border: 1px solid #e6e6e6;
            border-radius: 4px;
            margin-bottom: 8px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .table-select-group label:hover {
            border-color: #409eff;
            background: #f5faff;
        }
        .table-select-group input[type="radio"]:checked + span {
            color: #409eff;
            font-weight: 500;
        }
        .table-select-group input[type="radio"] {
            margin-right: 8px;
        }
        .table-desc {
            font-size: 12px;
            color: #999;
            margin-left: 22px;
            margin-top: 2px;
        }
        .format-select-group {
            display: flex;
            gap: 10px;
            margin-bottom: 25px;
        }
        .format-select-group label {
            flex: 1;
            text-align: center;
            padding: 10px 0;
            border: 1px solid #e6e6e6;
            border-radius: 4px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .format-select-group label:hover {
            border-color: #409eff;
        }
        .format-select-group input[type="radio"]:checked + span {
            color: #409eff;
            font-weight: 500;
        }
        .modal-btn-group {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
        }
        .modal-btn {
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
        }
        .modal-btn.confirm-btn {
            background: #409eff;
            color: white;
            border: 1px solid #409eff;
        }
        .modal-btn.confirm-btn:hover {
            background: #337ecc;
        }
        .modal-btn.cancel-btn {
            background: white;
            border: 1px solid #e6e6e6;
        }
        .modal-btn.cancel-btn:hover {
            border-color: #d9d9d9;
            background: #f5f5f5;
        }
    `);

    // ==================== 工具函数 ====================
    function throttle(func, delay = 500) {
        return function(...args) {
            if (!throttleTimer) {
                func.apply(this, args);
                throttleTimer = setTimeout(() => {
                    throttleTimer = null;
                }, delay);
            }
        };
    }

    function escapeCsvCell(text) {
        if (typeof text !== 'string') text = String(text);
        text = text.replace(/"/g, '""');
        return (text.includes(',') || text.includes('\n') || text.includes('"')) ? `"${text}"` : text;
    }

    /**
     * 修复版标题查找：跨容器查找离表格最近的上方标题
     * 原理：通过元素的页面位置（top值）判断，找到表格上方最近的标题
     */
    function getTableTitle(table) {
        // 获取表格在页面中的垂直位置
        const tableRect = table.getBoundingClientRect();
        const tableTop = tableRect.top + window.scrollY; // 加上滚动偏移，避免滚动影响

        // 获取页面中所有标题元素
        const allTitles = document.querySelectorAll(TITLE_SELECTOR);
        let closestTitle = null;
        let minDistance = Infinity;

        allTitles.forEach(title => {
            const titleRect = title.getBoundingClientRect();
            const titleTop = titleRect.top + window.scrollY;
            const titleBottom = titleTop + titleRect.height;

            // 只选择在表格上方的标题（标题底部 < 表格顶部）
            if (titleBottom < tableTop) {
                const distance = tableTop - titleBottom;
                // 找到距离表格最近的标题
                if (distance < minDistance) {
                    minDistance = distance;
                    closestTitle = title;
                }
            }
        });

        // 返回找到的标题或默认值
        return closestTitle ? closestTitle.textContent.trim() : `无标题表格`;
    }

    /**
     * 提取表格数据（含修复版标题提取）
     */
    function extractTables() {
        if (tableDataCache) return tableDataCache;

        const tables = document.querySelectorAll(TABLE_SELECTOR);
        const tableDataList = [];

        tables.forEach((table, index) => {
            const tableTitle = getTableTitle(table); // 调用修复版函数

            const header = [...table.querySelectorAll('thead th, tr th')].map(th => escapeCsvCell(th.textContent.trim()));
            const rows = [...table.querySelectorAll('tbody tr, tr:not(:has(th))')].map(tr =>
                [...tr.querySelectorAll('td')].map(td => escapeCsvCell(td.textContent.trim()))
            ).filter(row => row.length > 0);

            if (rows.length > 0) {
                tableDataList.push({
                    id: index + 1,
                    title: tableTitle,
                    header: header.length ? header : Array.from({ length: rows[0].length }, (_, i) => `列${i + 1}`),
                    rows: rows,
                    rowCount: rows.length,
                    colCount: header.length || rows[0].length
                });
            }
        });

        tableDataCache = tableDataList;
        return tableDataList;
    }

    function generateCsv(tableData) {
        return [tableData.header.join(','), ...tableData.rows.map(row => row.join(','))].join('\n');
    }

    function csvToExcelAndDownload(csvContent, fileName) {
        try {
            const csvBlob = new Blob([`\uFEFF${csvContent}`], { type: 'text/csv;charset=utf-8;' });
            const reader = new FileReader();
            reader.onload = function(e) {
                const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
                const a = document.createElement('a');
                a.href = URL.createObjectURL(new Blob([excelBuffer], { type: 'application/octet-stream' }));
                a.download = `${fileName}.xlsx`;
                a.click();
                URL.revokeObjectURL(a.href);
            };
            reader.readAsArrayBuffer(csvBlob);
        } catch (err) {
            alert(`Excel转换失败：${err.message}`);
        }
    }

    function downloadCsv(csvContent, fileName) {
        try {
            const blob = new Blob([`\uFEFF${csvContent}`], { type: 'text/csv;charset=utf-8;' });
            if (typeof GM_download === 'function') {
                GM_download({ url: URL.createObjectURL(blob), name: `${fileName}.csv`, saveAs: true });
            } else {
                const a = document.createElement('a');
                a.href = URL.createObjectURL(blob);
                a.download = `${fileName}.csv`;
                a.click();
                URL.revokeObjectURL(a.href);
            }
        } catch (err) {
            alert(`CSV下载失败：${err.message}`);
        }
    }

    /**
     * 创建交互式导出弹窗
     */
    function createExportModal() {
        const oldModal = document.getElementById(MODAL_ID);
        if (oldModal) oldModal.remove();

        const tableDataList = extractTables();
        if (tableDataList.length === 0) {
            alert('当前页面未识别到有效表格！');
            return;
        }

        const modal = document.createElement('div');
        modal.id = MODAL_ID;
        modal.innerHTML = `
            <div class="modal-content">
                <h3 class="modal-title">选择要导出的表格</h3>
                <div class="table-select-group">
                    ${tableDataList.map((table, idx) => `
                        <label>
                            <input type="radio" name="tableSelect" value="${idx}" ${idx === 0 ? 'checked' : ''}>
                            <span>${table.title}</span>
                        </label>
                        <div class="table-desc">${table.rowCount} 行 ${table.colCount} 列</div>
                    `).join('')}
                </div>

                <h3 class="modal-title">选择导出格式</h3>
                <div class="format-select-group">
                    <label>
                        <input type="radio" name="formatSelect" value="csv" checked>
                        <span>CSV 格式</span>
                    </label>
                    <label>
                        <input type="radio" name="formatSelect" value="xlsx">
                        <span>Excel 格式</span>
                    </label>
                </div>

                <div class="modal-btn-group">
                    <button class="modal-btn cancel-btn">取消</button>
                    <button class="modal-btn confirm-btn">确认导出</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);

        // 绑定事件
        modal.querySelector('.cancel-btn').addEventListener('click', () => modal.remove());
        modal.querySelector('.confirm-btn').addEventListener('click', () => {
            const selectedTableIdx = modal.querySelector('input[name="tableSelect"]:checked').value;
            const selectedFormat = modal.querySelector('input[name="formatSelect"]:checked').value;
            const targetTable = tableDataList[selectedTableIdx];

            const timeStr = new Date().toLocaleString().replace(/[/:\s]/g, '_');
            const fileName = `豆包表格_${targetTable.title}_${timeStr}`;
            const csvContent = generateCsv(targetTable);

            selectedFormat === 'csv' ? downloadCsv(csvContent, fileName) : csvToExcelAndDownload(csvContent, fileName);
            modal.remove();
        });

        modal.addEventListener('click', (e) => {
            if (e.target === modal) modal.remove();
        });
    }

    /**
     * 创建导出按钮
     */
    function createExportButton() {
        if (document.getElementById(BUTTON_ID)) return;
        const button = document.createElement('button');
        button.id = BUTTON_ID;
        button.textContent = '导出表格';
        button.onclick = () => {
            tableDataCache = null;
            createExportModal();
        };
        document.body.appendChild(button);
    }

    // ==================== 初始化逻辑 ====================
    function init() {
        createExportButton();
        const throttledObserver = throttle(() => {
            if (!document.getElementById(BUTTON_ID)) createExportButton();
            tableDataCache = null;
        }, 1000);

        const contentContainer = document.querySelector('.content, .main-content, .article-content') || document.body;
        new MutationObserver(throttledObserver).observe(contentContainer, { childList: true, subtree: true });
    }

    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        init();
    } else {
        window.addEventListener('DOMContentLoaded', init);
    }

})();
