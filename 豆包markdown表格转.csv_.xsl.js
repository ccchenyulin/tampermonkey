// ==UserScript==
// @name         豆包表格转CSV/Excel下载器
// @namespace    https://github.com/yourname/
// @version      2.0
// @description  识别豆包页面中的表格，导出为CSV并转换为Excel下载
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
    const BUTTON_ID = 'doubao-table-exporter-btn'; // 按钮唯一ID
    let tableDataCache = null; // 表格数据缓存
    let throttleTimer = null; // 节流定时器

    // ==================== 样式定义（改用GM_addStyle） ====================
    GM_addStyle(`
        #${BUTTON_ID} {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 9999;
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
    `);

    // ==================== 工具函数（优化性能） ====================
    /**
     * 节流函数：避免频繁触发
     * @param {Function} func 执行函数
     * @param {number} delay 延迟时间
     */
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

    /**
     * 转义CSV特殊字符
     */
    function escapeCsvCell(text) {
        if (typeof text !== 'string') text = String(text);
        text = text.replace(/"/g, '""');
        return (text.includes(',') || text.includes('\n') || text.includes('"')) ? `"${text}"` : text;
    }

    /**
     * 提取表格数据（带缓存）
     */
    function extractTables() {
        if (tableDataCache) return tableDataCache; // 优先使用缓存

        const tables = document.querySelectorAll(TABLE_SELECTOR);
        const tableDataList = [];
        tables.forEach((table, index) => {
            const header = [...table.querySelectorAll('thead th, tr th')].map(th => escapeCsvCell(th.textContent.trim()));
            const rows = [...table.querySelectorAll('tbody tr, tr:not(:has(th))')].map(tr =>
                [...tr.querySelectorAll('td')].map(td => escapeCsvCell(td.textContent.trim()))
            ).filter(row => row.length > 0);

            if (rows.length > 0) {
                tableDataList.push({
                    id: index + 1,
                    header: header.length ? header : Array.from({length: rows[0].length}, (_, i) => `列${i+1}`),
                    rows: rows
                });
            }
        });

        tableDataCache = tableDataList; // 缓存结果
        return tableDataList;
    }

    /**
     * 生成CSV
     */
    function generateCsv(tableData) {
        return [tableData.header.join(','), ...tableData.rows.map(row => row.join(','))].join('\n');
    }

    /**
     * CSV转Excel下载
     */
    function csvToExcelAndDownload(csvContent, fileName) {
        try {
            const csvBlob = new Blob([`\uFEFF${csvContent}`], { type: 'text/csv;charset=utf-8;' }); // BOM头解决中文乱码
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

    /**
     * 下载CSV（兼容不同油猴插件）
     */
    function downloadCsv(csvContent, fileName) {
        try {
            const blob = new Blob([`\uFEFF${csvContent}`], { type: 'text/csv;charset=utf-8;' });
            if (typeof GM_download === 'function') {
                GM_download({url: URL.createObjectURL(blob), name: `${fileName}.csv`, saveAs: true});
            } else {
                // 降级为原生下载
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
     * 表格选择弹窗
     */
    function showExportDialog() {
        const tableDataList = extractTables();
        if (tableDataList.length === 0) {
            alert('当前页面未识别到有效表格！');
            return;
        }

        const selectedIndex = prompt(
            `识别到 ${tableDataList.length} 张表格，请选择序号导出：\n${tableDataList.map(t => `表格${t.id}（${t.rows.length}行${t.header.length}列）`).join('\n')}\n\n输入序号（如1）：`,
            '1'
        );
        if (!selectedIndex) return;

        const targetTable = tableDataList[parseInt(selectedIndex) - 1];
        if (!targetTable) return alert('无效序号！');

        const fileName = `豆包表格_${new Date().toLocaleString().replace(/[/:\s]/g, '_')}`;
        const csvContent = generateCsv(targetTable);

        const exportType = prompt('选择导出格式：\n1 = CSV（推荐，无乱码）\n2 = Excel', '1');
        exportType === '1' ? downloadCsv(csvContent, fileName) : csvToExcelAndDownload(csvContent, fileName);
    }

    /**
     * 创建导出按钮（确保唯一）
     */
    function createExportButton() {
        if (document.getElementById(BUTTON_ID)) return; // 避免重复创建
        const button = document.createElement('button');
        button.id = BUTTON_ID;
        button.textContent = '导出表格';
        button.onclick = () => {
            tableDataCache = null; // 点击时清空缓存，确保提取最新表格
            showExportDialog();
        };
        document.body.appendChild(button);
    }

    // ==================== 初始化逻辑（优化监听） ====================
    // 页面加载完成后创建按钮
    function init() {
        createExportButton();
        // 监听DOM变化，但通过节流控制频率
        const throttledObserver = throttle(() => {
            if (!document.getElementById(BUTTON_ID)) createExportButton();
            tableDataCache = null; // DOM变化时清空缓存
        }, 1000); // 1秒内最多触发1次

        // 缩小监听范围：只监听豆包内容区域（需根据实际DOM调整，这里用常见的内容容器选择器）
        const contentContainer = document.querySelector('.content, .main-content, .article-content') || document.body;
        new MutationObserver(throttledObserver).observe(contentContainer, { childList: true, subtree: true });
    }

    // 确保在页面内容加载后执行
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        init();
    } else {
        window.addEventListener('DOMContentLoaded', init);
    }

})();