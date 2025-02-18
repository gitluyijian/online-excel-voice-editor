let recognition = null;
let filepath = null;
let originalFilename = null;
let currentCell = 0; // 当前列索引
let timer; // 定时器
let isProcessingCommand = false; // 新增状态标志
let autoMoveEnabled = true; // 新增状态标志
let restart = true;//是否程序自动重启语音
let buttonStop = false;

// 文件上传函数
function uploadFile() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];

    if (!file) {
        alert('请选择文件');
        return;
    }

    if (!file.name.match(/\.(xlsx|xls)$/)) {
        alert('请选择Excel文件(.xlsx或.xls)');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    fetch('/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        console.log('Received data:', data);
        if (data.error) {
            alert(data.error);
            return;
        }
        filepath = data.filepath;
        originalFilename = data.filename;
        displayExcelData(data);
    })
    .catch(error => {
        console.error('Error:', error);
        alert('上传失败: ' + error.message);
    });
}

// 显示Excel数据
function displayExcelData(data) {
    console.log('Displaying data:', data);
    const table = document.getElementById('excelTable');
    table.innerHTML = '';

    // 检查数据格式
    if (!data || !Array.isArray(data.columns) || data.columns.length === 0) {
        console.error('Invalid data format or missing columns:', data);
        alert('Excel文件必须包含表头');
        return;
    }

    const columns = data.columns;
    let displayData = Array.isArray(data.data) ? data.data : [];

    // 如果数据为空，则创建5行空数据
    if (displayData.length === 0) {
        displayData = Array(5).fill().map(() => Array(columns.length).fill(''));
    }

    // 创建表头，分为两行：
    // 第一行为字母提示（A, B, C, ...）
    // 第二行为原始Excel表头
    const thead = document.createElement('thead');

    // 第一行：字母提示
    const headerRow1 = document.createElement('tr');
    // 第一单元格为行号（跨两行显示）
    const rowNumberHeader = document.createElement('th');
    rowNumberHeader.textContent = '#';
    rowNumberHeader.rowSpan = 2;
    headerRow1.appendChild(rowNumberHeader);

    columns.forEach((column, index) => {
        const th = document.createElement('th');
        th.textContent = getColumnLabel(index);
        th.className = 'column-letter-header'; // 可在CSS中进行单独样式设置
        headerRow1.appendChild(th);
    });
    thead.appendChild(headerRow1);

    // 第二行：原始Excel表头
    const headerRow2 = document.createElement('tr');
    columns.forEach((column) => {
        const th = document.createElement('th');
        th.textContent = column;
        th.className = 'column-header'; // 原有的列头样式
        headerRow2.appendChild(th);
    });
    thead.appendChild(headerRow2);
    table.appendChild(thead);

    // 创建表体
    const tbody = document.createElement('tbody');
    displayData.forEach((rowData, rowIndex) => {
        const tr = document.createElement('tr');

        // 添加行号
        const rowNumberCell = document.createElement('td');
        rowNumberCell.textContent = (rowIndex + 1).toString();
        rowNumberCell.className = 'row-number';
        tr.appendChild(rowNumberCell);

        // 确保每行数据与列数匹配
        const normalizedRowData = Array(columns.length).fill('');
        if (Array.isArray(rowData)) {
            rowData.forEach((value, index) => {
                if (index < columns.length) {
                    normalizedRowData[index] = value;
                }
            });
        }

        // 添加数据单元格
        normalizedRowData.forEach((cellData, colIndex) => {
            const td = document.createElement('td');
            td.textContent = cellData !== null ? cellData : '';
            td.contentEditable = true;
            td.dataset.row = rowIndex + 1;
            td.dataset.col = getColumnLabel(colIndex);
            td.className = 'data-cell';

            // 添加单元格坐标提示，如 A1, B1 等
            td.title = `${td.dataset.col}${td.dataset.row}`;

            td.addEventListener('focus', () => {
                currentCell = td;
                highlightCurrentPosition(td); // 确保高亮当前单元格
                td.style.backgroundColor = 'green !important'; // 设置背景色为绿色
            });

            td.addEventListener('blur', () => {
                removeHighlight();
                td.style.backgroundColor = ''; // 恢复背景色
                //saveChanges();
            });

            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
}

// 获取列标签（A, B, C, ..., Z, AA, AB, ...）
function getColumnLabel(index) {
    let label = '';
    let num = index;

    while (num >= 0) {
        label = String.fromCharCode(65 + (num % 26)) + label;
        num = Math.floor(num / 26) - 1;
    }

    return label;
}

// 高亮当前位置
function highlightCurrentPosition(cell) {
    removeHighlight();

    const row = cell.dataset.row;
    const col = cell.dataset.col;

    // 高亮行号
    const rowNumberCell = document.querySelector(`.row-number:nth-child(1):nth-of-type(${row})`);
    if (rowNumberCell) rowNumberCell.classList.add('highlight');

    // 高亮列标题
    const colHeader = document.querySelector(`th.column-header:nth-child(${parseInt(cell.cellIndex) + 1})`);
    if (colHeader) colHeader.classList.add('highlight');

    // 设置当前单元格的背景色为绿色
    cell.style.backgroundColor = 'green !important'; // 确保当前单元格背景色为绿色

    // 更新状态显示
    updatePositionStatus(cell);
}

// 移除高亮
function removeHighlight() {
    document.querySelectorAll('.highlight').forEach(el => {
        el.classList.remove('highlight');
    });
}

// 更新位置状态显示
function updatePositionStatus(cell) {
    const statusElement = document.getElementById('positionStatus');
    if (statusElement) {
        statusElement.textContent = `当前位置: ${cell.dataset.col}${cell.dataset.row}`;
    }
}

// 获取当前表格数据的函数（只取第二行作为表头）
function getCurrentTableData() {
    const table = document.getElementById('excelTable');
    if (!table) return null;

    // 只取第二行的表头（原始Excel的表头）
    const headerRow = table.querySelector("thead tr:nth-child(2)");
    if (!headerRow) return null;
    const headers = Array.from(headerRow.querySelectorAll('th')).map(th => th.textContent.trim());

    // 获取所有数据行，并跳过第一列（行号）
    const rows = Array.from(table.querySelectorAll('tbody tr'))
        .map(row => {
            const cells = Array.from(row.querySelectorAll('td')).slice(1);
            return cells.map(cell => cell.textContent.trim());
        });
        //.filter(row => row.some(cell => cell !== '')); // 过滤掉全空行

    return {
        columns: headers,
        data: rows
    };
}

// 导出Excel
function exportExcel() {
    const data = getCurrentTableData(); // 获取当前表格数据
    if (!data) {
        alert('没有可导出的数据');
        return;
    }
    console.log('Exporting data:', data); // 调试输出
    fetch('/save', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            data: data, // 使用合并后的数据
            filepath: filepath,
            filename: originalFilename
        })
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
    })
    .then(result => {
        if (result.success) {
            filepath = result.filepath;
            // 保存成功后，触发下载
            window.location.href = `/download?filepath=${encodeURIComponent(result.filepath)}`;
            //点击导出主动触发停止录音
            stopRecognition(true);
            return fetch('/delete_files', {
                method: 'DELETE'
            });
        } else {
            throw new Error(result.error || '导出失败');
        }
    })
    .catch(error => {
        console.error('Export error:', error);
        alert('导出失败: ' + error.message);
    });
}

// 保存更改
function saveChanges() {
    if (!filepath) return;

    const data = getCurrentTableData();
    if (!data) return;

    fetch('/save', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            data: data,
            filepath: filepath,
            filename: originalFilename
        })
    })
    .then(response => response.json())
    .then(result => {
        if (result.success) {
            filepath = result.filepath;
        } else {
            console.error('保存失败:', result.error);
        }
    })
    .catch(error => {
        console.error('Error:', error);
    });
}

// 初始化语音识别
function initSpeechRecognition() {
    if ('webkitSpeechRecognition' in window) {
        recognition = new webkitSpeechRecognition();
        recognition.continuous = true;
        recognition.interimResults = false;
        recognition.lang = 'zh-CN';

        recognition.onstart = function() {
            console.log('语音识别已启动');
            document.getElementById('voiceStatus').textContent = '语音识别已开启';
        };

        recognition.onerror = function(event) {           
            // 如果错误为 "network"，则在3秒后尝试重新启动识别
            if (event.error === 'network') {
                console.log('检测到网络错误，3秒后重试...');
                setTimeout(function() {
                    try {
                        recognition.start();
                        console.log('重试启动语音识别');
                    } catch (err) {
                        console.error('重试启动失败：', err);
                    }
                }, 3000);
            }else if (event.error === 'no-speech') {
                console.log('语音识别错误：没有检测到语音');
                restart = true;
                // 不关闭识别，继续等待用户输入
            } else {
                console.error('语音识别错误：', event.error);
            }
        };

        recognition.onend = function() {
            console.log('语音识别已结束');
            //继续启动
            if(!buttonStop){
                restart = true;
                console.log('自动重启语音识别');
            }
            startRecognition();
            document.getElementById('voiceStatus').textContent = '语音识别已停止';
        };
        recognition.onresult = function(event) {
            if (isProcessingCommand) return; // 如果正在处理命令，则直接返回
            isProcessingCommand = true; // 设置为正在处理命令

            const result = event.results[event.results.length - 1][0].transcript.trim();
            console.log('识别结果：', result);
            const cleanedResult = result.replace(/[，。！？；、]$/, ''); // 去掉最后的标点符号

            // 清除之前的定时器
            resetTimer(); // 重置计时器

            // 根据识别内容执行相应操作
            if (cleanedResult.includes('转')) {
                const cellReference = cleanedResult.split('转')[1].trim(); // 获取单元格引用
                moveToCell(cellReference); // 移动到指定单元格
            } else if (cleanedResult.includes('下一列') || cleanedResult.includes('向右')) {
                moveToNextColumn();
            } else if (cleanedResult.includes('上一列') || cleanedResult.includes('向左')) {
                moveToPreColumn();
            } else if (cleanedResult.includes('下一行') || cleanedResult.includes('向下')) {
                moveToNextRow();
            } else if (cleanedResult.includes('上一行') || cleanedResult.includes('向上')) {
                moveToPreviousRow();
            } else {
                // 直接更新当前单元格内容，避免重复
                if (currentCell) {
                    const currentContent = currentCell.textContent;
                    if (currentContent !== cleanedResult) { // 只有在内容不同的情况下才更新
                        currentCell.textContent = cleanedResult; // 更新单元格内容
                        //saveChanges();
                    }
                }
            }

            // 处理完成后重置标志
            setTimeout(() => {
                isProcessingCommand = false; // 处理完成后重置标志
                autoMoveEnabled = true; // 恢复自动移动
            }, 1000); // 设置一个适当的延迟，避免快速重复触发
        };
    } else {
        alert('您的浏览器不支持语音识别功能');
    }
}

// 重置计时器
function resetTimer() {
    clearTimeout(timer);
    timer = setTimeout(() => {
        if (autoMoveEnabled) { // 检查状态标志
            if (currentCell && currentCell.textContent.trim() === '') {
                // 如果目标单元格没有内容，则不再自动跳转
                return;
            }
            moveToNextColumn();
            if (currentCell) {
                currentCell.style.backgroundColor = 'green !important'; // 设置当前单元格背景色为绿色
            }
        }
    }, 5000); // 3秒后自动转到下一个单元格
}

// 移动到下一行
function moveToNextRow() {
    if (currentCell) {
        const currentRow = currentCell.parentElement;
        const nextRow = currentRow.nextElementSibling;

        // 获取当前单元格的索引
        const currentCellIndex = Array.from(currentRow.children).indexOf(currentCell);

        if (nextRow) {
            const nextCell = nextRow.children[currentCellIndex]; // 获取下一行的当前列单元格

            currentCell.style.backgroundColor = ''; // 清除当前单元格的背景色
            currentCell = nextCell; // 移动到下一行的当前列单元格
            currentCell.style.backgroundColor = 'green !important'; // 设置新单元格背景色为绿色
            currentCell.focus();
            resetTimer(); // 立即重置定时器
            autoMoveEnabled = false; // 收到"下一列"指令时禁用自动移动
        } else {
            // 如果没有下一行，自动添加新的一行
            const newRow = document.createElement('tr');
            const columnCount = currentRow.children.length;

            // 添加行号单元格
            const rowNumberCell = document.createElement('td');
            rowNumberCell.textContent = currentRow.parentElement.children.length + 1; // 设置行号
            rowNumberCell.className = 'row-number'; // 添加行号样式
            newRow.appendChild(rowNumberCell);

            for (let i = 1; i < columnCount; i++) { // 从1开始，跳过行号单元格
                const newCell = document.createElement('td');
                newCell.contentEditable = true; // 允许编辑
                newCell.className = 'data-cell'; // 添加数据单元格样式
                newCell.style.backgroundColor = 'lightgray'; // 设置背景色
                newRow.appendChild(newCell);
            }

            currentRow.parentElement.appendChild(newRow); // 添加新行到表格
            const newCell = newRow.children[currentCellIndex]; // 获取新行的当前列单元格
            currentCell.style.backgroundColor = ''; // 清除当前单元格的背景色
            currentCell = newCell;
            currentCell.style.backgroundColor = 'green'; // 设置新单元格背景色为绿色
            currentCell.focus(); // 聚焦到新单元格
            resetTimer(); // 立即重置定时器
            autoMoveEnabled = false; // 收到"下一列"指令时禁用自动移动
        }
    }
}

// 移动到下一列
function moveToNextColumn() {
    if (currentCell) {
        const currentRow = currentCell.parentElement;
        const currentCellIndex = Array.from(currentRow.children).indexOf(currentCell);
        const table = document.getElementById('excelTable');
        const lastRow = table.rows[table.rows.length - 1]; // 获取最后一行
        // 检查当前单元格是否为最后一行的最后一列
        if (currentRow === lastRow && currentCellIndex === lastRow.children.length - 1) {
            // 如果是最后一行的最后一列，新增一行
            const newRow = document.createElement('tr');
            const columnCount = currentRow.children.length;

            // 添加行号单元格
            const rowNumberCell = document.createElement('td');
            rowNumberCell.textContent = table.rows.length + 1; // 设置行号
            rowNumberCell.className = 'row-number'; // 添加行号样式
            newRow.appendChild(rowNumberCell);

            for (let i = 1; i < columnCount; i++) { // 从1开始，跳过行号单元格
                const newCell = document.createElement('td');
                newCell.contentEditable = true; // 允许编辑
                newCell.className = 'data-cell'; // 添加数据单元格样式
                newCell.style.backgroundColor = 'lightgray'; // 设置背景色
                newRow.appendChild(newCell);
            }

            currentRow.parentElement.appendChild(newRow); // 添加新行到表格
            currentCell = newRow.children[1]; // 移动到新行的第一个数据单元格
            currentCell.style.backgroundColor = 'green'; // 设置新单元格背景色为绿色
            currentCell.focus(); // 聚焦到新单元格
            resetTimer(); // 立即重置定时器
            autoMoveEnabled = false; // 收到"下一列"指令时禁用自动移动
        } else if (currentCellIndex === currentRow.children.length - 1) {
            // 如果当前单元格是其他行的最后一列，移动到下一行的第一列（非导航列）
            const nextRow = currentRow.nextElementSibling;

            if (nextRow) {
                const firstCell = nextRow.children[1]; // 获取下一行的第一个数据单元格（跳过导航列）
                currentCell.style.backgroundColor = ''; // 清除当前单元格的背景色
                currentCell = firstCell; // 移动到下一行的第一个数据单元格
                currentCell.style.backgroundColor = 'green'; // 设置新单元格背景色为绿色
                currentCell.focus(); // 聚焦到新单元格
            }
        } else {
             // 如果当前单元格不是本行最后一列，正常向右移动一列
            const nextCell = currentRow.children[currentCellIndex + 1]; // 获取右侧单元格

            if (nextCell) {
                currentCell.style.backgroundColor = ''; // 清除当前单元格的背景色
                currentCell = nextCell; // 移动到右侧单元格
                currentCell.style.backgroundColor = 'green'; // 设置新单元格背景色为绿色
                currentCell.focus(); // 聚焦到新单元格
                resetTimer(); // 立即重置定时器
                autoMoveEnabled = false; // 收到"下一列"指令时禁用自动移动
            }
        }
    }
}

// 移动到上一列
function moveToPreColumn() {
    if (currentCell) {
        const currentRowIndex = currentCell.parentElement.rowIndex;
        const currentCellIndex = Array.from(currentCell.parentElement.children).indexOf(currentCell);

        // 如果当前是第一行第一列，或者当前是第一行，直接返回，不处理
        if (currentRowIndex === 0 && currentCellIndex === 0) {
            return;
        }

        currentCell.style.backgroundColor = ''; // 清除当前单元格的背景色
        const previousCell = currentCell.previousElementSibling;
        if (previousCell) {
            currentCell = previousCell;
            currentCell.style.backgroundColor = 'green !important'; // 设置新单元格背景色为绿色
            previousCell.focus();
            resetTimer(); // 更新内容后重置定时器
            autoMoveEnabled = false; // 收到"下一列"指令时禁用自动移动
        } else {
            moveToPreviousRow(); // 如果到达第一列，移动到上一行的第一列
        }
    }
}

// 移动到上一行
function moveToPreviousRow() {
    if (currentCell) {
        const currentRowIndex = currentCell.parentElement.rowIndex;

        // 如果当前是第一行，直接返回，不处理
        if (currentRowIndex === 0) {
            return;
        }

        const currentRow = currentCell.parentElement;
        const previousRow = currentRow.previousElementSibling;
        if (previousRow) {
            const currentCellIndex = Array.from(currentRow.children).indexOf(currentCell); // 获取当前单元格的索引
            const targetCell = previousRow.children[currentCellIndex]; // 获取上一行的当前列单元格

            currentCell.style.backgroundColor = ''; // 清除当前单元格的背景色
            currentCell = targetCell; // 移动到目标单元格
            currentCell.style.backgroundColor = 'green !important'; // 设置新单元格背景色为绿色
            currentCell.focus(); // 聚焦到新单元格
            resetTimer(); // 更新内容后重置定时器
            autoMoveEnabled = false; // 收到"下一列"指令时禁用自动移动
        }
    }
}

// 移动到指定单元格
function moveToCell(cellReference) {
    const regex = /^([A-Z]+)(\d+)$/; // 正则表达式匹配单元格格式（如B2）
    const match = cellReference.match(regex);
    if (match) {
        const columnLetter = match[1]; // 列字母
        const rowNumber = parseInt(match[2], 10); // 行数字

        // 计算列索引（A=0, B=1, C=2, ...）
        const columnIndex = columnLetter.charCodeAt(0) - 'A'.charCodeAt(0);
        const rowIndex = rowNumber - 1; // 行索引从0开始

        const table = document.querySelector('table'); // 假设你的表格是一个 <table> 元素
        const targetRow = table.rows[rowIndex];
        if (targetRow && targetRow.cells[columnIndex]) {
            currentCell.style.backgroundColor = ''; // 清除当前单元格的背景色
            currentCell = targetRow.cells[columnIndex]; // 移动到目标单元格
            currentCell.style.backgroundColor = 'green !important'; // 设置新单元格背景色为绿色
            currentCell.focus();
            resetTimer(); // 更新内容后重置定时器
            autoMoveEnabled = false; // 收到"下一列"指令时禁用自动移动
        }
    }
}

function startRecognition(start) {
    if(start){
        buttonStop = false;
    }
    if(start || restart){
        restart = start;
        recognition.start();
        const recordingIndicator = document.getElementById('recordingIndicator');
        recordingIndicator.style.display = 'block'; // 显示录音指示器
        recordingIndicator.style.backgroundColor = 'green !important'; // 设置为绿色
        recordingIndicator.classList.add('blinking'); // 添加闪烁效果
        document.getElementById('voiceStatus').textContent = '语音识别已开启';
    }
}

function stopRecognition(stop) {
    if(stop){
        restart = false;
        buttonStop = stop;
        console.log("按钮停止语音识别")
    }
    recognition.stop();
    const recordingIndicator = document.getElementById('recordingIndicator');
    recordingIndicator.style.display = 'none'; // 隐藏录音指示器
    recordingIndicator.classList.remove('blinking'); // 移除闪烁效果
    recordingIndicator.style.backgroundColor = 'red'; // 设置为红色
    document.getElementById('voiceStatus').textContent = '语音识别未开启';
}

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    initSpeechRecognition();
    
    // 设置自动保存
    setInterval(saveChanges, 10 * 60 * 1000); // 每5分钟自动保存
});