
class WPSOffice {
    constructor() {
        this.currentModule = 'writer';
        this.documents = {
            writer: { title: '未命名文档', content: '<p>开始输入文字...</p>' },
            spreadsheet: { data: [] },
            presentation: { slides: [] }
        };
        this.history = [];
        this.historyIndex = -1;
        this.recentFiles = [];
        this.currentSlideIndex = 0;
        this.zoomLevel = 100;
        this.init();
    }

    init() {
        this.loadFromStorage();
        this.setupTabs();
        this.setupToolbar();
        this.setupWriter();
        this.setupSpreadsheet();
        this.setupPresentation();
        this.setupFileCards();
        this.setupSaveButton();
        this.setupAllMenus();
        this.setupFileInput();
        this.setupImageInput();
        this.setupModal();
        this.updateRecentFiles();
    }

    loadFromStorage() {
        const savedData = localStorage.getItem('wpsOfficeData');
        if (savedData) {
            const data = JSON.parse(savedData);
            if (data.documents) this.documents = data.documents;
            if (data.recentFiles) this.recentFiles = data.recentFiles;
        }
    }

    saveToStorage() {
        localStorage.setItem('wpsOfficeData', JSON.stringify({
            documents: this.documents,
            recentFiles: this.recentFiles
        }));
    }

    setupTabs() {
        const tabs = document.querySelectorAll('.tab');
        tabs.forEach(tab => {
            tab.addEventListener('click', (e) => {
                if (e.target.classList.contains('tab-close')) return;
                const module = tab.dataset.module;
                this.switchModule(module);
            });
        });
    }

    switchModule(module) {
        this.currentModule = module;
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        document.querySelector('.tab[data-module="' + module + '"]').classList.add('active');
        document.querySelectorAll('.module-content').forEach(m => m.classList.add('hidden'));
        document.getElementById(module + 'Module').classList.remove('hidden');
    }

    setupToolbar() {
        document.getElementById('boldBtn').addEventListener('click', () => {
            document.execCommand('bold', false, null);
            this.saveToStorage();
        });

        document.getElementById('italicBtn').addEventListener('click', () => {
            document.execCommand('italic', false, null);
            this.saveToStorage();
        });

        document.getElementById('underlineBtn').addEventListener('click', () => {
            document.execCommand('underline', false, null);
            this.saveToStorage();
        });

        document.getElementById('strikeBtn').addEventListener('click', () => {
            document.execCommand('strikeThrough', false, null);
            this.saveToStorage();
        });

        document.getElementById('fontFamily').addEventListener('change', (e) => {
            document.execCommand('fontName', false, e.target.value);
            this.saveToStorage();
        });

        document.getElementById('fontSize').addEventListener('change', (e) => {
            document.execCommand('fontSize', false, e.target.value);
            this.saveToStorage();
        });

        document.getElementById('alignLeft').addEventListener('click', () => {
            document.execCommand('justifyLeft', false, null);
            this.saveToStorage();
        });

        document.getElementById('alignCenter').addEventListener('click', () => {
            document.execCommand('justifyCenter', false, null);
            this.saveToStorage();
        });

        document.getElementById('alignRight').addEventListener('click', () => {
            document.execCommand('justifyRight', false, null);
            this.saveToStorage();
        });

        document.getElementById('alignJustify').addEventListener('click', () => {
            document.execCommand('justifyFull', false, null);
            this.saveToStorage();
        });

        document.getElementById('bulletList').addEventListener('click', () => {
            document.execCommand('insertUnorderedList', false, null);
            this.saveToStorage();
        });

        document.getElementById('numberList').addEventListener('click', () => {
            document.execCommand('insertOrderedList', false, null);
            this.saveToStorage();
        });

        document.getElementById('fontColor').addEventListener('input', (e) => {
            document.execCommand('foreColor', false, e.target.value);
            this.saveToStorage();
        });

        document.getElementById('bgColor').addEventListener('input', (e) => {
            document.execCommand('hiliteColor', false, e.target.value);
            this.saveToStorage();
        });

        document.getElementById('undoBtn').addEventListener('click', () => {
            document.execCommand('undo', false, null);
        });

        document.getElementById('redoBtn').addEventListener('click', () => {
            document.execCommand('redo', false, null);
        });
    }

    setupWriter() {
        const docBody = document.getElementById('docBody');
        const docTitle = document.getElementById('docTitle');

        if (this.documents.writer.content) {
            docBody.innerHTML = this.documents.writer.content;
        }
        if (this.documents.writer.title) {
            docTitle.value = this.documents.writer.title;
        }

        docBody.addEventListener('input', () => {
            this.documents.writer.content = docBody.innerHTML;
            this.saveToStorage();
        });

        docTitle.addEventListener('input', () => {
            this.documents.writer.title = docTitle.value;
            this.saveToStorage();
        });

        docBody.addEventListener('click', (e) => {
            const link = e.target.closest('a');
            if (link && link.href) {
                e.preventDefault();
                window.open(link.href, '_blank');
                return;
            }

            const dateSpan = e.target.closest('.wps-date');
            if (dateSpan) {
                e.preventDefault();
                const newDate = prompt('请输入新日期（格式：YYYY-MM-DD）或留空使用当前日期：');
                if (newDate !== null) {
                    let dateObj;
                    if (newDate.trim()) {
                        dateObj = new Date(newDate);
                        if (isNaN(dateObj.getTime())) {
                            alert('日期格式无效，请使用 YYYY-MM-DD 格式');
                            return;
                        }
                    } else {
                        dateObj = new Date();
                    }
                    dateSpan.textContent = dateObj.toLocaleDateString('zh-CN');
                    dateSpan.dataset.date = dateObj.toISOString();
                    this.documents.writer.content = docBody.innerHTML;
                    this.saveToStorage();
                }
                return;
            }

            const timeSpan = e.target.closest('.wps-time');
            if (timeSpan) {
                e.preventDefault();
                const newTime = prompt('请输入新时间（格式：HH:MM）或留空使用当前时间：');
                if (newTime !== null) {
                    let dateObj;
                    if (newTime.trim()) {
                        const parts = newTime.split(':');
                        dateObj = new Date();
                        if (parts.length >= 1) dateObj.setHours(parseInt(parts[0]) || 0);
                        if (parts.length >= 2) dateObj.setMinutes(parseInt(parts[1]) || 0);
                    } else {
                        dateObj = new Date();
                    }
                    timeSpan.textContent = dateObj.toLocaleTimeString('zh-CN');
                    timeSpan.dataset.time = dateObj.toISOString();
                    this.documents.writer.content = docBody.innerHTML;
                    this.saveToStorage();
                }
            }
        });
    }

    setupSpreadsheet() {
        const tbody = document.getElementById('spreadsheetBody');
        for (let i = 1; i <= 20; i++) {
            const row = document.createElement('tr');
            const numCell = document.createElement('td');
            numCell.textContent = i;
            row.appendChild(numCell);
            for (let j = 0; j < 5; j++) {
                const cell = document.createElement('td');
                const input = document.createElement('input');
                input.addEventListener('input', () => {
                    this.saveSpreadsheetData();
                    this.saveToStorage();
                });
                cell.appendChild(input);
                row.appendChild(cell);
            }
            tbody.appendChild(row);
        }

        document.getElementById('addRow').addEventListener('click', () => {
            const rowCount = tbody.children.length + 1;
            const row = document.createElement('tr');
            const numCell = document.createElement('td');
            numCell.textContent = rowCount;
            row.appendChild(numCell);
            const colCount = document.querySelector('#spreadsheetTable thead tr').children.length - 1;
            for (let j = 0; j < colCount; j++) {
                const cell = document.createElement('td');
                const input = document.createElement('input');
                input.addEventListener('input', () => {
                    this.saveSpreadsheetData();
                    this.saveToStorage();
                });
                cell.appendChild(input);
                row.appendChild(cell);
            }
            tbody.appendChild(row);
            this.saveToStorage();
        });

        document.getElementById('addCol').addEventListener('click', () => {
            const thead = document.querySelector('#spreadsheetTable thead tr');
            const colCount = thead.children.length - 1;
            const colLetter = String.fromCharCode(65 + colCount);
            const th = document.createElement('th');
            th.textContent = colLetter;
            thead.appendChild(th);

            tbody.querySelectorAll('tr').forEach(row => {
                const cell = document.createElement('td');
                const input = document.createElement('input');
                input.addEventListener('input', () => {
                    this.saveSpreadsheetData();
                    this.saveToStorage();
                });
                cell.appendChild(input);
                row.appendChild(cell);
            });
            this.saveToStorage();
        });

        document.getElementById('deleteRow').addEventListener('click', () => {
            if (tbody.children.length > 1) {
                tbody.removeChild(tbody.lastChild);
                this.saveToStorage();
            }
        });

        document.getElementById('deleteCol').addEventListener('click', () => {
            const thead = document.querySelector('#spreadsheetTable thead tr');
            if (thead.children.length > 2) {
                thead.removeChild(thead.lastChild);
                tbody.querySelectorAll('tr').forEach(row => {
                    row.removeChild(row.lastChild);
                });
                this.saveToStorage();
            }
        });
    }

    saveSpreadsheetData() {
        const tbody = document.getElementById('spreadsheetBody');
        const data = [];
        tbody.querySelectorAll('tr').forEach(row => {
            const rowData = [];
            const inputs = row.querySelectorAll('input');
            inputs.forEach(input => {
                rowData.push(input.value);
            });
            data.push(rowData);
        });
        this.documents.spreadsheet.data = data;
    }

    setupPresentation() {
        if (this.documents.presentation.slides.length === 0) {
            this.addSlide('点击添加标题', '点击添加内容');
        } else {
            this.renderSlidesFromStorage();
        }
        
        document.getElementById('addSlide').addEventListener('click', () => {
            this.addSlide('新幻灯片', '内容');
        });
    }

    renderSlidesFromStorage() {
        const slidesList = document.getElementById('slidesList');
        slidesList.innerHTML = '';
        this.documents.presentation.slides.forEach((slide, index) => {
            const thumbnail = document.createElement('div');
            thumbnail.className = 'slide-thumbnail';
            thumbnail.textContent = '幻灯片 ' + (index + 1);
            thumbnail.dataset.index = index;
            if (index === this.currentSlideIndex) {
                thumbnail.classList.add('active');
            }
            thumbnail.addEventListener('click', () => {
                this.selectSlide(index);
            });
            slidesList.appendChild(thumbnail);
        });
        this.renderSlide(this.currentSlideIndex);
    }

    addSlide(title, content) {
        const slideIndex = this.documents.presentation.slides.length;
        this.documents.presentation.slides.push({ title: title, content: content });
        
        const slidesList = document.getElementById('slidesList');
        const thumbnail = document.createElement('div');
        thumbnail.className = 'slide-thumbnail';
        thumbnail.textContent = '幻灯片 ' + (slideIndex + 1);
        thumbnail.dataset.index = slideIndex;
        if (slideIndex === this.currentSlideIndex) {
            thumbnail.classList.add('active');
        }
        thumbnail.addEventListener('click', () => {
            this.selectSlide(slideIndex);
        });
        slidesList.appendChild(thumbnail);
        
        if (slideIndex === 0) {
            this.renderSlide(0);
        }
        
        this.saveToStorage();
    }

    selectSlide(index) {
        this.currentSlideIndex = index;
        document.querySelectorAll('.slide-thumbnail').forEach((thumb, i) => {
            thumb.classList.toggle('active', i === index);
        });
        this.renderSlide(index);
    }

    renderSlide(index) {
        if (!this.documents.presentation.slides[index]) return;
        const slide = this.documents.presentation.slides[index];
        const canvas = document.getElementById('slideCanvas');
        canvas.innerHTML = '<div class="slide-title" contenteditable="true">' + slide.title + '</div><div class="slide-content" contenteditable="true">' + slide.content + '</div>';
        
        canvas.querySelector('.slide-title').addEventListener('input', (e) => {
            this.documents.presentation.slides[index].title = e.target.textContent;
            this.saveToStorage();
        });
        
        canvas.querySelector('.slide-content').addEventListener('input', (e) => {
            this.documents.presentation.slides[index].content = e.target.textContent;
            this.saveToStorage();
        });
    }

    setupFileCards() {
        document.querySelectorAll('.file-type-card').forEach(card => {
            card.addEventListener('click', () => {
                const type = card.dataset.type;
                this.switchModule(type);
            });
        });
    }

    setupSaveButton() {
        document.getElementById('saveBtn').addEventListener('click', () => {
            this.saveDocument();
        });
    }

    saveDocument() {
        const title = this.documents.writer.title || '未命名文档';
        const timestamp = new Date().toLocaleString();
        
        this.recentFiles.unshift({
            id: Date.now(),
            title: title,
            type: this.currentModule,
            date: timestamp
        });
        
        if (this.recentFiles.length > 10) {
            this.recentFiles.pop();
        }
        
        this.saveToStorage();
        this.updateRecentFiles();
        alert('文档已保存！');
    }

    updateRecentFiles() {
        const container = document.getElementById('recentFiles');
        if (this.recentFiles.length === 0) {
            container.innerHTML = '<div class="empty-state">暂无最近文档</div>';
            return;
        }
        
        const icons = { writer: '📄', spreadsheet: '📊', presentation: '📽️' };
        container.innerHTML = this.recentFiles.map(file => {
            return '<div class="recent-file-item" data-id="' + file.id + '"><span style="font-size: 24px;">' + icons[file.type] + '</span><div><div style="font-size: 13px; font-weight: 500;">' + file.title + '</div><div style="font-size: 11px; color: #999;">' + file.date + '</div></div></div>';
        }).join('');
    }

    setupAllMenus() {
        const menuButtons = document.querySelectorAll('.menu-btn');
        const menus = {
            'file': 'fileMenu',
            'edit': 'editMenu',
            'view': 'viewMenu',
            'insert': 'insertMenu',
            'format': 'formatMenu',
            'tools': 'toolsMenu'
        };

        menuButtons.forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const menuName = btn.dataset.menu;
                const menuId = menus[menuName];
                
                document.querySelectorAll('.dropdown-menu').forEach(m => m.classList.remove('show'));
                
                if (menuId) {
                    const menu = document.getElementById(menuId);
                    const rect = btn.getBoundingClientRect();
                    menu.style.left = rect.left + 'px';
                    menu.style.top = rect.bottom + 'px';
                    menu.classList.toggle('show');
                }
            });
        });

        document.addEventListener('click', () => {
            document.querySelectorAll('.dropdown-menu').forEach(m => m.classList.remove('show'));
        });

        document.querySelectorAll('.dropdown-item[data-action]').forEach(item => {
            item.addEventListener('click', (e) => {
                e.stopPropagation();
                const action = item.dataset.action;
                this.handleMenuAction(action);
                document.querySelectorAll('.dropdown-menu').forEach(m => m.classList.remove('show'));
            });
        });
    }

    handleMenuAction(action) {
        switch(action) {
            case 'new':
                this.newDocument();
                break;
            case 'open':
                this.openDocument();
                break;
            case 'save':
                this.saveDocument();
                break;
            case 'export-html':
                this.exportDocument('html');
                break;
            case 'export-pdf':
                this.exportDocument('pdf');
                break;
            case 'export-doc':
                this.exportDocument('doc');
                break;
            case 'export-txt':
                this.exportDocument('txt');
                break;
            case 'print':
                this.printDocument();
                break;
            case 'undo':
                document.execCommand('undo', false, null);
                break;
            case 'redo':
                document.execCommand('redo', false, null);
                break;
            case 'cut':
                document.execCommand('cut', false, null);
                break;
            case 'copy':
                document.execCommand('copy', false, null);
                break;
            case 'paste':
                document.execCommand('paste', false, null);
                break;
            case 'selectAll':
                document.execCommand('selectAll', false, null);
                break;
            case 'clear':
                document.execCommand('delete', false, null);
                break;
            case 'zoomIn':
                this.zoomIn();
                break;
            case 'zoomOut':
                this.zoomOut();
                break;
            case 'zoomReset':
                this.zoomReset();
                break;
            case 'fullscreen':
                this.toggleFullscreen();
                break;
            case 'insertImage':
                this.insertImage();
                break;
            case 'insertLink':
                this.insertLink();
                break;
            case 'insertTable':
                this.insertTable();
                break;
            case 'insertHR':
                this.insertHTMLAtCursor('<hr style="margin: 10px 0; border: none; border-top: 1px solid #ccc;" />');
                break;
            case 'insertDate':
                this.insertDate();
                break;
            case 'insertTime':
                this.insertTime();
                break;
            case 'bold':
                document.execCommand('bold', false, null);
                this.saveToStorage();
                break;
            case 'italic':
                document.execCommand('italic', false, null);
                this.saveToStorage();
                break;
            case 'underline':
                document.execCommand('underline', false, null);
                this.saveToStorage();
                break;
            case 'strike':
                document.execCommand('strikeThrough', false, null);
                this.saveToStorage();
                break;
            case 'superscript':
                document.execCommand('superscript', false, null);
                this.saveToStorage();
                break;
            case 'subscript':
                document.execCommand('subscript', false, null);
                this.saveToStorage();
                break;
            case 'increaseIndent':
                document.execCommand('indent', false, null);
                this.saveToStorage();
                break;
            case 'decreaseIndent':
                document.execCommand('outdent', false, null);
                this.saveToStorage();
                break;
            case 'removeFormat':
                document.execCommand('removeFormat', false, null);
                this.saveToStorage();
                break;
            case 'wordCount':
                this.showWordCount();
                break;
            case 'findReplace':
                this.showFindReplace();
                break;
        }
    }

    newDocument() {
        if (confirm('确定要新建文档吗？当前未保存的内容将丢失。')) {
            if (this.currentModule === 'writer') {
                this.documents.writer = { title: '未命名文档', content: '<p>开始输入文字...</p>' };
                document.getElementById('docTitle').value = '';
                document.getElementById('docBody').innerHTML = '<p>开始输入文字...</p>';
            } else if (this.currentModule === 'spreadsheet') {
                this.documents.spreadsheet = { data: [] };
                const tbody = document.getElementById('spreadsheetBody');
                tbody.innerHTML = '';
                this.setupSpreadsheet();
            } else if (this.currentModule === 'presentation') {
                this.documents.presentation = { slides: [] };
                this.currentSlideIndex = 0;
                document.getElementById('slidesList').innerHTML = '';
                this.addSlide('点击添加标题', '点击添加内容');
            }
            this.saveToStorage();
        }
    }

    setupFileInput() {
        const fileInput = document.getElementById('fileInput');
        fileInput.addEventListener('change', (e) => {
            this.importFile(e.target.files[0]);
        });
    }

    openDocument() {
        document.getElementById('fileInput').click();
    }

    importFile(file) {
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = (e) => {
            const content = e.target.result;
            const filename = file.name.toLowerCase();
            
            if (filename.endsWith('.csv')) {
                this.importCSV(content);
            } else if (filename.endsWith('.html') || filename.endsWith('.htm')) {
                this.importHTML(content, filename);
            } else if (filename.endsWith('.txt')) {
                this.importTXT(content, filename);
            } else {
                alert('不支持的文件格式！');
            }
        };
        reader.readAsText(file);
    }

    importCSV(content) {
        this.switchModule('spreadsheet');
        const lines = content.split('\n').filter(line => line.trim());
        const tbody = document.getElementById('spreadsheetBody');
        tbody.innerHTML = '';
        
        lines.forEach((line, index) => {
            const row = document.createElement('tr');
            const numCell = document.createElement('td');
            numCell.textContent = index + 1;
            row.appendChild(numCell);
            
            const values = line.split(',');
            values.forEach(value => {
                const cell = document.createElement('td');
                const input = document.createElement('input');
                input.value = value.trim();
                input.addEventListener('input', () => {
                    this.saveSpreadsheetData();
                    this.saveToStorage();
                });
                cell.appendChild(input);
                row.appendChild(cell);
            });
            tbody.appendChild(row);
        });
        
        this.saveSpreadsheetData();
        this.saveToStorage();
        alert('CSV文件已导入！');
    }

    importHTML(content, filename) {
        this.switchModule('writer');
        
        const parser = new DOMParser();
        const doc = parser.parseFromString(content, 'text/html');
        const bodyContent = doc.body.innerHTML;
        
        const titleMatch = content.match(/<title>(.*?)<\/title>/i);
        const title = titleMatch ? titleMatch[1] : filename.replace('.html', '').replace('.htm', '');
        
        document.getElementById('docTitle').value = title;
        document.getElementById('docBody').innerHTML = bodyContent;
        
        this.documents.writer.title = title;
        this.documents.writer.content = bodyContent;
        this.saveToStorage();
        alert('HTML文件已导入！');
    }

    importTXT(content, filename) {
        this.switchModule('writer');
        
        const title = filename.replace('.txt', '');
        const paragraphs = content.split('\n').map(p => '<p>' + p + '</p>').join('');
        
        document.getElementById('docTitle').value = title;
        document.getElementById('docBody').innerHTML = paragraphs;
        
        this.documents.writer.title = title;
        this.documents.writer.content = paragraphs;
        this.saveToStorage();
        alert('文本文件已导入！');
    }

    exportDocument(format) {
        let content, filename, mimeType;
        const title = this.documents.writer.title || '未命名文档';
        
        if (this.currentModule === 'writer') {
            if (format === 'html') {
                const htmlContent = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + title + '</title></head><body><h1>' + title + '</h1>' + this.documents.writer.content + '</body></html>';
                content = htmlContent;
                filename = title + '.html';
                mimeType = 'text/html';
            } else if (format === 'txt') {
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = this.documents.writer.content;
                content = tempDiv.textContent || tempDiv.innerText || '';
                filename = title + '.txt';
                mimeType = 'text/plain';
            } else if (format === 'doc') {
                const htmlContent = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + title + '</title></head><body><h1>' + title + '</h1>' + this.documents.writer.content + '</body></html>';
                content = htmlContent;
                filename = title + '.doc';
                mimeType = 'application/msword';
            } else if (format === 'pdf') {
                this.printDocument();
                return;
            }
        } else if (this.currentModule === 'spreadsheet') {
            this.saveSpreadsheetData();
            const data = this.documents.spreadsheet.data;
            let csvContent = '';
            data.forEach(row => {
                csvContent += row.join(',') + '\n';
            });
            content = csvContent;
            filename = '表格.' + (format === 'txt' ? 'txt' : 'csv');
            mimeType = format === 'txt' ? 'text/plain' : 'text/csv';
        } else if (this.currentModule === 'presentation') {
            const slides = this.documents.presentation.slides;
            let htmlContent = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>演示文稿</title><style>body{font-family:Arial,sans-serif;}.slide{page-break-after:always;padding:50px;}h1{font-size:44px;}.content{font-size:24px;}</style></head><body>';
            slides.forEach(slide => {
                htmlContent += '<div class="slide"><h1>' + slide.title + '</h1><div class="content">' + slide.content + '</div></div>';
            });
            htmlContent += '</body></html>';
            content = htmlContent;
            filename = '演示文稿.' + (format === 'pdf' ? 'html' : format);
            mimeType = 'text/html';
            
            if (format === 'pdf') {
                this.printDocument();
                return;
            }
        }

        this.downloadFile(content, filename, mimeType);
    }

    downloadFile(content, filename, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    printDocument() {
        window.print();
    }

    zoomIn() {
        this.zoomLevel = Math.min(this.zoomLevel + 10, 200);
        this.applyZoom();
    }

    zoomOut() {
        this.zoomLevel = Math.max(this.zoomLevel - 10, 50);
        this.applyZoom();
    }

    zoomReset() {
        this.zoomLevel = 100;
        this.applyZoom();
    }

    applyZoom() {
        const mainContent = document.querySelector('.main-content');
        mainContent.style.transform = 'scale(' + (this.zoomLevel / 100) + ')';
        mainContent.style.transformOrigin = 'top left';
    }

    toggleFullscreen() {
        if (!document.fullscreenElement) {
            document.documentElement.requestFullscreen();
        } else {
            document.exitFullscreen();
        }
    }

    setupImageInput() {
        const imageInput = document.getElementById('imageInput');
        imageInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = (event) => {
                    this.insertHTMLAtCursor('<img src="' + event.target.result + '" style="max-width: 100%;" />');
                };
                reader.readAsDataURL(file);
            }
            imageInput.value = '';
        });
    }

    insertImage() {
        document.getElementById('imageInput').click();
    }

    insertLink() {
        const url = prompt('请输入链接地址：', 'https://');
        if (url) {
            this.insertHTMLAtCursor('<a href="' + url + '" target="_blank" style="color: #0066cc; text-decoration: underline;">' + url + '</a>');
        }
    }

    insertTable() {
        const rows = prompt('请输入行数：', '3');
        const cols = prompt('请输入列数：', '3');
        if (rows && cols) {
            let table = '<table style="border-collapse: collapse; width: 100%; margin: 10px 0;">';
            for (let i = 0; i < parseInt(rows); i++) {
                table += '<tr>';
                for (let j = 0; j < parseInt(cols); j++) {
                    table += '<td style="border: 1px solid #ccc; padding: 8px; min-width: 50px;">&nbsp;</td>';
                }
                table += '</tr>';
            }
            table += '</table>';
            this.insertHTMLAtCursor(table);
        }
    }

    insertDate() {
        const now = new Date();
        const dateStr = now.toLocaleDateString('zh-CN');
        this.insertHTMLAtCursor('<span class="wps-date" contenteditable="false" style="color: #0066cc; cursor: pointer; text-decoration: underline;" data-date="' + now.toISOString() + '">' + dateStr + '</span>');
    }

    insertTime() {
        const now = new Date();
        const timeStr = now.toLocaleTimeString('zh-CN');
        this.insertHTMLAtCursor('<span class="wps-time" contenteditable="false" style="color: #0066cc; cursor: pointer; text-decoration: underline;" data-time="' + now.toISOString() + '">' + timeStr + '</span>');
    }

    insertHTMLAtCursor(html) {
        const docBody = document.getElementById('docBody');
        docBody.focus();
        
        if (document.getSelection().rangeCount > 0) {
            const range = document.getSelection().getRangeAt(0);
            range.deleteContents();
            const temp = document.createElement('div');
            temp.innerHTML = html;
            const frag = document.createDocumentFragment();
            let node;
            while ((node = temp.firstChild)) {
                frag.appendChild(node);
            }
            range.insertNode(frag);
        } else {
            docBody.innerHTML += html;
        }
        
        this.documents.writer.content = docBody.innerHTML;
        this.saveToStorage();
    }

    setupModal() {
        const modal = document.getElementById('modal');
        const closeBtn = document.getElementById('modalClose');
        
        closeBtn.addEventListener('click', () => {
            modal.classList.remove('show');
        });
        
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('show');
            }
        });
    }

    showModal(content) {
        const modal = document.getElementById('modal');
        const modalBody = document.getElementById('modalBody');
        modalBody.innerHTML = content;
        modal.classList.add('show');
    }

    showWordCount() {
        const docBody = document.getElementById('docBody');
        const text = docBody.textContent || docBody.innerText || '';
        const chars = text.length;
        const charsNoSpace = text.replace(/\s/g, '').length;
        const words = text.split(/\s+/).filter(w => w).length;
        const paragraphs = docBody.querySelectorAll('p').length;

        const content = '<h3>字数统计</h3><div class="word-count-info"><p>字符数（含空格）：<span>' + chars + '</span></p><p>字符数（不含空格）：<span>' + charsNoSpace + '</span></p><p>词数：<span>' + words + '</span></p><p>段落数：<span>' + paragraphs + '</span></p></div>';
        this.showModal(content);
    }

    showFindReplace() {
        const content = '<h3>查找替换</h3><div class="form-group"><label>查找内容：</label><input type="text" id="findInput" placeholder="输入要查找的内容"></div><div class="form-group"><label>替换为：</label><input type="text" id="replaceInput" placeholder="输入替换后的内容"></div><div class="modal-buttons"><button class="modal-btn modal-btn-secondary" onclick="wpsOffice.closeModal()">取消</button><button class="modal-btn modal-btn-primary" onclick="wpsOffice.doReplace()">替换</button></div>';
        this.showModal(content);
    }

    closeModal() {
        document.getElementById('modal').classList.remove('show');
    }

    doReplace() {
        const findText = document.getElementById('findInput').value;
        const replaceText = document.getElementById('replaceInput').value;
        
        if (findText) {
            const docBody = document.getElementById('docBody');
            docBody.innerHTML = docBody.innerHTML.replace(new RegExp(findText, 'g'), replaceText);
            this.documents.writer.content = docBody.innerHTML;
            this.saveToStorage();
            this.closeModal();
            alert('替换完成！');
        }
    }
}

let wpsOffice;
document.addEventListener('DOMContentLoaded', () => {
    wpsOffice = new WPSOffice();
});
