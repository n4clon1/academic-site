// ==================== ПРОВЕРКА АВТОРИЗАЦИИ ====================
/*(async function checkAuthentication() {
    try {
        const response = await fetch('/api/check-auth');
        if (!response.ok) {
            window.location.href = '/login';
        }
    } catch (error) {
        console.error('Ошибка проверки аутентификации:', error);
        window.location.href = '/login';
    }
})();
*/
window.currentUser = { authenticated: true, username: 'test_user' };
// ==================== ПЕРЕКЛЮЧЕНИЕ ВКЛАДОК ====================
document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', function() {
        const tabName = this.getAttribute('data-tab');
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        
        document.querySelectorAll('.app-content').forEach(content => {
            content.classList.remove('active');
        });
        document.getElementById(tabName + '-app').classList.add('active');

        // При переключении на вкладку "Мои файлы" загружаем список
        if (tabName === 'myfiles') {
            loadMyFiles();
        }
    });
});

// ==================== ПРИЛОЖЕНИЕ 1: ПРЕПОДАВАТЕЛЬСКАЯ НАГРУЗКА ====================
(function() {
    // Обработчики событий
    document.getElementById('loadBtnWorkload').addEventListener('click', handleFileUploadWorkload);
    document.getElementById('addTeacherBtn').addEventListener('click', addTeacher);
    document.getElementById('addSelectedTeacherBtn').addEventListener('click', addSelectedTeacher);
    document.getElementById('addAllTeachersBtn').addEventListener('click', addAllTeachersFromList);
    document.getElementById('cancelModal').addEventListener('click', hideModal);
    document.getElementById('exportBtnWorkload').addEventListener('click', exportToExcel);
    document.getElementById('previewBtnWorkload').addEventListener('click', showWorkloadPreview);
    document.getElementById('restoreHistoryBtnWorkload').addEventListener('click', applyHistoricalAssignments);

    // Встроенный список преподавателей
    const BUILT_IN_TEACHERS = [
        "Яковлева", "Кареев", "Белова", "Рычихина", "Звонарева", "Мутаев", "Алферова",
        "Соловьева", "Хасбулатова", "Панкратова", "Смирнова", "Ульянкина", "Берендеева",
        "Коробова", "Ситникова", "Мартынов", "Задорожникова", "Когаловская", "Птицына"
    ];

    // Хранилище данных
    let currentData = {
        faculties: {},
        teachers: [],
        assignments: [],
        originalWorkbook: null,
        originalSheetData: null,
        headerRows: [],
        cellStyles: {},
        columnWidths: {},
        rowHeights: {},
        mergedCells: [],
        cellFormats: {},
        subjectNameSettings: {},
        subgroupSettings: {}
    };

    let currentAssignment = {
        directionId: null,
        isSubgroup: false,
        subjectName: '',
        directionCode: '',
        facultyName: '',
        directionData: null
    };

    // Инициализация выпадающего списка преподавателей
    function initTeacherSelect() {
        const select = document.getElementById('teacherSelect');
        if (!select) return;
        
        select.innerHTML = '<option value="">Выберите преподавателя из списка</option>';
        
        BUILT_IN_TEACHERS.forEach(teacher => {
            const option = document.createElement('option');
            option.value = teacher;
            option.textContent = teacher;
            select.appendChild(option);
        });
    }

    initTeacherSelect();

    document.getElementById('teacherSelect')?.addEventListener('change', function() {
        const selectedTeacher = this.value;
        if (selectedTeacher) {
            document.getElementById('teacherName').value = selectedTeacher;
        }
    });

    function addAllTeachersFromList() {
        let addedCount = 0;
        let skippedCount = 0;
        const baseTime = Date.now();
        
        BUILT_IN_TEACHERS.forEach((teacherName, index) => {
            if (!currentData.teachers.some(t => t.name === teacherName)) {
                const teacher = {
                    id: baseTime + index,
                    name: teacherName,
                    selected: true
                };
                currentData.teachers.push(teacher);
                addedCount++;
            } else {
                skippedCount++;
            }
        });
        
        updateTeachersDisplay();
        
        if (addedCount > 0) {
            showMessageWorkload(`Добавлено преподавателей: ${addedCount}${skippedCount > 0 ? `, пропущено (уже есть): ${skippedCount}` : ''}`, 'success');
        } else {
            showMessageWorkload('Все преподаватели из списка уже добавлены', 'info');
        }
        
        document.getElementById('teacherName').value = '';
        document.getElementById('teacherSelect').value = '';
    }

    function showMessageWorkload(message, type = 'success') {
        const errorDiv = document.getElementById('errorWorkload');
        errorDiv.textContent = message;
        errorDiv.style.display = 'block';
        
        if (type === 'success') {
            errorDiv.style.background = '#27ae60';
        } else if (type === 'info') {
            errorDiv.style.background = '#3498db';
        } else {
            errorDiv.style.background = '#e74c3c';
        }
        
        setTimeout(() => {
            errorDiv.style.display = 'none';
            errorDiv.style.background = '#e74c3c';
        }, 3000);
    }

    function addSelectedTeacher() {
        const select = document.getElementById('teacherSelect');
        const selectedTeacher = select.value;
        
        if (!selectedTeacher) {
            showErrorWorkload('Выберите преподавателя из списка');
            return;
        }
        
        if (currentData.teachers.some(t => t.name === selectedTeacher)) {
            showErrorWorkload('Преподаватель с таким ФИО уже добавлен');
            return;
        }
        
        const teacher = {
            id: Date.now(),
            name: selectedTeacher,
            selected: true
        };
        
        currentData.teachers.push(teacher);
        updateTeachersDisplay();
        
        select.value = '';
        document.getElementById('teacherName').value = '';
    }

    function formatTotalHours(value) {
        if (value === null || value === undefined || value === '') return value;
        if (typeof value === 'number') return value.toString().replace('.', ',');
        if (typeof value === 'string') return value.replace(/\./g, ',');
        return value;
    }

    function parseTotalHours(value) {
        if (value === null || value === undefined || value === '') return value;
        if (typeof value === 'number') return value;
        if (typeof value === 'string') {
            const numStr = value.replace(',', '.');
            const num = parseFloat(numStr);
            return isNaN(num) ? value : num;
        }
        return value;
    }

    function normalizeTotalHours(data) {
        if (!data) return data;
        return data.map(row => {
            if (!Array.isArray(row)) return row;
            const newRow = [...row];
            if (newRow.length > 34) {
                const cell = newRow[34];
                if (typeof cell === 'string' && cell.match(/^\d+[,]\d+$/)) {
                    newRow[34] = parseFloat(cell.replace(',', '.'));
                }
            }
            return newRow;
        });
    }

    function handleFileUploadWorkload() {
        const fileInput = document.getElementById('fileInputWorkload');
        const file = fileInput.files[0];

        if (!file) {
            showErrorWorkload('Пожалуйста, выберите файл');
            return;
        }

        showLoadingWorkload(true);
        document.getElementById('restoreHistoryBtnWorkload').style.display = 'none';
        currentData.faculties = {};
        currentData.assignments = [];
        currentData.subjectNameSettings = {};
        currentData.subgroupSettings = {};
        updateAssignmentsDisplay();

        const reader = new FileReader();

        reader.onload = function (e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {
                    type: 'array',
                    cellStyles: true,
                    cellNF: true,
                    raw: true
                });

                currentData.originalWorkbook = workbook;

                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                currentData.originalSheetData = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1,
                    raw: true,
                    defval: ''
                });

                currentData.originalSheetData = normalizeTotalHours(currentData.originalSheetData);
                extractAllCellStyles(worksheet);
                currentData.headerRows = currentData.originalSheetData.slice(0, 4);

                const jsonData = currentData.originalSheetData;
                processDataWorkload(jsonData);

                document.getElementById('exportOptionsWorkload').style.display = 'block';

            } catch (error) {
                showErrorWorkload('Ошибка при чтении файла: ' + error.message);
                console.error(error);
            }
        };

        reader.onerror = function () {
            showErrorWorkload('Ошибка при чтении файла');
        };

        reader.readAsArrayBuffer(file);
    }

    function extractAllCellStyles(worksheet) {
        currentData.cellStyles = {};
        currentData.cellFormats = {};

        const range = XLSX.utils.decode_range(worksheet['!ref']);

        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = worksheet[cellAddress];

                if (cell) {
                    if (cell.s) {
                        currentData.cellStyles[cellAddress] = JSON.parse(JSON.stringify(cell.s));
                    }
                    if (cell.z) {
                        currentData.cellFormats[cellAddress] = cell.z;
                    }
                    if (cell.t) {
                        if (!currentData.cellStyles[cellAddress]) {
                            currentData.cellStyles[cellAddress] = {};
                        }
                        currentData.cellStyles[cellAddress].t = cell.t;
                        if (cell.v !== undefined) {
                            currentData.cellStyles[cellAddress].originalValue = cell.v;
                        }
                    }
                }
            }
        }

        if (worksheet['!cols']) {
            currentData.columnWidths = JSON.parse(JSON.stringify(worksheet['!cols']));
        }
        if (worksheet['!rows']) {
            currentData.rowHeights = JSON.parse(JSON.stringify(worksheet['!rows']));
        }
        if (worksheet['!merges']) {
            currentData.mergedCells = JSON.parse(JSON.stringify(worksheet['!merges']));
        }
        currentData.range = worksheet['!ref'];
    }

    function processDataWorkload(data) {
        const rows = data.slice(4);
        const faculties = {};
        let currentFaculty = '';
        let currentSubject = null;
        let directionIndex = 0;

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            if (row[3] && !row[4]) {
                currentFaculty = String(row[3]).trim();
                currentSubject = null;
                if (!faculties[currentFaculty]) {
                    faculties[currentFaculty] = [];
                }
            }
            else if (row[3] && row[4] && currentFaculty) {
                currentSubject = {
                    name: String(row[3]).trim(),
                    directions: [],
                    rowIndex: i + 4
                };
                faculties[currentFaculty].push(currentSubject);

                const direction = addDirectionToSubject(currentSubject, row, directionIndex++, currentFaculty, i + 4);

                if (i + 1 < rows.length) {
                    const nextRow = rows[i + 1];
                    const hasNoDirectionCode = !nextRow[4] || String(nextRow[4]).trim() === '';
                    const hasSeminars = (nextRow[9] && nextRow[9] > 0) || (nextRow[17] && nextRow[17] > 0);

                    if (hasNoDirectionCode && hasSeminars) {
                        direction.hasSubgroup = true;
                        const subgroupId = `${direction.id}_subgroup`;
                        
                        direction.subgroupData = {
                            id: subgroupId,
                            groups: nextRow[7] || direction.groups,
                            autumn: {
                                lectures: nextRow[8] !== undefined ? nextRow[8] : 0,
                                seminars: nextRow[9] !== undefined ? nextRow[9] : 0,
                                labs: nextRow[10] !== undefined ? nextRow[10] : 0,
                                attestation: nextRow[11] || direction.autumn.attestation
                            },
                            spring: {
                                lectures: nextRow[12] !== undefined ? nextRow[12] : 0,
                                seminars: nextRow[13] !== undefined ? nextRow[13] : 0,
                                labs: nextRow[14] !== undefined ? nextRow[14] : 0,
                                attestation: nextRow[15] || direction.spring.attestation
                            },
                            currentControl: nextRow[16] || '',
                            total: nextRow[19] !== undefined ? parseTotalHours(nextRow[19]) : direction.total,
                            preExamConsultation: nextRow[17] || '',
                            examOrTest: nextRow[18] || '',
                            rowIndex: i + 5,
                            originalRowData: [...nextRow],
                            subjectName: direction.subjectName,
                            directionCode: direction.code
                        };
                        
                        currentData.subgroupSettings[subgroupId] = false;
                        i++;
                    }
                }
            }
            else if (!row[3] && row[4] && currentFaculty && currentSubject) {
                const direction = addDirectionToSubject(currentSubject, row, directionIndex++, currentFaculty, i + 4);
                currentData.subjectNameSettings[direction.id] = false;

                if (i + 1 < rows.length) {
                    const nextRow = rows[i + 1];
                    const hasNoDirectionCode = !nextRow[4] || String(nextRow[4]).trim() === '';
                    const hasSeminars = (nextRow[9] && nextRow[9] > 0) || (nextRow[17] && nextRow[17] > 0);

                    if (hasNoDirectionCode && hasSeminars) {
                        direction.hasSubgroup = true;
                        const subgroupId = `${direction.id}_subgroup`;
                        
                        direction.subgroupData = {
                            id: subgroupId,
                            groups: nextRow[7] || direction.groups,
                            autumn: {
                                lectures: nextRow[8] !== undefined ? nextRow[8] : 0,
                                seminars: nextRow[9] !== undefined ? nextRow[9] : 0,
                                labs: nextRow[10] !== undefined ? nextRow[10] : 0,
                                attestation: nextRow[11] || direction.autumn.attestation
                            },
                            spring: {
                                lectures: nextRow[12] !== undefined ? nextRow[12] : 0,
                                seminars: nextRow[13] !== undefined ? nextRow[13] : 0,
                                labs: nextRow[14] !== undefined ? nextRow[14] : 0,
                                attestation: nextRow[15] || direction.spring.attestation
                            },
                            currentControl: nextRow[16] || '',
                            total: nextRow[19] !== undefined ? parseTotalHours(nextRow[19]) : direction.total,
                            preExamConsultation: nextRow[17] || '',
                            examOrTest: nextRow[18] || '',
                            rowIndex: i + 5,
                            originalRowData: [...nextRow],
                            subjectName: direction.subjectName,
                            directionCode: direction.code
                        };
                        
                        currentData.subgroupSettings[subgroupId] = false;
                        i++;
                    }
                }
            }
        }

        currentData.faculties = faculties;
        displayDataWorkload(faculties);
    }

    function addDirectionToSubject(subject, row, id, faculty, rowIndex) {
        const direction = {
            id: id,
            faculty: faculty,
            subjectName: subject.name,
            code: String(row[4]).trim(),
            course: row[5],
            studentsCount: row[6],
            groups: row[7],
            autumn: {
                lectures: row[8] !== undefined ? row[8] : 0,
                seminars: row[9] !== undefined ? row[9] : 0,
                labs: row[10] !== undefined ? row[10] : 0,
                attestation: row[11] || ''
            },
            spring: {
                lectures: row[12] !== undefined ? row[12] : 0,
                seminars: row[13] !== undefined ? row[13] : 0,
                labs: row[14] !== undefined ? row[14] : 0,
                attestation: row[15] || ''
            },
            currentControl: row[16] || '',
            total: row[19] !== undefined ? parseTotalHours(row[19]) : 0,
            preExamConsultation: row[17] || '',
            examOrTest: row[18] || '',
            hasSubgroup: false,
            subgroupData: null,
            rowIndex: rowIndex,
            originalRowData: [...row]
        };

        subject.directions.push(direction);
        return direction;
    }

    function toggleSubjectNameSetting(directionId) {
        if (currentData.subjectNameSettings.hasOwnProperty(directionId)) {
            currentData.subjectNameSettings[directionId] = !currentData.subjectNameSettings[directionId];
        } else {
            currentData.subjectNameSettings[directionId] = true;
        }
        updateDirectionCheckboxDisplay(directionId);
    }

    function toggleSubgroupSetting(subgroupId) {
        if (currentData.subgroupSettings.hasOwnProperty(subgroupId)) {
            currentData.subgroupSettings[subgroupId] = !currentData.subgroupSettings[subgroupId];
        } else {
            currentData.subgroupSettings[subgroupId] = true;
        }
        updateSubgroupCheckboxDisplay(subgroupId);
    }

    function updateDirectionCheckboxDisplay(directionId) {
        const checkbox = document.getElementById(`subject_checkbox_${directionId}`);
        if (checkbox) {
            checkbox.checked = currentData.subjectNameSettings[directionId] || false;
        }
    }

    function updateSubgroupCheckboxDisplay(subgroupId) {
        const checkbox = document.getElementById(`subgroup_checkbox_${subgroupId}`);
        if (checkbox) {
            checkbox.checked = currentData.subgroupSettings[subgroupId] || false;
        }
    }

    function addTotalFormulas(ws, totalRows, headerRowsCount = 4) {
        for (let R = headerRowsCount; R < totalRows; R++) {
            const rowNum = R + 1;
            const formula = `=I${rowNum}+J${rowNum}+K${rowNum}+M${rowNum}+N${rowNum}+O${rowNum}+Q${rowNum}+R${rowNum}+S${rowNum}`;
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: 19 });
            
            if (!ws[cellAddress]) {
                ws[cellAddress] = {};
            }
            ws[cellAddress].t = 'n';
            ws[cellAddress].f = formula;
            ws[cellAddress].v = null;
        }
    }

    // ==================== НОВОЕ: Отправка файла на сервер ====================
    async function uploadWorkloadFileToServer(blob, filename) {
        const formData = new FormData();
        formData.append('file', blob, filename);
        formData.append('filename', filename);

        try {
            const response = await fetch('/api/upload-workload-file', {
                method: 'POST',
                body: formData
            });
            if (!response.ok) {
                const error = await response.json();
                console.error('Ошибка сохранения на сервере:', error);
                return false;
            }
            console.log('Файл преподавательской нагрузки сохранён на сервере');
            return true;
        } catch (error) {
            console.error('Ошибка сети при сохранении:', error);
            return false;
        }
    }

    function exportToExcel() {
        if (!currentData.originalSheetData) {
            showErrorWorkload('Сначала загрузите файл');
            return;
        }

        if (currentData.assignments.length === 0) {
            showErrorWorkload('Нет прикрепленных преподавателей для экспорта');
            return;
        }

        const selectedTeachers = currentData.teachers.filter(t => t.selected);
        if (selectedTeachers.length === 0) {
            showErrorWorkload('Выберите хотя бы одного преподавателя для экспорта');
            return;
        }

        const fileNameInput = document.getElementById('exportFileNameWorkload');
        let fileName = fileNameInput.value.trim();

        if (!fileName) {
            fileName = 'Нагрузка_преподавателей';
        }
        if (!fileName.toLowerCase().endsWith('.xlsx')) {
            fileName += '.xlsx';
        }

        showLoadingWorkload(true);

        try {
            const wb = XLSX.utils.book_new();
            const assignmentsByTeacher = {};

            currentData.assignments.forEach(assignment => {
                const teacher = currentData.teachers.find(t => t.id == assignment.teacherId);
                if (teacher && teacher.selected) {
                    if (!assignmentsByTeacher[assignment.teacherId]) {
                        assignmentsByTeacher[assignment.teacherId] = [];
                    }
                    assignmentsByTeacher[assignment.teacherId].push(assignment);
                }
            });

            if (Object.keys(assignmentsByTeacher).length === 0) {
                showLoadingWorkload(false);
                showErrorWorkload('У выбранных преподавателей нет прикрепленных направлений');
                return;
            }

            Object.keys(assignmentsByTeacher).forEach(teacherId => {
                const teacher = currentData.teachers.find(t => t.id == teacherId);
                if (!teacher || !teacher.selected) return;

                const teacherAssignments = assignmentsByTeacher[teacherId];
                const ws = createTeacherSheet(teacher, teacherAssignments);
                let sheetName = cleanSheetName(teacher.name);
                XLSX.utils.book_append_sheet(wb, ws, sheetName);
            });

            // Генерация Blob для скачивания и отправки на сервер
            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

            // Скачивание файла
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            // Отправка на сервер
            uploadWorkloadFileToServer(blob, fileName).then(success => {
                if (success) {
                    showMessageWorkload('Файл успешно экспортирован и сохранён на сервере', 'success');
                } else {
                    showMessageWorkload('Файл экспортирован, но не сохранён на сервере', 'info');
                }
            });

            showLoadingWorkload(false);

        } catch (error) {
            showLoadingWorkload(false);
            showErrorWorkload('Ошибка при экспорте: ' + error.message);
            console.error(error);
        }
    }

    function createTeacherSheet(teacher, assignments) {
        const ws = XLSX.utils.aoa_to_sheet([]);

        if (currentData.columnWidths) ws['!cols'] = currentData.columnWidths;
        if (currentData.rowHeights) ws['!rows'] = currentData.rowHeights;
        if (currentData.mergedCells) ws['!merges'] = currentData.mergedCells;

        let exportData = [];
        let currentRow = 0;

        for (let i = 0; i < currentData.headerRows.length; i++) {
            const headerRow = [];
            for (let col = 0; col < 20; col++) {
                let cellValue = currentData.headerRows[i][col];
                if (cellValue !== undefined && cellValue !== null && cellValue !== "" && cellValue !== " " && cellValue !== "\u00A0") {
                    headerRow[col] = cellValue;
                }
            }
            exportData.push(headerRow);
            currentRow++;
        }

        exportTeacherAssignments(ws, exportData, currentRow, assignments);
        XLSX.utils.sheet_add_aoa(ws, exportData, { origin: 'A1' });
        applySavedStylesToSheet(ws, exportData.length);
        addTotalFormulas(ws, exportData.length, 4);

        return ws;
    }

    function exportTeacherAssignments(ws, exportData, startRow, assignments) {
        const uniqueAssignments = new Map();

        assignments.forEach(assignment => {
            const key = `${assignment.directionId}_${assignment.isSubgroup}`;
            if (!uniqueAssignments.has(key)) {
                uniqueAssignments.set(key, []);
            }
            uniqueAssignments.get(key).push(assignment);
        });

        let rowIndex = startRow;
        let rowsAdded = 0;
        const rowsToExport = [];

        for (const [key, assignmentList] of uniqueAssignments) {
            const assignment = assignmentList[0];
            const directionData = getDirectionData(assignment.directionId);
            if (!directionData) continue;

            let sourceRowIndex = assignment.isSubgroup && directionData.hasSubgroup && directionData.subgroupData
                ? directionData.subgroupData.rowIndex
                : directionData.rowIndex;

            rowsToExport.push({
                sourceRowIndex: sourceRowIndex,
                directionData: directionData,
                isSubgroup: assignment.isSubgroup,
                subjectName: assignment.subjectName,
                directionId: assignment.directionId,
                subgroupId: assignment.isSubgroup && directionData.subgroupData ? directionData.subgroupData.id : null
            });
        }

        rowsToExport.sort((a, b) => a.sourceRowIndex - b.sourceRowIndex);

        for (let i = 0; i < rowsToExport.length; i++) {
            const rowData = rowsToExport[i];
            const originalRow = currentData.originalSheetData[rowData.sourceRowIndex];
            if (!originalRow) continue;
            
            const newRow = [];
            
            for (let col = 0; col < 20 && col < originalRow.length; col++) {
                const value = originalRow[col];
                if (value !== undefined && value !== null && value !== "" && value !== " " && value !== "\u00A0") {
                    newRow[col] = value;
                }
            }
            
            if (!rowData.isSubgroup) {
                const isSubjectNameEmpty = (!originalRow[3] || originalRow[3] === "" || originalRow[3] === null);
                const shouldAddSubjectName = currentData.subjectNameSettings[rowData.directionId] || false;
                
                if (isSubjectNameEmpty && shouldAddSubjectName && rowData.subjectName) {
                    newRow[3] = rowData.subjectName;
                }
            }
            
            if (rowData.isSubgroup) {
                const shouldAddForSubgroup = currentData.subgroupSettings[rowData.subgroupId] || false;
                
                if (shouldAddForSubgroup) {
                    if (rowData.subjectName && (!newRow[3] || newRow[3] === "")) {
                        newRow[3] = rowData.subjectName;
                    }
                    if (rowData.directionData && rowData.directionData.code && 
                        (!originalRow[4] || originalRow[4] === "" || originalRow[4] === null)) {
                        newRow[4] = rowData.directionData.code;
                    }
                }
            }
            
            if (newRow[34] !== undefined) {
                const totalValue = newRow[19];
                const cellAddress = XLSX.utils.encode_cell({ r: rowData.sourceRowIndex, c: 19 });
                const cellFormat = currentData.cellFormats[cellAddress];
                
                if (cellFormat && typeof totalValue === 'string' && totalValue.match(/^[\d\s,]+$/)) {
                    const numValue = parseTotalHours(totalValue);
                    if (typeof numValue === 'number') {
                        newRow[34] = numValue;
                    }
                }
            }

            const hasData = newRow.some(cell => cell !== undefined && cell !== null && cell !== "");
            if (hasData) {
                exportData.push(newRow);
                rowsAdded++;
                copyRowStylesNonEmpty(ws, rowData.sourceRowIndex, rowIndex, originalRow);
                rowIndex++;
            }
        }

        console.log(`Экспортировано ${rowsAdded} строк.`);
        return rowsAdded;
    }

    function copyRowStylesNonEmpty(ws, sourceRowIndex, targetRowIndex, originalRow) {
        for (let col = 0; col < originalRow.length; col++) {
            const value = originalRow[col];
            
            if (value !== undefined && value !== null && value !== "" && value !== " " && value !== "\u00A0") {
                const sourceCellAddress = XLSX.utils.encode_cell({ r: sourceRowIndex, c: col });
                const targetCellAddress = XLSX.utils.encode_cell({ r: targetRowIndex, c: col });

                if (currentData.cellStyles[sourceCellAddress]) {
                    if (!ws[targetCellAddress]) {
                        ws[targetCellAddress] = { v: value };
                    }
                    ws[targetCellAddress].s = JSON.parse(JSON.stringify(currentData.cellStyles[sourceCellAddress]));
                }

                if (currentData.cellFormats[sourceCellAddress]) {
                    if (!ws[targetCellAddress]) {
                        ws[targetCellAddress] = { v: value };
                    }
                    ws[targetCellAddress].z = currentData.cellFormats[sourceCellAddress];
                }
            }
        }
    }

    function copyRowStyles(ws, sourceRowIndex, targetRowIndex, colCount) {
        for (let col = 0; col < colCount; col++) {
            const sourceCellAddress = XLSX.utils.encode_cell({ r: sourceRowIndex, c: col });
            const targetCellAddress = XLSX.utils.encode_cell({ r: targetRowIndex, c: col });

            if (currentData.cellStyles[sourceCellAddress]) {
                if (!ws[targetCellAddress]) {
                    ws[targetCellAddress] = {};
                }
                ws[targetCellAddress].s = JSON.parse(JSON.stringify(currentData.cellStyles[sourceCellAddress]));
            }

            if (currentData.cellFormats[sourceCellAddress]) {
                if (!ws[targetCellAddress]) {
                    ws[targetCellAddress] = {};
                }
                ws[targetCellAddress].z = currentData.cellFormats[sourceCellAddress];
            }
            
            if (ws[targetCellAddress] && ws[targetCellAddress].v === undefined) {
                ws[targetCellAddress].v = null;
            }
        }
    }

    function cleanSheetName(name) {
        let cleaned = name.replace(/[\\/*?[\]:]/g, '');
        cleaned = cleaned.substring(0, 31);
        if (!cleaned.trim()) {
            cleaned = "Преподаватель";
        }
        return cleaned;
    }

    function getDirectionData(directionId) {
        for (const faculty in currentData.faculties) {
            for (const subject of currentData.faculties[faculty]) {
                for (const direction of subject.directions) {
                    if (direction.id == directionId) {
                        return direction;
                    }
                }
            }
        }
        return null;
    }

    function applySavedStylesToSheet(ws, totalRows) {
        for (let R = 0; R < 4 && R < totalRows; R++) {
            for (let C = 0; C < 20; C++) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                const originalAddress = XLSX.utils.encode_cell({ r: R, c: C });

                if (currentData.cellStyles[originalAddress]) {
                    if (!ws[cellAddress]) {
                        ws[cellAddress] = { v: '' };
                    }
                    ws[cellAddress].s = JSON.parse(JSON.stringify(currentData.cellStyles[originalAddress]));
                }

                if (currentData.cellFormats[originalAddress]) {
                    if (!ws[cellAddress]) {
                        ws[cellAddress] = { v: '' };
                    }
                    ws[cellAddress].z = currentData.cellFormats[originalAddress];
                }
            }
        }
    }

    function addTeacher() {
        const teacherNameInput = document.getElementById('teacherName');
        const teacherName = teacherNameInput.value.trim();

        if (!teacherName) {
            showErrorWorkload('Введите ФИО преподавателя');
            return;
        }

        if (currentData.teachers.some(t => t.name === teacherName)) {
            showErrorWorkload('Преподаватель с таким ФИО уже добавлен');
            return;
        }

        const teacher = {
            id: Date.now(),
            name: teacherName,
            selected: true
        };

        currentData.teachers.push(teacher);
        teacherNameInput.value = '';
        updateTeachersDisplay();
    }

    async function assignTeacherToCurrent(teacherId) {
        const assignment = {
            id: Date.now(),
            teacherId: teacherId,
            directionId: currentAssignment.directionId,
            isSubgroup: currentAssignment.isSubgroup,
            subjectName: currentAssignment.subjectName,
            directionCode: currentAssignment.directionCode,
            facultyName: currentAssignment.facultyName,
            assignedAt: new Date().toISOString()
        };

        currentData.assignments.push(assignment);
        updateTeachersDisplay();
        updateDirectionDisplays();
        hideModal();

        // Сохранение на сервер
        const teacher = currentData.teachers.find(t => t.id === teacherId);
        if (teacher) {
            try {
                await fetch('/api/assignments', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        teacher_name: teacher.name.trim(),
                        direction_code: currentAssignment.directionCode.trim(),
                        subject_name: currentAssignment.subjectName.trim(),
                        faculty_name: currentAssignment.facultyName.trim(),
                        is_subgroup: currentAssignment.isSubgroup,
                        course: (currentAssignment.directionData?.course || '').trim()
                    })
                });
            } catch (error) {
                console.error('Ошибка сохранения назначения:', error);
            }
        }
    }
    
    function hideModal() {
        document.getElementById('teacherModal').style.display = 'none';
        currentAssignment = {
            directionId: null,
            isSubgroup: false,
            subjectName: '',
            directionCode: '',
            facultyName: '',
            directionData: null
        };
    }

    async function removeAssignment(assignmentId) {
        const assignment = currentData.assignments.find(a => a.id === assignmentId);
        const teacher = assignment ? currentData.teachers.find(t => t.id === assignment.teacherId) : null;

        currentData.assignments = currentData.assignments.filter(a => a.id !== assignmentId);
        updateTeachersDisplay();
        updateDirectionDisplays();

        if (teacher && assignment) {
            try {
                await fetch('/api/assignments/delete-by-params', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        teacher_name: teacher.name,
                        direction_code: assignment.directionCode,
                        is_subgroup: assignment.isSubgroup
                    })
                });
            } catch (error) {
                console.error('Ошибка удаления назначения:', error);
            }
        }
    }
    

    function updateTeachersDisplay() {
        const teacherList = document.getElementById('teacherList');

        if (currentData.teachers.length === 0) {
            teacherList.innerHTML = '<div class="no-data">Пока нет добавленных преподавателей</div>';
        } else {
            teacherList.innerHTML = `
                <div class="teacher-selection-controls" style="margin-bottom: 15px;">
                    <button onclick="selectAllTeachers()" class="btn btn-small" style="background: #3498db; margin-right: 10px;">Выбрать всех</button>
                    <button onclick="deselectAllTeachers()" class="btn btn-small" style="background: #95a5a6;">Снять выбор</button>
                </div>
                <div class="teacher-columns">
            `;

            currentData.teachers.forEach(teacher => {
                const directionCount = currentData.assignments.filter(a => a.teacherId === teacher.id && !a.isSubgroup).length;
                const subgroupCount = currentData.assignments.filter(a => a.teacherId === teacher.id && a.isSubgroup).length;

                const teacherItem = document.createElement('div');
                teacherItem.className = 'teacher-item';
                teacherItem.innerHTML = `
                    <div style="display: flex; align-items: center; margin-bottom: 8px;">
                        <input type="checkbox" 
                               class="teacher-checkbox" 
                               id="teacher_${teacher.id}"
                               ${teacher.selected ? 'checked' : ''}
                               onchange="toggleTeacherSelection(${teacher.id})"
                               style="margin-right: 10px;">
                        <label for="teacher_${teacher.id}" class="teacher-name" style="flex: 1; cursor: pointer;">
                            ${teacher.name}
                        </label>
                    </div>
                    <div class="teacher-directions">
                        Напр.: ${directionCount}, Подгр.: ${subgroupCount}
                    </div>
                    <button class="btn btn-danger btn-small" onclick="removeTeacher(${teacher.id})">Удалить</button>
                `;
                teacherList.querySelector('.teacher-columns').appendChild(teacherItem);
            });

            teacherList.innerHTML += '</div>';
        }

        updateAssignmentsDisplay();
    }

    function toggleTeacherSelection(teacherId) {
        const teacher = currentData.teachers.find(t => t.id === teacherId);
        if (teacher) {
            teacher.selected = !teacher.selected;
        }
    }

    function selectAllTeachers() {
        currentData.teachers.forEach(teacher => {
            teacher.selected = true;
        });
        updateTeachersDisplay();
    }

    function deselectAllTeachers() {
        currentData.teachers.forEach(teacher => {
            teacher.selected = false;
        });
        updateTeachersDisplay();
    }

    function removeTeacher(teacherId) {
        if (confirm('Удалить преподавателя и все его прикрепления?')) {
            currentData.teachers = currentData.teachers.filter(t => t.id !== teacherId);
            currentData.assignments = currentData.assignments.filter(a => a.teacherId !== teacherId);
            updateTeachersDisplay();
            updateDirectionDisplays();
        }
    }

    function showAssignModal(directionId, isSubgroup = false) {
        if (currentData.teachers.length === 0) {
            showErrorWorkload('Сначала добавьте преподавателей');
            return;
        }

        let directionData = null;
        let facultyName = '';
        let subjectName = '';

        for (const faculty in currentData.faculties) {
            for (const subject of currentData.faculties[faculty]) {
                for (const direction of subject.directions) {
                    if (direction.id == directionId) {
                        directionData = direction;
                        facultyName = faculty;
                        subjectName = subject.name;
                        break;
                    }
                }
                if (directionData) break;
            }
            if (directionData) break;
        }

        if (!directionData) {
            showErrorWorkload('Направление не найдено');
            return;
        }

        currentAssignment = {
            directionId: directionId,
            isSubgroup: isSubgroup,
            subjectName: subjectName,
            directionCode: directionData.code,
            facultyName: facultyName,
            directionData: directionData
        };

        const modalTitle = document.getElementById('modalTitle');
        if (isSubgroup) {
            modalTitle.textContent = `Прикрепить подгруппу к преподавателю`;
        } else {
            modalTitle.textContent = `Прикрепить направление к преподавателю`;
        }
        modalTitle.innerHTML += `<br><small>${subjectName} - ${directionData.code}</small>`;

        const modalTeachersList = document.getElementById('modalTeachersList');
        modalTeachersList.innerHTML = '';

        currentData.teachers.forEach(teacher => {
            const teacherAssignments = currentData.assignments.filter(a => a.teacherId === teacher.id);
            const directionCount = teacherAssignments.filter(a => !a.isSubgroup).length;
            const subgroupCount = teacherAssignments.filter(a => a.isSubgroup).length;

            const alreadyAssigned = currentData.assignments.some(a =>
                a.teacherId === teacher.id &&
                a.directionId == directionId &&
                a.isSubgroup === isSubgroup
            );

            const teacherOption = document.createElement('div');
            teacherOption.className = 'modal-teacher-item';
            if (alreadyAssigned) {
                teacherOption.style.opacity = '0.6';
                teacherOption.style.backgroundColor = '#e8f4fc';
            }

            teacherOption.innerHTML = `
                <div style="font-weight: bold; margin-bottom: 5px;">${teacher.name}</div>
                <div style="font-size: 0.8em; color: #666;">
                    Напр.: ${directionCount}, Подгр.: ${subgroupCount}
                    ${alreadyAssigned ? '<br><span style="color: #27ae60;">✓ Уже прикреплен</span>' : ''}
                </div>
            `;

            if (!alreadyAssigned) {
                teacherOption.addEventListener('click', () => {
                    assignTeacherToCurrent(teacher.id);
                });
                teacherOption.style.cursor = 'pointer';
            } else {
                teacherOption.style.cursor = 'not-allowed';
            }

            modalTeachersList.appendChild(teacherOption);
        });

        document.getElementById('teacherModal').style.display = 'flex';
    }

    function updateAssignmentsDisplay() {
        const assignmentsList = document.getElementById('assignmentsList');
        const assignmentsByTeacher = {};

        currentData.assignments.forEach(assignment => {
            if (!assignmentsByTeacher[assignment.teacherId]) {
                assignmentsByTeacher[assignment.teacherId] = [];
            }
            assignmentsByTeacher[assignment.teacherId].push(assignment);
        });

        if (currentData.assignments.length === 0) {
            assignmentsList.innerHTML = '<div class="no-data">Выберите направление и прикрепите к преподавателю</div>';
        } else {
            assignmentsList.innerHTML = '';

            Object.keys(assignmentsByTeacher).forEach(teacherId => {
                const teacher = currentData.teachers.find(t => t.id == teacherId);
                if (!teacher) return;

                const teacherAssignments = assignmentsByTeacher[teacherId];

                const teacherBox = document.createElement('div');
                teacherBox.className = 'teacher-assignment-box';

                const teacherHeader = document.createElement('div');
                teacherHeader.className = 'teacher-assignment-header';
                teacherHeader.innerHTML = `
                    <div class="teacher-assignment-name">${teacher.name}</div>
                    <div class="teacher-assignment-count">Прикреплено: ${teacherAssignments.length}</div>
                `;

                const assignmentsContainer = document.createElement('div');
                assignmentsContainer.className = 'assignment-items-list';

                teacherAssignments.forEach(assignment => {
                    const assignmentItem = document.createElement('div');
                    assignmentItem.className = 'assignment-item-box';

                    const typeClass = assignment.isSubgroup ? 'assignment-item-type subgroup' : 'assignment-item-type';
                    const typeText = assignment.isSubgroup ? 'Подгруппа' : 'Направление';

                    assignmentItem.innerHTML = `
                        <div class="${typeClass}">${typeText}</div>
                        <div class="assignment-item-subject">${assignment.subjectName}</div>
                        <div class="assignment-item-direction">${assignment.directionCode}</div>
                        <div class="assignment-item-faculty">Факультет: ${assignment.facultyName}</div>
                        <div class="assignment-item-remove" onclick="removeAssignment(${assignment.id})">Удалить</div>
                    `;
                    assignmentsContainer.appendChild(assignmentItem);
                });

                teacherBox.appendChild(teacherHeader);
                teacherBox.appendChild(assignmentsContainer);
                assignmentsList.appendChild(teacherBox);
            });
        }
    }

    function updateDirectionDisplays() {
        document.querySelectorAll('.direction-item').forEach(element => {
            const directionId = element.getAttribute('data-id');
            const directionData = getDirectionData(directionId);
            if (!directionData) return;

            const directionAssignments = currentData.assignments.filter(a => a.directionId == directionId);
            const mainAssignments = directionAssignments.filter(a => !a.isSubgroup);
            const subgroupAssignments = directionAssignments.filter(a => a.isSubgroup);

            updateDirectionTeacherDisplay(element, mainAssignments, false);

            if (directionData.hasSubgroup) {
                const subgroupElement = element.querySelector('.subgroup-info');
                if (subgroupElement) {
                    updateDirectionTeacherDisplay(subgroupElement, subgroupAssignments, true);
                }
            }
        });
    }

    function updateDirectionTeacherDisplay(element, assignments, isSubgroup) {
        const oldDisplay = element.querySelector(isSubgroup ? '.subgroup-teachers' : '.direction-teachers');
        if (oldDisplay) {
            oldDisplay.remove();
        }

        if (assignments && assignments.length > 0) {
            const teachersDisplay = document.createElement('div');
            teachersDisplay.className = isSubgroup ? 'subgroup-teachers' : 'direction-teachers';

            const typeText = isSubgroup ? 'Преподаватели подгруппы:' : 'Преподаватели направления:';
            let teachersHTML = `<span class="teacher-label">${typeText}</span><br>`;

            assignments.forEach(assignment => {
                const teacher = currentData.teachers.find(t => t.id == assignment.teacherId);
                if (teacher) {
                    teachersHTML += `
                        <div style="margin: 5px 0; padding: 5px; background: ${isSubgroup ? '#fff8e1' : '#e8f4fc'}; border-radius: 3px;">
                            ${teacher.name}
                            <button class="btn btn-danger btn-small" onclick="removeAssignment(${assignment.id})" style="margin-left: 10px; float: right;">×</button>
                        </div>
                    `;
                }
            });

            teachersDisplay.innerHTML = teachersHTML;
            element.appendChild(teachersDisplay);
        }
    }

    function displayDataWorkload(faculties) {
        const output = document.getElementById('outputWorkload');
        output.innerHTML = '';

        if (Object.keys(faculties).length === 0) {
            showErrorWorkload('Не удалось найти данные о факультетах и дисциплинах');
            return;
        }

        let totalSubjects = 0;
        let totalDirections = 0;
        let totalHours = 0;
        let totalSubgroups = 0;

        for (const faculty in faculties) {
            faculties[faculty].forEach(subject => {
                totalSubjects++;
                totalDirections += subject.directions.length;
                subject.directions.forEach(direction => {
                    totalHours += parseFloat(direction.total) || 0;
                    if (direction.hasSubgroup) {
                        totalSubgroups++;
                        if (direction.subgroupData && direction.subgroupData.total) {
                            totalHours += parseFloat(direction.subgroupData.total) || 0;
                        }
                    }
                });
            });
        }

        const statsHTML = `
            <div class="stats">
                <p>Всего факультетов: ${Object.keys(faculties).length}</p>
                <p>Всего дисциплин: ${totalSubjects}</p>
                <p>Всего направлений: ${totalDirections} (из них с подгруппами: ${totalSubgroups})</p>
                <p>Общее количество часов: ${formatTotalHours(totalHours)}</p>
            </div>
        `;

        output.innerHTML = statsHTML;

        for (const faculty in faculties) {
            const facultySection = document.createElement('div');
            facultySection.className = 'faculty-section';

            let facultySubjectCount = 0;
            let facultyDirectionCount = 0;

            faculties[faculty].forEach(subject => {
                facultySubjectCount++;
                facultyDirectionCount += subject.directions.length;
            });

            const facultyHeader = document.createElement('div');
            facultyHeader.className = 'faculty-header';
            facultyHeader.innerHTML = `
                ${faculty}
                <span>${facultySubjectCount} дисциплин, ${facultyDirectionCount} групп</span>
            `;

            const facultyContent = document.createElement('div');
            facultyContent.className = 'faculty-content';

            faculties[faculty].forEach(subject => {
                const subjectItem = document.createElement('div');
                subjectItem.className = 'subject-item';

                let subjectHTML = `
                    <div class="subject-name">
                        ${subject.name}
                        <span class="directions-count">${subject.directions.length} направления</span>
                    </div>
                `;

                subject.directions.forEach((direction, index) => {
                    const directionClass = direction.hasSubgroup ? 'direction-item has-subgroup' : 'direction-item';
                    const isSubjectNameEmpty = !direction.originalRowData[3] || direction.originalRowData[3] === '';
                    const subjectNameSetting = currentData.subjectNameSettings[direction.id] || false;

                    subjectHTML += `
                        <div class="${directionClass}" data-id="${direction.id}">
                            <button class="assign-btn" onclick="showAssignModal('${direction.id}', false)">
                                Прикрепить направление
                            </button>
                    `;
                    
                    if (isSubjectNameEmpty) {
                        subjectHTML += `
                            <div style="margin-bottom: 10px; padding: 5px; background: #f0f7ff; border-radius: 3px; border-left: 3px solid #3498db;">
                                <input type="checkbox" 
                                       id="subject_checkbox_${direction.id}" 
                                       onchange="toggleSubjectNameSetting('${direction.id}')"
                                       ${subjectNameSetting ? 'checked' : ''}
                                       style="margin-right: 8px; cursor: pointer;">
                                <label for="subject_checkbox_${direction.id}" style="cursor: pointer; color: #2c3e50;">
                                    <strong>Добавить название дисциплины при экспорте</strong>
                                    <span style="display: block; font-size: 0.85em; color: #666; margin-top: 3px;">
                                        (В исходном файле название дисциплины отсутствует)
                                    </span>
                                </label>
                            </div>
                        `;
                    }
                    
                    subjectHTML += `
                            <div class="detail-item">
                                <span class="detail-label">Направление ${index + 1}:</span> ${direction.code}
                            </div>
                            <div class="subject-details">
                                <div class="detail-item">
                                    <span class="detail-label">Курс:</span> ${direction.course}
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Студентов:</span> ${direction.studentsCount}
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Групп/подгрупп:</span> ${direction.groups || ' '}
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Осенний семестр:</span> 
                                    Лекц.: ${direction.autumn.lectures}, 
                                    Сем., практ.: ${direction.autumn.seminars}, 
                                    Лаб.: ${direction.autumn.labs},
                                    Форма пром. аттест: ${direction.autumn.attestation}
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Весенний семестр:</span> 
                                    Лекц.: ${direction.spring.lectures}, 
                                    Сем., практ.: ${direction.spring.seminars}, 
                                    Лаб.: ${direction.spring.labs},
                                    Форма пром. аттест: ${direction.spring.attestation}
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Текущий контроль:</span> ${direction.currentControl || ' '}
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Итого часов:</span> ${formatTotalHours(direction.total)}
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Предэкз. конс.:</span> ${direction.preExamConsultation || ' '}
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Экзамен, зачет, К:</span> ${direction.examOrTest || ' '}
                                </div>
                            </div>
                    `;

                    if (direction.hasSubgroup && direction.subgroupData) {
                        const subgroupSetting = currentData.subgroupSettings[direction.subgroupData.id] || false;
                        
                        subjectHTML += `
                            <div class="subgroup-info">
                                <div style="margin-bottom: 10px; padding: 5px; background: #fff3e0; border-radius: 3px; border-left: 3px solid #f39c12;">
                                    <input type="checkbox" 
                                           id="subgroup_checkbox_${direction.subgroupData.id}" 
                                           onchange="toggleSubgroupSetting('${direction.subgroupData.id}')"
                                           ${subgroupSetting ? 'checked' : ''}
                                           style="margin-right: 8px; cursor: pointer;">
                                    <label for="subgroup_checkbox_${direction.subgroupData.id}" style="cursor: pointer; color: #2c3e50;">
                                        <strong>Добавить название дисциплины и направления при экспорте</strong>
                                    </label>
                                </div>
                                
                                <span class="subgroup-label">Информация о подгруппе:</span>
                                <div class="subject-details">
                                    <div class="detail-item">
                                        <span class="detail-label">Групп/подгрупп:</span> ${direction.subgroupData.groups || ' '}
                                    </div>
                                    <div class="detail-item">
                                        <span class="detail-label">Осенний семестр (подгруппа):</span> 
                                        Лекц.: ${direction.subgroupData.autumn.lectures}, 
                                        Сем., практ.: ${direction.subgroupData.autumn.seminars}, 
                                        Лаб.: ${direction.subgroupData.autumn.labs},
                                        Форма пром. аттест: ${direction.subgroupData.autumn.attestation}
                                    </div>
                                    <div class="detail-item">
                                        <span class="detail-label">Весенний семестр (подгруппа):</span> 
                                        Лекц.: ${direction.subgroupData.spring.lectures}, 
                                        Сем., практ.: ${direction.subgroupData.spring.seminars}, 
                                        Лаб.: ${direction.subgroupData.spring.labs},
                                        Форма пром. аттест: ${direction.subgroupData.spring.attestation}
                                    </div>
                                    <div class="detail-item">
                                        <span class="detail-label">Текущий контроль (подгруппа):</span> ${direction.subgroupData.currentControl || ' '}
                                    </div>
                                    <div class="detail-item">
                                        <span class="detail-label">Итого часов (подгруппа):</span> ${formatTotalHours(direction.subgroupData.total)}
                                    </div>
                                    <div class="detail-item">
                                        <span class="detail-label">Предэкз. конс.(подгруппа):</span> ${direction.subgroupData.preExamConsultation || ' '}
                                    </div>
                                    <div class="detail-item">
                                        <span class="detail-label">Экзамен,зачет, К (подгруппа):</span> ${direction.subgroupData.examOrTest || ' '}
                                    </div>
                                </div>
                                <button class="subgroup-assign-btn" onclick="showAssignModal('${direction.id}', true)">
                                    Прикрепить подгруппу
                                </button>
                            </div>
                        `;
                    }

                    subjectHTML += `</div>`;
                });

                subjectItem.innerHTML = subjectHTML;
                facultyContent.appendChild(subjectItem);
            });

            facultyHeader.addEventListener('click', function () {
                const content = this.nextElementSibling;
                content.style.display = content.style.display === 'block' ? 'none' : 'block';
            });

            facultySection.appendChild(facultyHeader);
            facultySection.appendChild(facultyContent);
            output.appendChild(facultySection);
        }

        updateDirectionDisplays();
        updateAssignmentsDisplay();
        document.getElementById('restoreHistoryBtnWorkload').style.display = 'inline-block';
        showLoadingWorkload(false);
    }

    async function applyHistoricalAssignments() {
        try {
            const response = await fetch('/api/assignments');
            if (!response.ok) throw new Error('Ошибка загрузки истории');
            const saved = await response.json();

            if (!saved.length) {
                showMessageWorkload('Нет сохранённых назначений для восстановления', 'info');
                return;
            }

             // Подготовим карту для быстрого поиска направлений по ключу: faculty|subject|code
            const directionMap = new Map();
            for (const facultyName in currentData.faculties) {
                const subjects = currentData.faculties[facultyName];
                for (const subject of subjects) {
                    const subjectNameClean = (subject.name || '').trim().toLowerCase();
                    for (const dir of subject.directions) {
                        const codeClean = (dir.code || '').trim().toLowerCase();
                        const key = `${facultyName.trim().toLowerCase()}|${subjectNameClean}|${codeClean}`;
                        directionMap.set(key, dir);
                    }
                }
            }
            let restored = 0;
            let notFound = 0;
            let skippedExists = 0;

             // Обрабатываем каждое сохранённое назначение
            for (const a of saved) {
                const facultyClean = (a.faculty_name || '').trim().toLowerCase();
                const subjectClean = (a.subject_name || '').trim().toLowerCase();
                const codeClean = (a.direction_code || '').trim().toLowerCase();
                const key = `${facultyClean}|${subjectClean}|${codeClean}`;

                const directionData = directionMap.get(key);
                if (!directionData) {
                    notFound++;
                    continue;
                }
                // Проверяем, существует ли подгруппа для is_subgroup = true
                if (a.is_subgroup && (!directionData.hasSubgroup || !directionData.subgroupData)) {
                    notFound++;
                    continue;
                }

                // Добавляем преподавателя, если его ещё нет
                let teacher = currentData.teachers.find(t => t.name.trim().toLowerCase() === a.teacher_name.trim().toLowerCase());
                if (!teacher) {
                    teacher = {
                        id: Date.now() + Math.random(),
                        name: a.teacher_name.trim(),
                        selected: true
                    };
                    currentData.teachers.push(teacher);
                }

                // Проверяем дубликат назначения
                const exists = currentData.assignments.some(
                    asgn => asgn.teacherId === teacher.id &&
                            asgn.directionId == directionData.id &&
                            asgn.isSubgroup === a.is_subgroup
                );
                if (exists) {
                    skippedExists++;
                    continue;
                }
                // Создаём новое назначение
                currentData.assignments.push({
                    id: a.id,
                    teacherId: teacher.id,
                    directionId: directionData.id,
                    isSubgroup: a.is_subgroup,
                    subjectName: directionData.subjectName,
                    directionCode: directionData.code,
                    facultyName: a.faculty_name || directionData.faculty,
                    assignedAt: new Date().toISOString()
                });
                restored++;
            }
            // Обновляем интерфейс
            updateTeachersDisplay();
            updateDirectionDisplays();
            updateAssignmentsDisplay();

            let message = `Восстановлено назначений: ${restored}`;
            if (notFound) message += `, не найдено в текущем файле: ${notFound}`;
            if (skippedExists) message += `, уже существует: ${skippedExists}`;
            showMessageWorkload(message, 'success');

        } catch (error) {
            showErrorWorkload('Ошибка восстановления истории: ' + error.message);
        }
    }

    function showLoadingWorkload(show) {
        document.getElementById('loadingWorkload').style.display = show ? 'block' : 'none';
    }

    function showErrorWorkload(message) {
        const errorDiv = document.getElementById('errorWorkload');
        errorDiv.textContent = message;
        errorDiv.style.display = 'block';
        errorDiv.style.background = '#e74c3c';
        showLoadingWorkload(false);
        setTimeout(() => {
            errorDiv.style.display = 'none';
        }, 5000);
    }

    function showWorkloadPreview() {
    if (currentData.assignments.length === 0) {
        showErrorWorkload('Нет прикрепленных направлений для предпросмотра');
        return;
    }

    const selectedTeachers = currentData.teachers.filter(t => t.selected);
    if (selectedTeachers.length === 0) {
        showErrorWorkload('Выберите хотя бы одного преподавателя');
        return;
    }

    const previewMap = new Map(); // teacherId -> rows[]

    currentData.assignments.forEach(assignment => {
        const teacher = selectedTeachers.find(t => t.id == assignment.teacherId);
        if (!teacher) return;

        const directionData = getDirectionData(assignment.directionId);
        if (!directionData) return;

        let sourceRow;
        if (assignment.isSubgroup && directionData.subgroupData) {
            sourceRow = directionData.subgroupData.originalRowData;
        } else {
            sourceRow = directionData.originalRowData;
        }

        const colMap = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19];
        const row = colMap.map(col => {
            let val = sourceRow[col] !== undefined ? sourceRow[col] : '';
            if (col === 3 && !assignment.isSubgroup) {
                const isEmpty = !sourceRow[3] || sourceRow[3] === '';
                if (isEmpty && currentData.subjectNameSettings[assignment.directionId]) {
                    val = assignment.subjectName;
                }
            }
            if (assignment.isSubgroup && col === 3) {
                if (currentData.subgroupSettings[directionData.subgroupData?.id]) {
                    if (!val || val === '') val = assignment.subjectName;
                }
            }
            if (assignment.isSubgroup && col === 4) {
                if (currentData.subgroupSettings[directionData.subgroupData?.id]) {
                    if (!val || val === '') val = directionData.code;
                }
            }
            return val;
        });

        if (!previewMap.has(teacher.id)) {
            previewMap.set(teacher.id, []);
        }
        previewMap.get(teacher.id).push({
            teacher,
            row,
            isSubgroup: assignment.isSubgroup
        });
    });

    let html = '';
    for (const [teacherId, rows] of previewMap.entries()) {
        const teacher = selectedTeachers.find(t => t.id == teacherId);
        html += `<div class="preview-teacher-title">Преподаватель: ${teacher.name}</div>`;
        html += `<table class="preview-table">
            <thead>
                <tr>
                    <th>Наименование дисциплин/практик,
                     руководство курсовыми/НИР/ВКР, участие в ГИА и др.</th>
                    <th>Код и наименование направления,
                    направленность,
                    форма обучения
                    (очная/очно-заочная/заочная)</th>
                    <th>Курс</th>
                    <th>Кол-во студ.</th>
                    <th>Групп/подгр.</th>
                    <th>Осень Лекц.</th>
                    <th>Осень Сем., Практ.</th>
                    <th>Осень Лаб.</th>
                    <th>Осень Форма пром. аттест.</th>
                    <th>Весна Лекц.</th>
                    <th>Весна Сем., Практ.</th>
                    <th>Весна Лаб.</th>
                    <th>Весна Форма пром. аттест.</th>
                    <th>Текущий контроль</th>
                    <th>Предэкз. конс.</th>
                    <th>Экзамен, зачет, К</th>
                    <th>Итого</th>
                </tr>
            </thead>
            <tbody>`;
        rows.forEach(r => {
            const cls = r.isSubgroup ? 'subgroup-row' : '';
            html += `<tr class="${cls}">`;
            r.row.forEach(cell => html += `<td>${cell}</td>`);
            html += `</tr>`;
        });
        html += `</tbody></table>`;
    }

    document.getElementById('previewTitle').textContent = 'Предпросмотр нагрузки преподавателей';
    document.getElementById('previewContent').innerHTML = html;
    document.getElementById('previewModal').style.display = 'flex';
    }
    

    window.addEventListener('click', function (event) {
        const modal = document.getElementById('teacherModal');
        if (event.target === modal) {
            hideModal();
        }
        const previewModal = document.getElementById('previewModal');
        if (event.target === previewModal) {
            previewModal.style.display = 'none';
        }
        const fileSelectModal = document.getElementById('fileSelectModal');
        if (event.target === fileSelectModal) {
            fileSelectModal.style.display = 'none';
        }
    });

    // Глобальные функции
    window.toggleSubjectNameSetting = toggleSubjectNameSetting;
    window.toggleSubgroupSetting = toggleSubgroupSetting;
    window.showAssignModal = showAssignModal;
    window.removeAssignment = removeAssignment;
    window.toggleTeacherSelection = toggleTeacherSelection;
    window.selectAllTeachers = selectAllTeachers;
    window.deselectAllTeachers = deselectAllTeachers;
    window.removeTeacher = removeTeacher;
    window.showWorkloadPreview = showWorkloadPreview;
    window.applyHistoricalAssignments = applyHistoricalAssignments;

})();

// ==================== ПРИЛОЖЕНИЕ 2: РАСПРЕДЕЛЕНИЕ ПО КАФЕДРАМ ====================
(function() {
    const SERVER_URL = 'http://localhost:5000';

    let coursesData = {};
    let currentViewMode = 'course';
    let currentDeptForAppend = null;
    const deptCourseFilter = {};

    const fileInputDisciplines = document.getElementById('fileInputDisciplines');
    const loadBtnDisciplines = document.getElementById('loadBtnDisciplines');
    const loadingDisciplines = document.getElementById('loadingDisciplines');
    const errorDisciplines = document.getElementById('errorDisciplines');
    const outputDisciplines = document.getElementById('outputDisciplines');
    const viewToggleDisciplines = document.getElementById('viewToggleDisciplines');
    const viewByCourseBtn = document.getElementById('viewByCourseBtn');
    const viewByDepartmentBtn = document.getElementById('viewByDepartmentBtn');

    loadBtnDisciplines.addEventListener('click', handleFileUploadDisciplines);
    viewByCourseBtn.addEventListener('click', () => switchView('course'));
    viewByDepartmentBtn.addEventListener('click', () => switchView('department'));

    function handleFileUploadDisciplines() {
        const file = fileInputDisciplines.files[0];
        if (!file) {
            showErrorDisciplines('Пожалуйста, выберите файл');
            return;
        }

        showLoadingDisciplines(true);
        coursesData = {};

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                workbook.SheetNames.forEach(sheetName => {
                    if (sheetName.includes('Курс')) {
                        const worksheet = workbook.Sheets[sheetName];
                        const sheetData = XLSX.utils.sheet_to_json(worksheet, {
                            header: 1,
                            defval: '',
                            blankrows: false
                        });
                        parseSheet(sheetData, sheetName);
                    }
                });

                if (Object.keys(coursesData).length === 0) {
                    showErrorDisciplines('Не найдено ни одного листа с названием "Курс 1–4"');
                    showLoadingDisciplines(false);
                    return;
                }

                currentViewMode = 'course';
                displayDataDisciplines();
                viewToggleDisciplines.style.display = 'flex';
                showLoadingDisciplines(false);
            } catch (err) {
                showErrorDisciplines('Ошибка при чтении файла: ' + err.message);
                console.error(err);
                showLoadingDisciplines(false);
            }
        };

        reader.onerror = () => {
            showErrorDisciplines('Ошибка чтения файла');
            showLoadingDisciplines(false);
        };

        reader.readAsArrayBuffer(file);
    }

    function parseSheet(rows, sheetName) {
        const courseNumberMatch = sheetName.match(/\d+/);
        if (!courseNumberMatch) return;
        const targetCourse = parseInt(courseNumberMatch[0], 10);

        const disciplines = [];
        let isDisciplineSection = false;
        let isPracticeSection = false;
        let skipUntilPractices = false;

        rows.forEach(row => {
            const firstCell = row[0] !== undefined ? String(row[0]).trim() : '';
            const thirdCell = row[2] !== undefined ? String(row[2]).trim() : '';

            if (!isDisciplineSection && !isPracticeSection && !skipUntilPractices) {
                if (thirdCell.includes('ДИСЦИПЛИНЫ (МОДУЛИ)')) {
                    isDisciplineSection = true;
                    return;
                }
                return;
            }

            if (skipUntilPractices) {
                if (firstCell === '-2' && thirdCell.includes('ПРАКТИКИ')) {
                    skipUntilPractices = false;
                    isPracticeSection = true;
                }
                return;
            }

            if (isDisciplineSection) {
                if (firstCell === '-2') {
                    if (thirdCell.includes('ФОРМЫ')) {
                        isDisciplineSection = false;
                        skipUntilPractices = true;
                        return;
                    }
                    if (thirdCell.includes('ПРАКТИКИ')) {
                        isDisciplineSection = false;
                        isPracticeSection = true;
                        return;
                    }
                }

                const nameCell = row[4];
                if (!nameCell || String(nameCell).trim() === '') return;

                const rowCourse = parseInt(firstCell, 10);
                if (!isNaN(rowCourse) && rowCourse !== targetCourse) return;

                const deptCode = row[42] !== undefined ? String(row[42]).trim() : '';
                const deptFullName = row[43] !== undefined ? String(row[43]).trim() : '';

                disciplines.push({
                    name: String(nameCell).trim(),
                    isPractice: false,
                    controlAutumn: row[6] ? String(row[6]).trim() : '',
                    controlSpring: row[18] ? String(row[18]).trim() : '',
                    lecAutumn: (row[9] !== undefined && row[9] !== '' && row[9] !== 0) ? row[9] : '',
                    labAutumn: (row[10] !== undefined && row[10] !== '' && row[10] !== 0) ? row[10] : '',
                    pracAutumn: (row[11] !== undefined && row[11] !== '' && row[11] !== 0) ? row[11] : '',
                    lecSpring: (row[21] !== undefined && row[21] !== '' && row[21] !== 0) ? row[21] : '',
                    labSpring: (row[22] !== undefined && row[22] !== '' && row[22] !== 0) ? row[22] : '',
                    pracSpring: (row[23] !== undefined && row[23] !== '' && row[23] !== 0) ? row[23] : '',
                    departmentCode: deptCode,
                    departmentFullName: deptFullName
                });
                return;
            }

            if (isPracticeSection) {
                if (firstCell === '-2' && thirdCell.includes('ГОСУДАРСТВЕННАЯ')) {
                    isPracticeSection = false;
                    return;
                }

                const rowCourse = parseInt(firstCell, 10);
                if (!isNaN(rowCourse) && rowCourse !== targetCourse) return;

                let practiceName = '';
                if (row[4] && String(row[4]).trim() !== '') {
                    practiceName = String(row[4]).trim();
                } else if (row[3] && String(row[3]).trim() !== '') {
                    practiceName = String(row[3]).trim();
                } else if (row[5] && String(row[5]).trim() !== '') {
                    practiceName = String(row[5]).trim();
                }

                if (!practiceName || practiceName.includes('ПРАКТИКИ') || practiceName.includes('ГОСУДАРСТВЕННАЯ')) return;

                const deptCode = row[42] !== undefined ? String(row[42]).trim() : '';
                const deptFullName = row[43] !== undefined ? String(row[43]).trim() : '';

                disciplines.push({
                    name: practiceName,
                    isPractice: true,
                    controlAutumn: row[6] ? String(row[6]).trim() : '',
                    controlSpring: row[18] ? String(row[18]).trim() : '',
                    pracAutumn: row[11] || '',
                    pracSpring: row[23] || '',
                    lecAutumn: '', labAutumn: '', lecSpring: '', labSpring: '',
                    departmentCode: deptCode,
                    departmentFullName: deptFullName
                });
            }
        });

        if (disciplines.length > 0) {
            coursesData[sheetName] = disciplines;
        }
    }

    function switchView(mode) {
        currentViewMode = mode;
        if (mode === 'course') {
            displayDataDisciplines();
        } else {
            displayByDepartment();
        }
    }

    function displayDataDisciplines() {
        let totalDisciplines = 0;
        let totalPractices = 0;
        let html = '<div class="stats">';

        for (const course in coursesData) {
            const disciplines = coursesData[course];
            const practices = disciplines.filter(d => d.isPractice).length;
            totalDisciplines += disciplines.length - practices;
            totalPractices += practices;
        }
        
        html += `<p>Всего курсов: ${Object.keys(coursesData).length}</p>`;
        html += `<p>Всего дисциплин: ${totalDisciplines}</p>`;
        html += `<p>Всего практик: ${totalPractices}</p>`;
        html += '</div>';

        const sortedCourses = Object.keys(coursesData).sort((a, b) => {
            const numA = parseInt(a.match(/\d+/)?.[0] || 0);
            const numB = parseInt(b.match(/\d+/)?.[0] || 0);
            return numA - numB;
        });

        sortedCourses.forEach(course => {
            const disciplines = coursesData[course];
            const regularDisciplines = disciplines.filter(d => !d.isPractice);
            const practices = disciplines.filter(d => d.isPractice);
            
            html += `
                <div class="course-section">
                    <div class="course-header" onclick="toggleCourse(this)">
                        <span class="course-title">${course}</span>
                        <div class="course-stats">
                            ${regularDisciplines.length > 0 ? `<span class="course-stat">${regularDisciplines.length} дисциплин</span>` : ''}
                            ${practices.length > 0 ? `<span class="course-stat">${practices.length} практик</span>` : ''}
                        </div>
                    </div>
                    <div class="course-content">
            `;
            
            if (regularDisciplines.length > 0) {
                html += `
                        <h3 class="subsection-title">Дисциплины</h3>
                        <table class="discipline-table">
                            <thead>
                                <tr>
                                    <th rowspan="2">№</th>
                                    <th rowspan="2">Наименование дисциплины</th>
                                    <th colspan="4" style="background-color: #e67e22;">Осенний семестр</th>
                                    <th colspan="4" style="background-color: #27ae60;">Весенний семестр</th>
                                    <th rowspan="2">Код каф.</th>
                                    <th rowspan="2">Наименование кафедры</th>
                                </tr>
                                <tr>
                                    <th>Контроль</th>
                                    <th>Лек</th>
                                    <th>Лаб</th>
                                    <th>Пр</th>
                                    <th>Контроль</th>
                                    <th>Лек</th>
                                    <th>Лаб</th>
                                    <th>Пр</th>
                                </tr>
                            </thead>
                            <tbody>
                `;

                regularDisciplines.forEach((discipline, idx) => {
                    const controlAutumnDisplay = discipline.controlAutumn || '';
                    const controlSpringDisplay = discipline.controlSpring || '';
                    const deptCode = discipline.departmentCode || '—';
                    const deptFull = discipline.departmentFullName || '—';
                    
                    html += `
                        <tr>
                            <td>${idx + 1}</td>
                            <td>${discipline.name}</td>
                            <td class="autumn-control">${controlAutumnDisplay}</td>
                            <td class="autumn-lecture">${discipline.lecAutumn}</td>
                            <td class="autumn-lab">${discipline.labAutumn}</td>
                            <td class="autumn-practice">${discipline.pracAutumn}</td>
                            <td class="spring-control">${controlSpringDisplay}</td>
                            <td class="spring-lecture">${discipline.lecSpring}</td>
                            <td class="spring-lab">${discipline.labSpring}</td>
                            <td class="spring-practice">${discipline.pracSpring}</td>
                            <td>${deptCode}</td>
                            <td>${deptFull}</td>
                        </tr>
                    `;
                });

                html += `</tbody></table>`;
            }
            
            if (practices.length > 0) {
                html += `
                        <h3 class="subsection-title" style="margin-top: 20px;">Практики</h3>
                        <table class="discipline-table practice-table">
                            <thead>
                                <tr>
                                    <th>№</th>
                                    <th>Наименование практики</th>
                                    <th>Контроль</th>
                                    <th>Код каф.</th>
                                    <th>Наименование кафедры</th>
                                </tr>
                            </thead>
                            <tbody>
                `;

                practices.forEach((practice, idx) => {
                    const control = practice.controlAutumn || practice.controlSpring || 'ЗаО';
                    const deptCode = practice.departmentCode || '—';
                    const deptFull = practice.departmentFullName || '—';
                    
                    html += `
                        <tr class="practice-row">
                            <td>${idx + 1}</td>
                            <td><strong>${practice.name}</strong></td>
                            <td>${control}</td>
                            <td>${deptCode}</td>
                            <td>${deptFull}</td>
                        </tr>
                    `;
                });

                html += `</tbody></table>`;
            }

            html += `</div></div>`;
        });

        outputDisciplines.innerHTML = html;

        const firstCourse = document.querySelector('.course-content');
        if (firstCourse) firstCourse.style.display = 'none';
    }

    function displayByDepartment() {
        const deptMap = new Map();

        for (const course in coursesData) {
            coursesData[course].forEach(item => {
                const code = item.departmentCode || 'Без кафедры';
                const fullName = item.departmentFullName || (code === 'Без кафедры' ? '—' : code);
                
                if (!deptMap.has(code)) {
                    deptMap.set(code, { fullName, disciplines: [], practices: [] });
                }
                const deptData = deptMap.get(code);
                
                const itemWithCourse = { ...item, course };
                
                if (item.isPractice) {
                    deptData.practices.push(itemWithCourse);
                } else {
                    deptData.disciplines.push(itemWithCourse);
                }
            });
        }

        let html = '<div class="stats"><p>Группировка по кафедрам</p></div>';
        
        const sortedDepts = Array.from(deptMap.entries()).sort((a, b) => {
            const aNum = parseInt(a[0]);
            const bNum = parseInt(b[0]);
            if (!isNaN(aNum) && !isNaN(bNum)) return aNum - bNum;
            return a[0].localeCompare(b[0]);
        });

        sortedDepts.forEach(([code, data]) => {
            const { fullName, disciplines, practices } = data;
            const totalItems = disciplines.length + practices.length;
            if (totalItems === 0) return;

            html += `
                <div class="department-section" data-dept-code="${code}" data-dept-name="${fullName.replace(/"/g, '&quot;')}">
                    <div class="department-header" onclick="toggleDepartment(this)">
                        <div>
                            <span class="department-code">Кафедра ${code}</span>
                            <span class="department-name">${fullName}</span>
                        </div>
                        <div class="department-stats">
                            <span>${disciplines.length} дисц. / ${practices.length} практ.</span>
                            <button class="btn-export-dept" onclick="event.stopPropagation(); exportDepartmentExcel('${code}', '${fullName.replace(/'/g, "\\'")}')" title="Скачать Excel для этой кафедры"> 📥 Excel
                            </button>
                            <button class="btn-preview-dept" onclick="event.stopPropagation(); showDepartmentPreview('${code}', '${fullName.replace(/'/g, "\\'")}')" title="Предпросмотр данных кафедры"> 👁 Предпросмотр
                            </button>
                            <button class="btn-append-dept" onclick="event.stopPropagation(); showFileSelectModal('${code}', '${fullName.replace(/'/g, "\\'")}')" title="Добавить данные в существующий файл"> 📎 Добавить в файл
                            </button>
                        </div>
                        <div class="dept-course-filter" style="margin-top:5px;" onclick="event.stopPropagation();">
                            <span style="font-weight:bold; margin-right:8px;">Курс:</span>
                            <button class="btn btn-small dept-course-btn active" data-course="all" onclick="setDeptCourse('${code.replace(/'/g, "\\'")}', 'all')">Все</button>
                            <button class="btn btn-small dept-course-btn" data-course="Курс 1" onclick="setDeptCourse('${code.replace(/'/g, "\\'")}', 'Курс 1')">Курс 1</button>
                            <button class="btn btn-small dept-course-btn" data-course="Курс 2" onclick="setDeptCourse('${code.replace(/'/g, "\\'")}', 'Курс 2')">Курс 2</button>
                            <button class="btn btn-small dept-course-btn" data-course="Курс 3" onclick="setDeptCourse('${code.replace(/'/g, "\\'")}', 'Курс 3')">Курс 3</button>
                            <button class="btn btn-small dept-course-btn" data-course="Курс 4" onclick="setDeptCourse('${code.replace(/'/g, "\\'")}', 'Курс 4')">Курс 4</button>
                        </div>
                    </div>
                    <div class="department-content">
            `;

            if (disciplines.length > 0) {
                html += `
                        <h4>Дисциплины</h4>
                        <table class="discipline-table">
                            <thead>
                                <tr>
                                    <th>Курс</th>
                                    <th>Наименование</th>
                                    <th colspan="4" style="background-color: #e67e22;">Осенний семестр</th>
                                    <th colspan="4" style="background-color: #27ae60;">Весенний семестр</th>
                                </tr>
                                <tr>
                                    <th></th>
                                    <th></th>
                                    <th>Контроль</th><th>Лек</th><th>Лаб</th><th>Пр</th>
                                    <th>Контроль</th><th>Лек</th><th>Лаб</th><th>Пр</th>
                                </tr>
                            </thead>
                            <tbody>
                `;
                disciplines.forEach(d => {
                    html += `<tr data-course="${d.course}">
                        <td>${d.course}</td>
                        <td>${d.name}</td>
                        <td class="autumn-control">${d.controlAutumn || ''}</td>
                        <td class="autumn-lecture">${d.lecAutumn}</td>
                        <td class="autumn-lab">${d.labAutumn}</td>
                        <td class="autumn-practice">${d.pracAutumn}</td>
                        <td class="spring-control">${d.controlSpring || ''}</td>
                        <td class="spring-lecture">${d.lecSpring}</td>
                        <td class="spring-lab">${d.labSpring}</td>
                        <td class="spring-practice">${d.pracSpring}</td>
                    </tr>`;
                });
                html += `</tbody></table>`;
            }

            if (practices.length > 0) {
                html += `
                        <h4 style="margin-top:15px;">Практики</h4>
                        <table class="discipline-table practice-table">
                            <thead>
                                <tr>
                                    <th>Курс</th>
                                    <th>Наименование</th>
                                    <th>Контроль</th>
                                </tr>
                            </thead>
                            <tbody>
                `;
                practices.forEach(p => {
                    const ctrl = p.controlAutumn || p.controlSpring || 'ЗаО';
                    html += `<tr data-course="${p.course}"><td>${p.course}</td><td>${p.name}</td><td>${ctrl}</td></tr>`;
                });
                html += `</tbody></table>`;
            }

            html += `</div></div>`;
        });

        outputDisciplines.innerHTML = html;
    }

    function showDepartmentPreview(deptCode, deptName) {
        const deptData = { disciplines: [], practices: [] };

        for (const course in coursesData) {
            coursesData[course].forEach(item => {
                const code = item.departmentCode || 'Без кафедры';
                if (code === deptCode) {
                    const itemWithCourse = { ...item, course };
                    if (item.isPractice) {
                        deptData.practices.push(itemWithCourse);
                    } else {
                        deptData.disciplines.push(itemWithCourse);
                    }
                }
            });
        }

        if (deptData.disciplines.length === 0 && deptData.practices.length === 0) {
            showErrorDisciplines('Нет данных для предпросмотра');
            return;
        }

        // Фильтр по курсу
        const courseFilter = deptCourseFilter[deptCode] || 'all';
        let disc = deptData.disciplines;
        let pract = deptData.practices;
        if (courseFilter !== 'all') {
            disc = disc.filter(d => d.course === courseFilter);
            pract = pract.filter(p => p.course === courseFilter);
        }
        if (disc.length === 0 && pract.length === 0) {
            showErrorDisciplines('Нет данных для выбранного курса');
            return;
        }

        let html = `<h3>Кафедра ${deptCode}: ${deptName}</h3>`;

        if (disc.length > 0) {
            html += `<h4>Дисциплины</h4>`;
            html += `<table class="preview-table">
                <thead>
                    <tr>
                        <th>Наименование дисциплин/практик,
                        руководство курсовыми/НИР/ВКР, участие в ГИА и др.</th>
                        <th>Код и наименование направления,
                        направленность,
                        форма обучения
                        (очная/очно-заочная/заочная)</th>
                        <th>Курс</th>
                        <th>Кол-во студ.</th>
                        <th>Групп/подгр.</th>
                        <th>Осень Лекц.</th>
                        <th>Осень Сем., Практ.</th>
                        <th>Осень Лаб.</th>
                        <th>Осень Форма пром. аттест.</th>
                        <th>Весна Лекц.</th>
                        <th>Весна Сем., Практ.</th>
                        <th>Весна Лаб.</th>
                        <th>Весна Форма пром. аттест.</th>
                        <th>Предэкз. конс.</th>
                        <th>Экзамен, зачет, К</th>
                        <th>Итого</th>
                    </tr>
                </thead>
                <tbody>`;

        disc.forEach(d => {
            // Приведение к числу, если возможно
            const toNum = val => { const n = parseFloat(val); return isNaN(n) ? 0 : n; };
            const totalHours = toNum(d.lecAutumn) + toNum(d.labAutumn) + toNum(d.pracAutumn) +
                               toNum(d.lecSpring) + toNum(d.labSpring) + toNum(d.pracSpring);
            const totalStr = totalHours > 0 ? totalHours.toString().replace('.', ',') : '';

            html += `<tr>
                <td>${d.name}</td>
                <td></td>
                <td>${d.course}</td>
                <td></td>
                <td></td>
                <td>${d.lecAutumn}</td>
                <td>${d.pracAutumn}</td>
                <td>${d.labAutumn}</td>
                <td>${d.controlAutumn || ''}</td>
                <td>${d.lecSpring}</td>
                <td>${d.pracSpring}</td>
                <td>${d.labSpring}</td>
                <td>${d.controlSpring || ''}</td>
                <td></td>
                <td></td>
                <td></td>
            </tr>`;
        });
        html += `</tbody></table>`;
    }

    if (pract.length > 0) {
        html += `<h4>Практики</h4>`;
        html += `<table class="preview-table">
            <thead>
                <tr>
                    <th>Курс</th>
                    <th>Наименование</th>
                    <th>Контроль</th>
                </tr>
            </thead>
            <tbody>`;
        pract.forEach(p => {
            const ctrl = p.controlAutumn || p.controlSpring || 'ЗаО';
            html += `<tr><td>${p.course}</td><td>${p.name}</td><td>${ctrl}</td></tr>`;
        });
        html += `</tbody></table>`;
    }

    document.getElementById('previewTitle').textContent = `Предпросмотр: Кафедра ${deptCode}`;
    document.getElementById('previewContent').innerHTML = html;
    document.getElementById('previewModal').style.display = 'flex';
    }
    
        // ---------------------------- Модальное окно выбора файла ----------------------------
    function showFileSelectModal(deptCode, deptName) {
        currentDeptForAppend = { code: deptCode, name: deptName, data: null };

        // Собираем и фильтруем данные кафедры
        const courseFilter = deptCourseFilter[deptCode] || 'all';
        const deptData = { disciplines: [], practices: [] };
        for (const course in coursesData) {
            coursesData[course].forEach(item => {
                const code = item.departmentCode || 'Без кафедры';
                if (code === deptCode) {
                    const itemWithCourse = { ...item, course };
                    if (item.isPractice) {
                        deptData.practices.push(itemWithCourse);
                    } else {
                        deptData.disciplines.push(itemWithCourse);
                    }
                }
            });
        }
        let disc = deptData.disciplines;
        let pract = deptData.practices;
        if (courseFilter !== 'all') {
            disc = disc.filter(d => d.course === courseFilter);
            pract = pract.filter(p => p.course === courseFilter);
        }
        currentDeptForAppend = {
            code: deptCode,
            name: deptName,
            data: { disciplines: disc, practices: pract },
            course: courseFilter !== 'all' ? courseFilter : null
        };

        // Загружаем список файлов пользователя
        const listContainer = document.getElementById('fileSelectList');
        listContainer.innerHTML = '<div class="loading">Загрузка списка файлов...</div>';
        document.getElementById('fileSelectModal').style.display = 'flex';

        fetch('/api/my-files')
            .then(response => response.json())
            .then(files => {
                if (files.length === 0) {
                    listContainer.innerHTML = '<div class="no-data">Нет доступных файлов</div>';
                    return;
                }
                let html = '<div style="display: grid; gap: 10px;">';
                files.forEach(file => {
                    const date = new Date(file.created_at).toLocaleString('ru-RU');
                    html += `
                        <div style="padding: 10px; border: 1px solid #ddd; border-radius: 5px; cursor: pointer; transition: background 0.3s;"
                             onclick="appendToSelectedFile(${file.id})"
                             onmouseover="this.style.background='#f0f7ff'"
                             onmouseout="this.style.background='white'">
                            <strong>${file.filename}</strong><br>
                            <small>${file.file_type === 'workload' ? 'Нагрузка' : 'Кафедры'} | ${date} | ${(file.file_size/1024).toFixed(2)} КБ</small>
                        </div>`;
                });
                html += '</div>';
                listContainer.innerHTML = html;
            })
            .catch(error => {
                listContainer.innerHTML = `<div class="error">Ошибка загрузки: ${error.message}</div>`;
            });
    }

    window.showFileSelectModal = showFileSelectModal;

    async function appendToSelectedFile(fileId) {
        if (!currentDeptForAppend || !currentDeptForAppend.data) {
            alert('Нет данных для добавления');
            return;
        }

        const modal = document.getElementById('fileSelectModal');
        const listContainer = document.getElementById('fileSelectList');
        listContainer.innerHTML = '<div class="loading">Добавление данных...</div>';

        try {
            const response = await fetch(`/api/append-to-file/${fileId}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    departmentData: currentDeptForAppend.data,
                    course: currentDeptForAppend.course
                })
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || 'Ошибка при добавлении');
            }

            const result = await response.json();
            alert(result.message || 'Данные успешно добавлены');
            modal.style.display = 'none';
        } catch (error) {
            listContainer.innerHTML = `<div class="error">Ошибка: ${error.message}</div>`;
        }
    }

    window.appendToSelectedFile = appendToSelectedFile;

    async function exportDepartmentExcel(deptCode, deptName) {
        const deptData = { disciplines: [], practices: [] };
        
        for (const course in coursesData) {
            coursesData[course].forEach(item => {
                const code = item.departmentCode || 'Без кафедры';
                
                if (code === deptCode) {
                    const itemWithCourse = { ...item, course };
                    if (item.isPractice) {
                        deptData.practices.push(itemWithCourse);
                    } else {
                        deptData.disciplines.push(itemWithCourse);
                    }
                }
            });
        }
        
        // Фильтр по курсу
        const courseFilter = deptCourseFilter[deptCode] || 'all';
        let filteredDisciplines = deptData.disciplines;
        let filteredPractices = deptData.practices;
        if (courseFilter !== 'all') {
            filteredDisciplines = deptData.disciplines.filter(d => d.course === courseFilter);
            filteredPractices = deptData.practices.filter(p => p.course === courseFilter);
        }
        if (filteredDisciplines.length === 0 && filteredPractices.length === 0) {
            showErrorDisciplines('Нет данных для выбранного курса');
            return;
        }

        const buttons = document.querySelectorAll('.btn-export-dept');
        let targetButton = null;
        
        buttons.forEach(btn => {
            const deptSection = btn.closest('.department-section');
            if (deptSection) {
                const deptCodeElement = deptSection.querySelector('.department-code');
                if (deptCodeElement && deptCodeElement.textContent.includes(deptCode)) {
                    targetButton = btn;
                    btn.textContent = '⏳';
                    btn.disabled = true;
                }
            }
        });

        try {
            const response = await fetch(`${SERVER_URL}/generate-department-excel`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    departmentData: { disciplines: filteredDisciplines, practices: filteredPractices },
                    departmentName: deptName,
                    departmentCode: deptCode,
                    course: courseFilter !== 'all' ? courseFilter : null
                })
            });
            
            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || 'Ошибка при генерации файла');
            }
            
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            
            const safeName = deptName.replace(/[^a-zA-Zа-яА-Я0-9]/g, '_');
            a.download = `Кафедра_${deptCode}_${safeName}.xlsx`;
            
            document.body.appendChild(a);
            a.click();
            
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
        } catch (error) {
            showErrorDisciplines(`Ошибка при экспорте: ${error.message}`);
            console.error('Export error:', error);
        } finally {
            if (targetButton) {
                targetButton.textContent = '📥 Excel';
                targetButton.disabled = false;
            }
        }
    }

    function showLoadingDisciplines(show) {
        loadingDisciplines.style.display = show ? 'block' : 'none';
        if (show) {
            loadingDisciplines.innerHTML = '<p>⏳ Обработка данных...</p>';
        }
    }

    function showErrorDisciplines(msg) {
        errorDisciplines.textContent = msg;
        errorDisciplines.style.display = 'block';
        setTimeout(() => {
            errorDisciplines.style.display = 'none';
        }, 5000);
    }

    window.toggleCourse = function(header) {
        const content = header.nextElementSibling;
        content.style.display = content.style.display === 'block' ? 'none' : 'block';
    };

    window.toggleDepartment = function(header) {
        const content = header.nextElementSibling;
        content.style.display = content.style.display === 'block' ? 'none' : 'block';
    };

    window.setDeptCourse = function(deptCode, course) {
        deptCourseFilter[deptCode] = course;

        const section = document.querySelector(`.department-section[data-dept-code="${deptCode.replace(/"/g, '\\"')}"]`);
        if (!section) return;
        const btns = section.querySelectorAll('.dept-course-btn');
        btns.forEach(btn => {
            const btnCourse = btn.getAttribute('data-course');
            btn.classList.toggle('active', btnCourse === course);
        });
        const rows = section.querySelectorAll('tr[data-course]');
        rows.forEach(row => {
            const rowCourse = row.getAttribute('data-course');
            if (course === 'all' || rowCourse === course) {
                row.style.display = '';
            } else {
                row.style.display = 'none';
            }
        });
    };
        
    window.exportDepartmentExcel = exportDepartmentExcel;
    window.showDepartmentPreview = showDepartmentPreview;

})();

// ==================== ПРИЛОЖЕНИЕ 3: МОИ ФАЙЛЫ ====================
function loadMyFiles() {
    const container = document.getElementById('filesList');
    container.innerHTML = '<div class="loading">Загрузка списка файлов...</div>';

    fetch('/api/my-files')
        .then(response => {
            if (!response.ok) throw new Error('Ошибка загрузки');
            return response.json();
        })
        .then(files => {
            if (files.length === 0) {
                container.innerHTML = '<div class="no-data">У вас пока нет сохранённых файлов</div>';
                return;
            }

            let html = `
                <table class="discipline-table">
                    <thead>
                        <tr>
                            <th>Имя файла</th>
                            <th>Тип</th>
                            <th>Дата создания</th>
                            <th>Размер</th>
                            <th>Действия</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            files.forEach(file => {
                const typeText = file.file_type === 'workload' ? 'Преподавательская нагрузка' : 'Распределение по кафедрам';
                const date = new Date(file.created_at).toLocaleString('ru-RU');
                const size = file.file_size ? (file.file_size / 1024).toFixed(2) + ' КБ' : '—';

                html += `
                    <tr>
                        <td>${file.filename}</td>
                        <td>${typeText}</td>
                        <td>${date}</td>
                        <td>${size}</td>
                        <td>
                            <button class="btn btn-small btn-success" onclick="downloadUserFile(${file.id})">Скачать</button>
                            <button class="btn btn-small btn-danger" onclick="deleteUserFile(${file.id})">Удалить</button>
                        </td>
                    </tr>
                `;
            });

            html += '</tbody></table>';
            container.innerHTML = html;
        })
        .catch(error => {
            container.innerHTML = `<div class="error">Ошибка: ${error.message}</div>`;
        });
}

window.downloadUserFile = function(fileId) {
    window.location.href = `/api/download-file/${fileId}`;
};

window.deleteUserFile = function(fileId) {
    if (!confirm('Удалить файл?')) return;

    fetch(`/api/delete-file/${fileId}`, { method: 'DELETE' })
        .then(response => {
            if (response.ok) {
                loadMyFiles(); // Обновить список
            } else {
                alert('Ошибка при удалении файла');
            }
        })
        .catch(error => {
            alert('Ошибка сети: ' + error.message);
        });
};

// ==================== АВТОРИЗАЦИЯ: ВЫХОД И ИМЯ ПОЛЬЗОВАТЕЛЯ ====================
document.getElementById('logoutBtn')?.addEventListener('click', async () => {
    await fetch('/logout');
    window.location.href = '/login';
});

fetch('/api/check-auth')
    .then(res => res.json())
    .then(data => {
        if (data.authenticated) {
            document.getElementById('currentUser').textContent = `👤 ${data.username}`;
        }
    });

// Если вкладка "Мои файлы" активна при загрузке (редко), загружаем список
if (document.getElementById('myfiles-app')?.classList.contains('active')) {
    loadMyFiles();
}
