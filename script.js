document.getElementById('loadBtn').addEventListener('click', handleFileUpload);
document.getElementById('addTeacherBtn').addEventListener('click', addTeacher);
document.getElementById('cancelModal').addEventListener('click', hideModal);
document.getElementById('exportBtn').addEventListener('click', exportToExcel);

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
    cellFormats: {}
};

// Переменные для текущего выбранного направления/подгруппы
let currentAssignment = {
    directionId: null,
    isSubgroup: false,
    subjectName: '',
    directionCode: '',
    facultyName: '',
    directionData: null
};

function handleFileUpload() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        showError('Пожалуйста, выберите файл');
        return;
    }

    showLoading(true);
    currentData.faculties = {};
    currentData.assignments = [];
    updateAssignmentsDisplay();

    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {
                type: 'array',
                cellStyles: true,
                cellNF: true // Получаем числовые форматы
            });

            // Сохраняем оригинальную книгу
            currentData.originalWorkbook = workbook;

            // Получаем данные из первого листа
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Сохраняем оригинальные данные листа
            currentData.originalSheetData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                raw: false,
                defval: ''
            });

            // Извлекаем ВСЕ стили форматирования
            extractAllCellStyles(worksheet);

            // Сохраняем заголовочные строки (первые 4 строки)
            currentData.headerRows = currentData.originalSheetData.slice(0, 4);

            // Преобразуем в JSON для обработки
            const jsonData = currentData.originalSheetData;

            // Обрабатываем данные
            processData(jsonData);

            // Показываем опции экспорта
            document.getElementById('exportOptions').style.display = 'block';

        } catch (error) {
            showError('Ошибка при чтении файла: ' + error.message);
            console.error(error);
        }
    };

    reader.onerror = function () {
        showError('Ошибка при чтении файла');
    };

    reader.readAsArrayBuffer(file);
}

function extractAllCellStyles(worksheet) {
    // Извлекаем ВСЕ стили ячеек из рабочего листа
    currentData.cellStyles = {};
    currentData.cellFormats = {};

    // Получаем диапазон ячеек
    const range = XLSX.utils.decode_range(worksheet['!ref']);

    // Проходим по всем ячейкам и сохраняем их стили и форматы
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cellAddress];

            if (cell) {
                // Сохраняем стиль ячейки
                if (cell.s) {
                    currentData.cellStyles[cellAddress] = JSON.parse(JSON.stringify(cell.s));
                }

                // Сохраняем числовой формат
                if (cell.z) {
                    currentData.cellFormats[cellAddress] = cell.z;
                }

                // Сохраняем тип данных
                if (cell.t) {
                    if (!currentData.cellStyles[cellAddress]) {
                        currentData.cellStyles[cellAddress] = {};
                    }
                    currentData.cellStyles[cellAddress].t = cell.t;
                }
            }
        }
    }

    // Сохраняем ширины столбцов
    if (worksheet['!cols']) {
        currentData.columnWidths = JSON.parse(JSON.stringify(worksheet['!cols']));
    }

    // Сохраняем высоты строк
    if (worksheet['!rows']) {
        currentData.rowHeights = JSON.parse(JSON.stringify(worksheet['!rows']));
    }

    // Сохраняем объединенные ячейки
    if (worksheet['!merges']) {
        currentData.mergedCells = JSON.parse(JSON.stringify(worksheet['!merges']));
    }

    // Сохраняем диапазон ячеек
    currentData.range = worksheet['!ref'];
}

function processData(data) {
    // Пропускаем первые 4 строки (заголовки)
    const rows = data.slice(4);

    const faculties = {};
    let currentFaculty = '';
    let currentSubject = null;
    let directionIndex = 0;

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];

        // Проверяем, является ли строка названием факультета
        if (row[3] && !row[4]) {
            currentFaculty = row[3];
            currentSubject = null;
            if (!faculties[currentFaculty]) {
                faculties[currentFaculty] = [];
            }
        }
        // Проверяем, является ли строка новым предметом
        else if (row[3] && row[4] && currentFaculty) {
            currentSubject = {
                name: row[3],
                directions: [],
                rowIndex: i + 4 // Сохраняем индекс строки в исходных данных
            };
            faculties[currentFaculty].push(currentSubject);

            // Добавляем первое направление с уникальным ID
            const direction = addDirectionToSubject(currentSubject, row, directionIndex++, currentFaculty, i + 4);

            // Проверяем подгруппу
            if (i + 1 < rows.length) {
                const nextRow = rows[i + 1];
                const hasNoDirectionCode = !nextRow[4] || String(nextRow[4]).trim() === '';
                const hasSeminars = (nextRow[9] && nextRow[9] > 0) || (nextRow[17] && nextRow[17] > 0);

                if (hasNoDirectionCode && hasSeminars) {
                    direction.hasSubgroup = true;
                    direction.subgroupData = {
                        groups: nextRow[7] || direction.groups,
                        autumn: {
                            lectures: nextRow[8] || 0,
                            seminars: nextRow[9] || 0,
                            labs: nextRow[10] || 0,
                            attestation: nextRow[15] || direction.autumn.attestation
                        },
                        spring: {
                            lectures: nextRow[16] || 0,
                            seminars: nextRow[17] || 0,
                            labs: nextRow[18] || 0,
                            attestation: nextRow[23] || direction.spring.attestation
                        },
                        total: nextRow[34] || direction.total,
                        preExamConsultation: nextRow[32] || '',
                        examOrTest: nextRow[33] || '',
                        rowIndex: i + 5 // Индекс строки подгруппы
                    };
                    i++;
                }
            }
        }
        // Дополнительное направление для текущего предмета
        else if (!row[3] && row[4] && currentFaculty && currentSubject) {
            const direction = addDirectionToSubject(currentSubject, row, directionIndex++, currentFaculty, i + 4);

            if (i + 1 < rows.length) {
                const nextRow = rows[i + 1];
                const hasNoDirectionCode = !nextRow[4] || String(nextRow[4]).trim() === '';
                const hasSeminars = (nextRow[9] && nextRow[9] > 0) || (nextRow[17] && nextRow[17] > 0);

                if (hasNoDirectionCode && hasSeminars) {
                    direction.hasSubgroup = true;
                    direction.subgroupData = {
                        groups: nextRow[7] || direction.groups,
                        autumn: {
                            lectures: nextRow[8] || 0,
                            seminars: nextRow[9] || 0,
                            labs: nextRow[10] || 0,
                            attestation: nextRow[15] || direction.autumn.attestation
                        },
                        spring: {
                            lectures: nextRow[16] || 0,
                            seminars: nextRow[17] || 0,
                            labs: nextRow[18] || 0,
                            attestation: nextRow[23] || direction.spring.attestation
                        },
                        total: nextRow[34] || direction.total,
                        preExamConsultation: nextRow[32] || '',
                        examOrTest: nextRow[33] || '',
                        rowIndex: i + 5 // Индекс строки подгруппы
                    };
                    i++;
                }
            }
        }
    }

    currentData.faculties = faculties;
    displayData(faculties);
}

function addDirectionToSubject(subject, row, id, faculty, rowIndex) {
    const direction = {
        id: id,
        faculty: faculty,
        subjectName: subject.name,
        code: row[4],
        course: row[5],
        studentsCount: row[6],
        groups: row[7],
        autumn: {
            lectures: row[8] || 0,
            seminars: row[9] || 0,
            labs: row[10] || 0,
            attestation: row[15] || ''
        },
        spring: {
            lectures: row[16] || 0,
            seminars: row[17] || 0,
            labs: row[18] || 0,
            attestation: row[23] || ''
        },
        total: row[34] || 0,
        preExamConsultation: row[32] || '',
        examOrTest: row[33] || '',
        hasSubgroup: false,
        subgroupData: null,
        rowIndex: rowIndex,
        originalRowData: row // Сохраняем оригинальные данные строки
    };

    subject.directions.push(direction);
    return direction;
}

function exportToExcel() {
    if (!currentData.originalSheetData) {
        showError('Сначала загрузите файл');
        return;
    }

    if (currentData.assignments.length === 0) {
        showError('Нет прикрепленных преподавателей для экспорта');
        return;
    }

    // Получаем имя файла от пользователя
    const fileNameInput = document.getElementById('exportFileName');
    let fileName = fileNameInput.value.trim();

    if (!fileName) {
        fileName = 'Нагрузка_преподавателей';
    }

    // Убедимся, что имя файла заканчивается на .xlsx
    if (!fileName.toLowerCase().endsWith('.xlsx')) {
        fileName += '.xlsx';
    }

    showLoading(true);

    try {
        // Создаем новую книгу Excel
        const wb = XLSX.utils.book_new();

        // Группируем прикрепления по преподавателям
        const assignmentsByTeacher = {};

        currentData.assignments.forEach(assignment => {
            if (!assignmentsByTeacher[assignment.teacherId]) {
                assignmentsByTeacher[assignment.teacherId] = [];
            }
            assignmentsByTeacher[assignment.teacherId].push(assignment);
        });

        // Для каждого преподавателя создаем отдельный лист
        Object.keys(assignmentsByTeacher).forEach(teacherId => {
            const teacher = currentData.teachers.find(t => t.id == teacherId);
            if (!teacher) return;

            const teacherAssignments = assignmentsByTeacher[teacherId];

            // Создаем лист для преподавателя
            const ws = createTeacherSheet(teacher, teacherAssignments);

            // Очищаем имя преподавателя для имени листа
            let sheetName = cleanSheetName(teacher.name);

            // Добавляем лист в книгу
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
        });

        // Скачиваем файл с пользовательским именем
        XLSX.writeFile(wb, fileName);

        showLoading(false);
        alert('Файл успешно экспортирован! В файле созданы отдельные листы для каждого преподавателя.');

    } catch (error) {
        showLoading(false);
        showError('Ошибка при экспорте: ' + error.message);
        console.error(error);
    }
}

function createTeacherSheet(teacher, assignments) {
    // Создаем новый рабочий лист
    const ws = XLSX.utils.aoa_to_sheet([]);

    // Восстанавливаем ширины столбцов
    if (currentData.columnWidths) {
        ws['!cols'] = currentData.columnWidths;
    }

    // Восстанавливаем высоты строк
    if (currentData.rowHeights) {
        ws['!rows'] = currentData.rowHeights;
    }

    // Восстанавливаем объединенные ячейки
    if (currentData.mergedCells) {
        ws['!merges'] = currentData.mergedCells;
    }

    let exportData = [];
    let currentRow = 0;

    // Копируем заголовочные строки
    for (let i = 0; i < currentData.headerRows.length; i++) {
        exportData.push([...currentData.headerRows[i]]);
        currentRow++;
    }

    // Экспорт направлений для конкретного преподавателя
    exportTeacherAssignments(ws, exportData, currentRow, assignments);

    // Записываем данные в рабочий лист
    XLSX.utils.sheet_add_aoa(ws, exportData, { origin: 'A1' });

    // Применяем сохраненные стили ко всем ячейкам
    applySavedStylesToSheet(ws, exportData.length);

    return ws;
}

function exportTeacherAssignments(ws, exportData, startRow, assignments) {
    // Группируем прикрепления по направлению для устранения дубликатов
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

    // Собираем все исходные строки, которые нужно экспортировать
    const rowsToExport = [];

    for (const [key, assignmentList] of uniqueAssignments) {
        const assignment = assignmentList[0];
        const directionData = getDirectionData(assignment.directionId);

        if (!directionData) continue;

        // Определяем, какую строку брать (основную или подгруппу)
        let sourceRowIndex = assignment.isSubgroup && directionData.hasSubgroup && directionData.subgroupData
            ? directionData.subgroupData.rowIndex
            : directionData.rowIndex;

        // Копируем строку из оригинала
        const originalRow = currentData.originalSheetData[sourceRowIndex];
        if (!originalRow) continue;

        // Добавляем информацию о строке для дальнейшей обработки
        rowsToExport.push({
            sourceRowIndex: sourceRowIndex,
            originalRow: originalRow,
            directionData: directionData,
            isSubgroup: assignment.isSubgroup
        });
    }

    // Сортируем строки по их оригинальному индексу (чтобы сохранить порядок)
    rowsToExport.sort((a, b) => a.sourceRowIndex - b.sourceRowIndex);

    // Теперь обрабатываем строки, сохраняя оригинальные данные
    for (let i = 0; i < rowsToExport.length; i++) {
        const rowData = rowsToExport[i];
        const newRow = [...rowData.originalRow];

        // НЕ изменяем название дисциплины (колонка 3, индекс 2)
        // Оставляем оригинальное значение из файла

        // Обрезаем массив до 35 колонок (по количеству колонок в таблице)
        if (newRow.length > 35) {
            newRow.splice(35, newRow.length - 35);
        } else if (newRow.length < 35) {
            // Дополняем до 35 колонок пустыми значениями
            while (newRow.length < 35) {
                newRow.push('');
            }
        }

        // Добавляем строку в данные для экспорта
        exportData.push(newRow);
        rowsAdded++;

        // Копируем стили для этой строки из оригинала
        copyRowStyles(ws, rowData.sourceRowIndex, rowIndex, newRow.length);

        rowIndex++;
    }

    console.log(`Экспортировано ${rowsAdded} строк с сохранением оригинальных данных`);
    return rowsAdded;
}

function cleanSheetName(name) {
    // Очищаем имя для использования в качестве имени листа Excel
    // Excel имеет ограничения: не более 31 символа, нельзя использовать: \ / ? * [ ] :
    let cleaned = name.replace(/[\\/*?[\]:]/g, '');
    cleaned = cleaned.substring(0, 31);

    // Если имя пустое после очистки, используем "Лист"
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

function copyRowStyles(ws, sourceRowIndex, targetRowIndex, colCount) {
    // Копируем стили для каждой ячейки в строке
    for (let col = 0; col < colCount; col++) {
        const sourceCellAddress = XLSX.utils.encode_cell({ r: sourceRowIndex, c: col });
        const targetCellAddress = XLSX.utils.encode_cell({ r: targetRowIndex, c: col });

        // Копируем стиль
        if (currentData.cellStyles[sourceCellAddress]) {
            if (!ws[targetCellAddress]) {
                ws[targetCellAddress] = { v: '' };
            }
            ws[targetCellAddress].s = JSON.parse(JSON.stringify(currentData.cellStyles[sourceCellAddress]));
        }

        // Копируем числовой формат
        if (currentData.cellFormats[sourceCellAddress]) {
            if (!ws[targetCellAddress]) {
                ws[targetCellAddress] = { v: '' };
            }
            ws[targetCellAddress].z = currentData.cellFormats[sourceCellAddress];
        }
    }
}

function applySavedStylesToSheet(ws, totalRows) {
    // Применяем стили заголовочных строк
    for (let R = 0; R < 4 && R < totalRows; R++) {
        for (let C = 0; C < 36; C++) {
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
        showError('Введите ФИО преподавателя');
        return;
    }

    // Проверяем, нет ли уже такого преподавателя
    if (currentData.teachers.some(t => t.name === teacherName)) {
        showError('Преподаватель с таким ФИО уже добавлен');
        return;
    }

    const teacher = {
        id: Date.now(),
        name: teacherName
    };

    currentData.teachers.push(teacher);
    teacherNameInput.value = '';
    updateTeachersDisplay();
}

function assignTeacherToCurrent(teacherId) {
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

function removeAssignment(assignmentId) {
    currentData.assignments = currentData.assignments.filter(a => a.id !== assignmentId);
    updateTeachersDisplay();
    updateDirectionDisplays();
}

function updateTeachersDisplay() {
    const teacherList = document.getElementById('teacherList');

    if (currentData.teachers.length === 0) {
        teacherList.innerHTML = '<div class="no-data">Пока нет добавленных преподавателей</div>';
    } else {
        teacherList.innerHTML = '<div class="teacher-columns">';

        currentData.teachers.forEach(teacher => {
            const directionCount = currentData.assignments.filter(a => a.teacherId === teacher.id && !a.isSubgroup).length;
            const subgroupCount = currentData.assignments.filter(a => a.teacherId === teacher.id && a.isSubgroup).length;

            const teacherItem = document.createElement('div');
            teacherItem.className = 'teacher-item';
            teacherItem.innerHTML = `
                <div class="teacher-name">${teacher.name}</div>
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
        showError('Сначала добавьте преподавателей');
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
        showError('Направление не найдено');
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

function displayData(faculties) {
    const output = document.getElementById('output');
    output.innerHTML = '';

    if (Object.keys(faculties).length === 0) {
        showError('Не удалось найти данные о факультетах и дисциплинах');
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
            <p>Общее количество часов: ${totalHours}</p>
        </div>
    `;

    output.innerHTML = statsHTML;

    for (const faculty in faculties) {
        const facultySection = document.createElement('div');
        facultySection.className = 'faculty-section';

        let facultySubjectCount = 0;
        let facultyDirectionCount = 0;
        let facultySubgroupCount = 0;

        faculties[faculty].forEach(subject => {
            facultySubjectCount++;
            facultyDirectionCount += subject.directions.length;
            subject.directions.forEach(direction => {
                if (direction.hasSubgroup) {
                    facultySubgroupCount++;
                }
            });
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

                const directionAssignments = currentData.assignments.filter(a => a.directionId == direction.id);
                const mainAssignments = directionAssignments.filter(a => !a.isSubgroup);
                const subgroupAssignments = directionAssignments.filter(a => a.isSubgroup);

                subjectHTML += `
                    <div class="${directionClass}" data-id="${direction.id}">
                        <button class="assign-btn" onclick="showAssignModal('${direction.id}', false)">
                            Прикрепить направление
                        </button>
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
                                <span class="detail-label">Итого часов:</span> ${direction.total}
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
                    subjectHTML += `
                        <div class="subgroup-info">
                            <div class="subgroup-info-text">
                                <strong>Дисциплина:</strong> ${direction.subjectName}<br>
                                <strong>Направление:</strong> ${direction.code}
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
                                    <span class="detail-label">Итого часов (подгруппа):</span> ${direction.subgroupData.total}
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

    showLoading(false);
}

function showLoading(show) {
    document.getElementById('loading').style.display = show ? 'block' : 'none';
}

function showError(message) {
    const errorDiv = document.getElementById('error');
    errorDiv.textContent = message;
    errorDiv.style.display = 'block';
    showLoading(false);
    setTimeout(() => {
        errorDiv.style.display = 'none';
    }, 5000);
}

window.addEventListener('click', function (event) {
    const modal = document.getElementById('teacherModal');
    if (event.target === modal) {
        hideModal();
    }
});