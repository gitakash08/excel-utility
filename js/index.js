let selectedFile;
let modifiedData;
let isProcessing;
let fileNameDisplay;
let downloadName;

let userColumns = ['DIV/RLY', 'SECTION', 'STATION', 'LINE', 'XOVER', 'LRRR',
    'FROM KM', 'FROM MET', 'TO KM', 'TO MET', 'LENGTH(km)', 'SECCODE', 'LINECODE',
    'STNCODE', 'LOOPLINECODE', 'XOVERID', 'ROLLING MONTH', 'LAYING MONTH', 'SUPPLIER',
    'ACCUMULATED GMT', 'RAIL SECTION', 'GRADE OF STEEL', 'GMT CARRIED AT LAYING'];

handleFile = () => {
    let fileInput = document.getElementById('fileInput');
    selectedFile = fileInput.files[0];
    downloadName = selectedFile.name.split('.')[0]
    displayFileName(selectedFile.name);
    processingFun();
}
displayFileName = (fileName) => {
    fileNameDisplay = document.getElementById('fileNameDisplay');
    fileNameDisplay.textContent = `Selected File: ${fileName}`;
}
mapUserColumnsToExcelColumns = (userColumns_, headerRow) => {
    if (userColumns_.length !== headerRow.length) {
        alert("Extra Column in Excel sheet...");
        return false;
    }
    let columnMapping = {};
    for (let i = 0; i < userColumns_.length; ++i) {
        let userColumn = userColumns_[i];
        let excelColumnIndex = headerRow.indexOf(userColumn);
        if (excelColumnIndex === -1) {
            alert(`${userColumn} column is not present in the Excel Sheet 
                   and Wrong Column name is ${headerRow[i]}`);
            return false; // Terminate if any user column is not found in Excel
        }
        columnMapping[userColumn] = excelColumnIndex;
    }
    return columnMapping;
}
readExcelFile = (file) => {
    let reader = new FileReader();
    reader.onload = function (e) {
        let data = e.target.result;
        let workbook = XLSX.read(data, { type: 'binary' });
        let sheetName = workbook.SheetNames[0];
        let sheet = workbook.Sheets[sheetName];

        // Skip the first four rows
        let range = XLSX.utils.decode_range(sheet['!ref']);
        range.s.r += 4; // Skip 4 rows
        sheet['!ref'] = XLSX.utils.encode_range(range);

        let headerData = XLSX.utils.sheet_to_json(sheet, { range: 1, header: 1 })[3];
        let headerRow = [];
        for (let C = range.s.c; C <= range.e.c; ++C) {
            if (headerData[C] !== '#') {
                let columnName = headerData[C] || '';
                headerRow.push(columnName);
            }
        }
        let columnMapping = mapUserColumnsToExcelColumns(userColumns, headerRow);

        if (!columnMapping) {
            alert('Columns not matched.Try again with correct excel file.');
            return;
        }

        // Create a new column at the beginning with heading "UID"
        let uidColumnIndex = range.s.c; // Index for the new "UID" column
        range.e.c++; // Increase the ending column index for existing columns

        // Add the "UID" column heading to the first row
        let uidHeadingCell = XLSX.utils.encode_cell({ r: range.s.r, c: uidColumnIndex });
        sheet[uidHeadingCell] = { t: 's', v: 'UID' };

        // Add auto-incremented values in the "UID" column
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            let cell = XLSX.utils.encode_cell({ r: R, c: uidColumnIndex });
            sheet[cell] = { t: 'n', v: R - range.s.r }; // Auto-incremented value
        }

        // Create an array to store rows to delete
        let rowsToDelete = [];
        // Loop through the rows and mark those with longer values in the "LINE" column
        for (let R = range.s.r; R <= range.e.r; ++R) {
            if (R === range.s.r) {
                // Skip the first row (header row)
                continue;
            }

            let lineCell = XLSX.utils.encode_cell({ r: R, c: columnMapping['LINE'] });
            let lineValue = sheet[lineCell]?.v || ''; // Get the value in the "LINE" column

            // Check if the "LINE" value has more than two characters
            if (lineValue.length > 2) {
                rowsToDelete.push(R);
            }
        }

        // Delete the marked rows
        for (let i = rowsToDelete.length - 1; i >= 0; i--) {
            let currentRow = rowsToDelete[i];
            for (let C = range.s.c; C <= range.e.c; ++C) {
                let cell = XLSX.utils.encode_cell({ r: currentRow, c: C });
                delete sheet[cell];
            }
        }

        let newColumnNames = ['DIST_FROM', 'DIST_TO', 'DIST_M', 'LROUTE'];
        newColumnNames.forEach((columnName, index) => {
            let columnIndex = range.e.c + index;
            let cell = XLSX.utils.encode_cell({ r: range.s.r, c: columnIndex });
            sheet[cell] = { t: 's', v: columnName };
        });

        range.e.c += newColumnNames.length;
        sheet['!ref'] = XLSX.utils.encode_range(range);

        // Calculate values for Dist_From column
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            
            let fromKmCell = XLSX.utils.encode_cell({ r: R, c: columnMapping['FROM KM'] + 1 });
            let fromMetCell = XLSX.utils.encode_cell({ r: R, c: columnMapping['FROM MET'] + 1 });
            let fromKmValue = sheet[fromKmCell]?.v || '';
            let fromMetValue = sheet[fromMetCell]?.v || '';

            let toKmCell = XLSX.utils.encode_cell({ r: R, c: columnMapping['TO KM'] + 1 });
            let toMetCell = XLSX.utils.encode_cell({ r: R, c: columnMapping['TO MET'] + 1 });
            let toKmValue = sheet[toKmCell]?.v || '';
            let toMetValue = sheet[toMetCell]?.v || '';

            let secCodeCell = XLSX.utils.encode_cell({ r: R, c: columnMapping['SECCODE'] + 1 });
            let lineCodeCell = XLSX.utils.encode_cell({ r: R, c: columnMapping['LINECODE'] + 1 });

            let secCodeValue = sheet[secCodeCell]?.v || '';
            let lineCodeValue = sheet[lineCodeCell]?.v || '';

            let distFromCell = XLSX.utils.encode_cell({ r: R, c: 24 });
            let distFromValue = fromKmValue * 1000 + fromMetValue;

            let distToCell = XLSX.utils.encode_cell({ r: R, c: 25 });
            let distToValue = toKmValue * 1000 + toMetValue;

            let distMCell = XLSX.utils.encode_cell({ r: R, c: 26 });
            let distMValue = distToValue - distFromValue;

            let lrouteCell = XLSX.utils.encode_cell({ r: R, c: 27 });
            let lrouteValue = secCodeValue + '' + lineCodeValue;

            sheet[distFromCell] = { t: 'n', v: distFromValue };
            sheet[distToCell] = { t: 'n', v: distToValue };
            sheet[distMCell] = { t: 'n', v: distMValue };
            sheet[lrouteCell] = { t: 's', v: lrouteValue };

            let rowsToBeDeleted = [];
            // Check if the "Dist_M" value is negative
            if (distMValue < 0) {
                rowsToBeDeleted.push(R);
                //Checking if LROUTE having null
            } else if (lrouteValue == '') {
                rowsToBeDeleted.push(R);
            }
            for (let i = rowsToBeDeleted.length - 1; i >= 0; i--) {
                const currentRow = rowsToBeDeleted[i];
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell = XLSX.utils.encode_cell({ r: currentRow, c: C });
                    delete sheet[cell];
                }
            }
        }
        // Update the sheet reference range
        sheet['!ref'] = XLSX.utils.encode_range(range);
        modifiedData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        modifiedData = modifiedData.filter(d => d[0] != undefined)
    };
    reader.readAsBinaryString(file);
    return isProcessing = true;
}
processData = (modifiedData) => {
    if (selectedFile) {
        let updatedSheet = XLSX.utils.aoa_to_sheet(modifiedData);
        let csvContent = XLSX.utils.sheet_to_csv(updatedSheet); // Convert the sheet to CSV format 
        let blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        let link = document.createElement('a');
        let fileName = `${downloadName}_updated.csv`;
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        alert("Updated file saved successfully!");
    } else {
        alert("Please select a file first...");
    }
}
processingFun = () => {
    if (selectedFile) {
        isProcessing = readExcelFile(selectedFile);
        if (isProcessing) return alert(`Processing done!! now you can save CSV file...`);
    } else {
        alert("Please select a file first...");
    }
}
saveFile = () => {
    processData(modifiedData);
    document.getElementById('fileInput').value = '';
    fileNameDisplay.textContent = '';
    selectedFile = null;
    modifiedData = null;
}