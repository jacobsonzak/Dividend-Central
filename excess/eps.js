import * as XLSX from 'xlsx';

function processExcelFile(sheet) {
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, cellDates: true });

    let result = {};
    json.forEach((row, index) => {
        result[index + 1] = row.map(cell => {
            if (cell instanceof Date) {
                return cell.toISOString().split('T')[0];
            } else if (typeof cell === 'number') {
                return cell;
            }
            return cell;
        });
    });
    
    return result;
}

function loadData(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const workbook = XLSX.read(e.target.result, { type: 'binary' });
                let data = {};
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const point = processExcelFile(worksheet);
                    reworkData(point);
                    data[sheetName] = point;
                });
                resolve(data);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = (error) => reject(error);
        
        reader.readAsArrayBuffer(file);
    });
}

function reworkData(data) {
    delete data['1'];
    for (let val in data) {
        data[val][1] = Number(data[val][1]);
    }
}

export { loadData, processExcelFile, reworkData };




import * as XLSX from 'xlsx';

function processExcelFile(sheet) {

    // Convert sheet to a 2D array
    //const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, cellDates: true });

    // Create the result object
    let result = {};
    json.forEach((row, index) => {
        result[index + 1] = row.map(cell => {
            if (cell instanceof Date) {
                // Format date as YYYY-MM-DD
                return cell.toISOString().split('T')[0];
            } else if (typeof cell === 'number') {
                // Ensure the value is treated as a number
                return cell;
            }
            return cell; // Leave other types as they are
        });
    });
    
    return result;
}

function loadData(filePath){
    // Read the file as a binary string
    const workbook = XLSX.readFile(filePath);
    var len = workbook.SheetNames.length;
    let data = {};
    for(let i = 0; i < len; i++){
        const sheetName = workbook.SheetNames[i];
        const worksheet = workbook.Sheets[sheetName];
        const point = processExcelFile(worksheet);
        reworkData(point);
        data[sheetName] = point;
    }
    return data;
}

function reworkData(data){
    delete data['1'];
    for(val in data){
        data[val][1] = Number(data[val][1]);
    }
}



export { loadData, processExcelFile, reworkData };