document.getElementById('excelFile').addEventListener('change', handleFile, false);

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) {
        alert('Please select an Excel file.');
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        processFile(workbook);
    };
    reader.readAsArrayBuffer(file);
}

function processFile(workbook) {
    const sheetNames = workbook.SheetNames;
    let allData = {};

    sheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        allData[sheetName] = calculateRatesAndRatios(jsonData);
    });

    displayResults(allData);
}

function calculateRatesAndRatios(data) {
    let dividends = [];

    data.forEach((row, index) => {
        if (index > 0 && row[4] !== undefined && !isNaN(row[4])) {
            dividends.push(parseFloat(row[4]));
        }
    });

    let rateChanges = [];
    if (dividends.length === 0) {
        rateChanges.push('No valid dividend data found in column F.');
    } else {
        for (let i = 1; i < dividends.length; i++) {
            const rate = ((dividends[i] - dividends[i - 1]) / dividends[i - 1]) * 100;
            rateChanges.push(`Rate of change in dividends from period ${i} to ${i + 1}: ${rate.toFixed(2)}%`);
        }
    }

    return rateChanges;
}

function displayResults(allData) {
    const resultsDiv = document.getElementById('results');
    resultsDiv.innerHTML = '';

    for (const sheetName in allData) {
        const sheetResults = allData[sheetName];
        const sheetDiv = document.createElement('div');
        sheetDiv.innerHTML = `<h2>${sheetName}</h2><ul>${sheetResults.map(item => `<li>${item}</li>`).join('')}</ul>`;
        resultsDiv.appendChild(sheetDiv);
    }
}


