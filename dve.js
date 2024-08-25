//document.getElementById('excelFile').addEventListener('change', processFile, false);
/*
function handleFile(event) {
    const file1 = event.target.files[0];
    const file2 = event.target.files[1];
    if (!file1 || !file2) {
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
}*/

function processFile() {
    console.log('processFile function called');
    const fileInput = document.getElementById('excelFile');
    const files = fileInput.files;
    

    console.log('Files selected:', files);

    if (files.length !== 2) {
        alert('Please select exactly two Excel files.');
        return;
    }
    // Log the file names
    for (let i = 0; i < files.length; i++) {
        console.log(`Selected file ${i + 1}: ${files[i].name}`);
    }

    const reader1 = new FileReader();
    const reader2 = new FileReader();

    reader1.onload = function (e) {
        const data1 = new Uint8Array(e.target.result);
        const workbook1 = XLSX.read(data1, { type: 'array' });

        // Load data from the first workbook
        const dataFromFile1 = loadData(workbook1, 1);

        reader2.onload = function (e) {
            const data2 = new Uint8Array(e.target.result);
            const workbook2 = XLSX.read(data2, { type: 'array' });

            // Load data from the second workbook
            const dataFromFile2 = loadData(workbook2, 4);

            // Process and display results
            const combinedData = divOverEarn(dataFromFile2, dataFromFile1);
            displayResults(combinedData);
        };

        reader2.readAsArrayBuffer(files[1]); // Read the second file
    };

    reader1.readAsArrayBuffer(files[0]); // Read the first file
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




function divOverEarn(div, earn){
    //console.log(`div is ${div} and earn is ${earn}`)
    const li = {}
    for(val in earn){ //loop thru sheet names
        let l = [];
        const elist = earn[val]; //object conatining rows
        //console.log(`elist is ${elist}`)
        const dlist = div[val]; //object containing rows
        //console.log(`dlist is ${dlist}`)
        const elen = Object.keys(elist).length; //number of keys in elist
        //console.log(`elen is ${elen}`)
        const enew = firstNKeys(elist, elen); //make into a list
        //console.log(`enew is ${enew}`)
        const dnew = getObjectValues(firstNKeys(dlist, elen), true); //pull the last vals from from dlist into a reversed list
        //console.log(`dnew is ${dnew}`)
        for(let i = 0; i < elen; i++){
            const num = enew[i][1];
            console.log(`enum is ${num}`)
            if (dnew[i]) {
                const denom = dnew[i][4];
                //console.log(`denom is ${denom}`)
                if (denom !== 0) {
                    l.push(num / denom);
                    //console.log(`num / denom is ${l[i]}`)
                } else {
                    console.warn(`Division by zero for index ${i}`);
                    l.push('Inf'); // or handle it differently
                }
        }
        li[val] = l;
        //console.log(`${li}`)



    }
}
    return li;
}


function getLastNKeys(data, n) {
    // Extract all keys from the object
    const allKeys = Object.keys(data);
    
    // Get the last n keys
    const lastNKeys = allKeys.slice(-n);
    
    // Create a new object with only the last n keys
    const result = {};
    lastNKeys.forEach(key => {
        result[key] = data[key];
    });
    
    return result;
}


function getObjectValues(data, reverse = false) {
    // Extract all values from the object
    let values = Object.values(data);
    
    // Reverse the values if the reverse option is true
    if (reverse) {
        values = values.reverse();
    }
    return values;
}



function firstNKeys(data, n, reverse = false) {
    // Get an array of keys from the object
    const keys = Object.keys(data);

    // Slice the first n keys
    const firstKeys = keys.slice(0, n);

    // Extract the values corresponding to the first n keys
    let values = firstKeys.map(key => data[key]);

    // Reverse the values if the reverse option is true
    if (reverse) {
        values = values.reverse();
    }

    return values;
}


function loadData(workbook, index){
    // Read the file as a binary string
    var len = workbook.SheetNames.length;
    let data = {};
    for(let i = 0; i < len; i++){
        const sheetName = workbook.SheetNames[i];
        const worksheet = workbook.Sheets[sheetName];
        const point = processExcelFile(worksheet);
        reworkData(point, index);
        data[sheetName] = point;
    }
    return data;
}

function reworkData(data, index){
    delete data[index.toString()];
    for(val in data){
        data[val][index] = Number(data[val][index]);
    }
}

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