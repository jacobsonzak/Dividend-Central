//import { loadData as loadEarningsData } from './eps.js';
//import { loadData as loadDividendsData } from './dps.js';
import * as XLSX from 'xlsx';

//const XLSX = require('xlsx');
//const fs = require('fs');


function processExcelFile1(sheet) {

    // Convert sheet to a 2D array
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

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

function loadData1(filePath){
    // Read the file as a binary string
    const workbook = XLSX.readFile(filePath);
    var len = workbook.SheetNames.length;
    let data = {};
    for(let i = 0; i < len; i++){
        const sheetName = workbook.SheetNames[i];
        const worksheet = workbook.Sheets[sheetName];
        const point = processExcelFile1(worksheet);
        reworkData1(point);
        data[sheetName] = point;
    }
    return data;
}

function reworkData1(data){
    delete data['1'];
    for(val in data){
        data[val][1] = Number(data[val][1]);
    }
}


function processExcelFile2(sheet) {
    

    // Convert sheet to a 2D array
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

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

function loadData2(filePath){
    // Read the file as a binary string
    const workbook = XLSX.readFile(filePath);
    var len = workbook.SheetNames.length;
    let data = {};
    for(let i = 0; i < len; i++){
        const sheetName = workbook.SheetNames[i];
        const worksheet = workbook.Sheets[sheetName];
        const point = processExcelFile2(worksheet);
        reworkData2(point);
        data[sheetName] = point;
    }
    return data;
}

function reworkData2(data){
    for(val in data){
        data[val][4] = Number(data[val][4]);
    }
}








//const e_data = loadData1('10 DIV EPS .xlsx');
//const d_data = loadData2('10 MORNINGSTAR DIV STOCKS HISTORY.xlsx');


function divOverEarn(div, earn){
    //let li = []; //list to hold all values at the end
    const li = {}
    for(val in earn){ //loop thru sheet names
        let l = [];
        const elist = earn[val]; //object conatining rows
        const dlist = div[val]; //object containing rows
        const elen = Object.keys(elist).length; //number of keys in elist
        const enew = firstNKeys(elist, elen); //make into a list
        const dnew = getObjectValues(firstNKeys(dlist, elen), true); //pull the last vals from from dlist into a reversed list
        for(let i = 0; i < elen; i++){
            const num = enew[i][1];
            if(dnew[i]){
                const denom = dnew[i][4];
                l.push(num/denom);}
        }
        li[val] = l;



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

//const what = divOverEarn(d_data, e_data);
//console.log(what);

document.getElementById('uploadForm').addEventListener('submit', async function(event) {
    event.preventDefault(); // Prevent the form from submitting normally

    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];
    //const e_data = loadData1('10 DIV EPS .xlsx');
    //const d_data = loadData2('10 MORNINGSTAR DIV STOCKS HISTORY.xlsx');

    if (file1 && file2) {
        try {
            //const e_data = await loadEarningsData(file2);
            const e_data = await loadData1(file2);
            const d_data = await loadData2(file1);
            const data = divOverEarn(d_data, e_data);
            displayResults(data);
        } catch (error) {
            console.error('Error:', error);
        }
    } else {
        console.error('Please select both files.');
    }
});

function displayResults(allData) {
    console.log('Display Results Function Called');
    const resultsDiv = document.getElementById('results');
    if (!resultsDiv) {
        console.error('Element with ID "results" not found');
        return;
    }

    resultsDiv.innerHTML = '';

    for (const sheetName in allData) {
        const sheetResults = allData[sheetName];
        console.log(`Sheet Name: ${sheetName}`);
        console.log(`Sheet Results:`, sheetResults);
        const sheetDiv = document.createElement('div');
        sheetDiv.innerHTML = `<h2>${sheetName}</h2><ul>${sheetResults.map(item => `<li>${item}</li>`).join('')}</ul>`;
        resultsDiv.appendChild(sheetDiv);
    }
}