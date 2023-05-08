import fs from 'fs';
import path from 'path';
import * as XLSX from 'xlsx';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
import { fileURLToPath } from 'url';
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
//Seteando  el sistema de archivo para XLSX
XLSX.set_fs(fs);
const workbookName = 'Reporte General - EDIC.xlsx';
const worksheetName = 'Reporte General';
const firstSurnameXlsxHeader = 'Apellido Paterno';
const lastSurnameXlsxHeader = 'Apellido Materno';
const firstnameXlsxHeader = 'Nombre (s)';
const emailXlsxHeader = 'Correo electrónico';
const resultsXlsxHeader = 'Resultado';
// Variables para cada oarchivo
const jsonsDirectoryPath = path.join(__dirname, 'jsons');
const originaXLSXPath = path.join(__dirname, 'origin/originalData.xlsx');
// Arreglo de
const directoryContent = fs.readdirSync(jsonsDirectoryPath);
// Arreglo de rutas
const pathsArray = directoryContent.map(json => path.join(jsonsDirectoryPath, json));
// Arreglo de resultados JSON
const jsonResults = pathsArray.map(path => require(path));
jsonResults.forEach(json => json.ranges.forEach(range => { if (range.score < 3)
    range.score = range.score + 1; }));
const workbook = XLSX.readFile(originaXLSXPath);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const originalUsersArr = [];
for (const address in sheet) {
    //console.log(address);
    if (address.includes('C')) {
        const _completeName = sheet[address].v;
        const wordsArr = _completeName.split(' ');
        const capitalizedWordsArr = wordsArr.map(word => word.toUpperCase());
        //const capitalizedCompletedName = capitalizedWordsArr.join(' ');
        let splitedAddress = address.split('');
        splitedAddress[0] = 'D';
        const emailAddress = splitedAddress.join('');
        const _email = sheet[emailAddress].v;
        const user = {
            email: _email,
            fullName: capitalizedWordsArr
        };
        originalUsersArr.push(user);
    }
}
// Arreglo de usuarios por fila
/*
jsonResults.forEach(result => {
    
    for(let i = 0; i < originalUsersArr.length; i++) {
        
        let app : number = 0;

        const normalizedAndSplitedFullName = originalUsersArr[i].fullName.split(' ').map(name => name.toLowerCase()).map(e => e.normalize("NFD").replace(/[\u0300-\u036f]/g, ""));

        normalizedAndSplitedFullName.forEach(name => {
            if(result.username.includes(name)) { app++ }
        })

        let inserted = false;
        

        if(app >= 2 && !inserted) {

            const calculateTotal = (userRanges : Array<Range>) : number => {
                let subtotal : number = 0;
                userRanges.forEach( range => {
                    subtotal += range.score;
                })
                let total = ((subtotal * 10) / 21) + 1.5;
                total = parseFloat(total.toFixed(2));
                return total >= 10 ? 10 :  total;
            }
    
            const total : number = calculateTotal(result.ranges);

            userRows.push(
                {
                    firstSurname: normalizedAndSplitedFullName[normalizedAndSplitedFullName.length - 3].toUpperCase(),
                    lastSurname: normalizedAndSplitedFullName[normalizedAndSplitedFullName.length - 2].toUpperCase(),
                    firstname: normalizedAndSplitedFullName[0].toUpperCase(),
                    email: originalUsersArr[i].email,
                    results: total
                })
                break;
            }
    }

});
*/
const lowerFullNames = originalUsersArr.map(user => user.fullName.map(name => name.toLowerCase()).map(e => e.normalize("NFD").replace(/[\u0300-\u036f]/g, "")));
const extractedNames = lowerFullNames.map(names => {
    const tempArr = [];
    tempArr.push(names[0]);
    tempArr.push(names[1]);
    return tempArr;
});
// console.log(lowerFullNames);
const userRows = jsonResults.map(json => {
    const name = json.username.replace('@edicmexico.onmicrosoft.com', '');
    console.log(name);
    for (let i = 0; i < lowerFullNames.length; i++) {
        let appearances = 0;
        for (let j = 0; j < lowerFullNames[i].length; j++) {
            if (json.username.includes(lowerFullNames[i][j])) {
                appearances++;
            }
        }
        if (appearances > 1) {
            const calculateTotal = (userRanges) => {
                let subtotal = 0;
                userRanges.forEach(range => {
                    subtotal += range.score;
                });
                let total = ((subtotal * 10) / 21) + 1.5;
                total = parseFloat(total.toFixed(2));
                return total >= 10 ? 10 : total;
            };
            const user = {
                firstSurname: originalUsersArr[i].fullName[originalUsersArr[i].fullName.length - 3],
                lastSurname: originalUsersArr[i].fullName[originalUsersArr[i].fullName.length - 2],
                firstname: originalUsersArr[i].fullName[0],
                email: originalUsersArr[i].email,
                results: calculateTotal(json.ranges)
            };
            originalUsersArr.splice(i, 1);
            lowerFullNames.splice(i, 1);
            return user;
            break;
        }
    }
    return {
        email: '',
        firstSurname: '',
        lastSurname: '',
        firstname: '',
        results: 0
    };
});
console.log(userRows.length);
/* Create worksheet */
var data = [
    [firstSurnameXlsxHeader, lastSurnameXlsxHeader, firstnameXlsxHeader, emailXlsxHeader, resultsXlsxHeader]
];
let userRowAoA = userRows.map(user => {
    return Object.values(user);
});
// userRowAoA.unshift(data[0]);
userRowAoA.unshift(data[0]);
let newWorksheet = XLSX.utils.aoa_to_sheet(userRowAoA);
/* Create workbook */
let newWorkbook = XLSX.utils.book_new();
/* Add the worksheet to the workbook */
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, worksheetName);
/* Write to file */
XLSX.writeFile(newWorkbook, workbookName);
//var data = XLSX.write(newWorkbook);
