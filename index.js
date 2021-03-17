const XLSX = require('xlsx');

 
const workbook = XLSX.readFileSync("DadosColetados.xlsx", { cellFormula: false, cellHTML: false });
const sheetNames = workbook.SheetNames;

const planilha = sheetNames[0];
const rows = XLSX.utils.sheet_to_json(workbook.Sheets[planilha], { defval: ""});

//let sheets = workbook.worksheets;
//console.log(sheets.load("Form Responses 1/Email Address"));

const b = XLSX.utils.sheet_to_csv(workbook.Sheets[planilha], { defval: ""});
//const a = XLSX.utils.workbook.Sheets[planilha];



//console.log(b);


//const imprimir = ''



rows.forEach(row => {
    const colunas = Object.keys(row);
    let imprimir = `${row[colunas[4]]} foi dirigido por ${row[colunas[3]]} em ${row[colunas[7]]}`;
    //email = imprimir + email
    //console.log(rows)
    //console.log(imprimir)
    //console.log(email)
    const teste = row[colunas[5]];
    //console.log(colunas[8])    
});
//console.log(rows)
//console.log(teste)

 
//console.log(email)
 
console.log(sheetNames);

