const XLSX = require('xlsx');//Importando pacote de manipulação de planilhas

 
const workbook = XLSX.readFileSync("DadosColetados.xlsx", { cellFormula: false, cellHTML: false });//Importando o arquivo xlsx(Excel)
const sheetNames = workbook.SheetNames;//Buscando o nomes das planilhas(Ficam no canto inferior dentro do arquivo excel

const planilha = sheetNames[0];//Criando constante para a primeira planilha
const rows = XLSX.utils.sheet_to_json(workbook.Sheets[planilha], { defval: ""});//Importando a primeira planilha para JSON para pode manipular

//let sheets = workbook.worksheets; 
//console.log(sheets.load("Form Responses 1/Email Address"));

const b = XLSX.utils.sheet_to_csv(workbook.Sheets[planilha], { defval: ""});
//const a = XLSX.utils.workbook.Sheets[planilha];



//console.log(b);


//const imprimir = ''



rows.forEach(row => {
    const colunas = Object.keys(row);
    let imprimir = `${row[colunas[4]]} + ${row[colunas[3]]} + ${row[colunas[7]]}`;//Percorrer o JSON para criar uma frase por linha
    //email = imprimir + email
    console.log(rows)//Imprimir todo o JSON
    console.log(imprimir)//Percorrer o JSON para imprimir uma frase por linha
    //console.log(email)
    const teste = row[colunas[5]];
    //console.log(colunas[8])    
});
console.log(rows)//Imprimindo todo o JSON
//console.log(teste)

 
//console.log(email)
 
console.log(sheetNames);//Imprimindo o nome da planilha

