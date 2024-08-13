const express = require('express');
const xlsx = require('xlsx');
const fs = require('fs-extra');
const path = require('path');
const cron = require('node-cron');

const app = express();

// Função para obter o diretório do executável
const getExecutableDir = () => {
    return __dirname; // Retorna o diretório onde o script está sendo executado
};

// Função para converter serial de data do Excel para formato legível
function excelDateToJSDate(serial) {
    const excelEpoch = new Date(1899, 11, 30); // Data base do Excel
    const date = new Date(excelEpoch.getTime() + (serial * 86400000)); // 86400000 ms por dia
    const year = date.getFullYear();
    const month = ("0" + (date.getMonth() + 1)).slice(-2); // Adiciona zero à esquerda
    const day = ("0" + date.getDate()).slice(-2); // Adiciona zero à esquerda
    return `${year}-${month}-${day}`; // Formato final da data
}

// Função para converter planilha Excel em JSON
function excelToJson(excelFilePath, jsonFilePath, startRow = 3) {
    console.log(`Iniciando a conversão do Excel para JSON, a partir da linha ${startRow}.`);
    if (fs.existsSync(jsonFilePath)) {
        console.log('Arquivo JSON existente encontrado. Removendo...');
        fs.removeSync(jsonFilePath);
    }
    console.log('Carregando o arquivo Excel:', excelFilePath);
    const workbook = xlsx.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    const data = jsonData.slice(startRow - 1).map(row => ({
        cliente: row[0],    // Coluna A - Nome do Cliente
        debito: row[1],     // Coluna B - Débito
        credito: row[2],    // Coluna C - Crédito
        data: excelDateToJSDate(row[3])  // Coluna D - Data convertida
    }));
    fs.writeJson(jsonFilePath, data, { spaces: 2 })
        .then(() => console.log(`Arquivo JSON salvo em: ${jsonFilePath}`))
        .catch(err => console.error('Erro ao salvar o arquivo JSON:', err));
}

// Caminho para os arquivos Excel e os arquivos JSON de saída
const excelFilePathCredDeb = path.join(getExecutableDir(), 'CRED E DEB.xlsx');
const jsonFilePathCredDeb = path.join(getExecutableDir(), 'output.json');
const excelFilePathBlocos = path.join(getExecutableDir(), 'BLOCOS.xlsx');
const jsonFilePathBlocos = path.join(getExecutableDir(), 'responses.json');

// Agendar a execução a cada hora, das 07:30 até 17:30
cron.schedule('30 7-17 * * *', () => {
    console.log(`Cron job iniciado às ${new Date().toLocaleTimeString()}`);
    console.log('Iniciando a conversão do Excel para JSON (CRED E DEB).');
    excelToJson(excelFilePathCredDeb, jsonFilePathCredDeb);
}, {
    timezone: "America/Sao_Paulo"
});

cron.schedule('30 7-17 * * *', () => {
    console.log(`Cron job iniciado às ${new Date().toLocaleTimeString()}`);
    console.log('Iniciando a conversão do Excel para JSON (BLOCOS).');
    excelToJsonBlocos(excelFilePathBlocos, jsonFilePathBlocos);
}, {
    timezone: "America/Sao_Paulo"
});

// API para acessar dados do JSON
app.get('/creddeb', (req, res) => {
    const jsonFilePath = path.join(getExecutableDir(), 'output.json');
    if (fs.existsSync(jsonFilePath)) {
        const data = fs.readJsonSync(jsonFilePath);
        res.json(data);
    } else {
        res.status(404).json({ error: 'Arquivo JSON de CRED E DEB não encontrado' });
    }
});

app.get('/blocos', (req, res) => {
    const jsonFilePath = path.join(getExecutableDir(), 'responses.json');
    if (fs.existsSync(jsonFilePath)) {
        const data = fs.readJsonSync(jsonFilePath);
        res.json(data);
    } else {
        res.status(404).json({ error: 'Arquivo JSON de BLOCOS não encontrado' });
    }
});

// Inicia o servidor
const PORT = process.env.PORT || 3333;
app.listen(PORT, () => {
    console.log(`API rodando em http://localhost:${PORT}`);
});

