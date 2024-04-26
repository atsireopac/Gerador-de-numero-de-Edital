const XLSX = require('xlsx');

// Função para verificar se o número de edital já existe no banco de dados
function verificarExistenciaEdital(numeroEdital) {
    // Carregar o arquivo Excel
    const workbook = XLSX.readFile('./Desenvolvimento/ambientes/Microsoft VS Code/workspace/CursoEmVideo/Gerador-de-numero-de-Edital/Pasta-Copiar.xlsx');
    
    // Selecionar a primeira planilha do arquivo
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Converter a planilha para um objeto JSON
    const data = XLSX.utils.sheet_to_json(sheet);

    // Verificar se o número de edital já existe no banco de dados
    for (let i = 0; i < data.length; i++) {
        if (data[i].NumeroEdital === numeroEdital) {
            return true; // O número de edital já existe
        }
    }

    return false; // O número de edital não existe
}

// Função para gerar o número de acompanhamento do edital
function gerarNumeroEdital(numeroOrdem, anoCadastro) {
    // Verifica se o número de ordem e o ano de cadastro foram fornecidos
    if (!numeroOrdem || !anoCadastro) {
        return "Número de ordem e ano de cadastro são obrigatórios.";
    }

    // Verifica se o número de ordem é um número válido
    if (isNaN(numeroOrdem)) {
        return "O número de ordem deve ser um número válido.";
    }

    // Verifica se o ano de cadastro é um número válido e tem quatro dígitos
    if (isNaN(anoCadastro) || anoCadastro.toString().length !== 4) {
        return "O ano de cadastro deve ser um número válido com quatro dígitos.";
    }

    // Formata o número de ordem e o ano de cadastro no formato desejado
    const numeroEdital = numeroOrdem + "/" + anoCadastro;

    // Verifica se o número de edital já existe no banco de dados
    if (verificarExistenciaEdital(numeroEdital)) {
        return "O número de edital já existe no banco de dados.";
    }

    return numeroEdital;
}

// Exemplo de uso da função para gerar um número de acompanhamento de edital
var numeroEdital = gerarNumeroEdital(51, 2024);
console.log(numeroEdital); // Saída: 51/2024
