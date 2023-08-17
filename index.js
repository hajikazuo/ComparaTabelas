const XLSX = require('xlsx');
const unidecode = require('unidecode');

// Função para comparar as notas das duas tabelas
function compararNotas(tabela1, tabela2) {
  const comparacoes = [];

  tabela1.forEach((linha1) => {
    const linha2Nome1 = tabela2.find((linha) => unidecode(linha.nome1.toLowerCase()) === unidecode(linha1.nome1.toLowerCase()));
    const linha2Nome2 = tabela2.find((linha) => unidecode(linha.nome2.toLowerCase()) === unidecode(linha1.nome2.toLowerCase()));

    const comparacaoLinha = {
      nome1: linha1.nome1,
      AVM1Tabela1: linha1.AVM1,
      AVB1Tabela1: linha1.AVB1,
      TB1Tabela1: linha1.TB1,
      AP1Tabela1: linha1.AP1 || 'N/A',
      AVM1Tabela2: linha2Nome1 ? linha2Nome1.AVM1 : 'Não encontrado',
      AVB1Tabela2: linha2Nome1 ? linha2Nome1.AVB1 : 'Não encontrado',
      TB1Tabela2: linha2Nome1 ? linha2Nome1.TB1 : 'Não encontrado',
      AP1Tabela2: linha2Nome1 ? (linha2Nome1.AP1 || 'N/A') : 'Não encontrado',
      AVM1Igual: linha1.AVM1 === (linha2Nome1 ? linha2Nome1.AVM1 : null) ? 'OK' : 'Diferente',
      AVB1Igual: linha1.AVB1 === (linha2Nome1 ? linha2Nome1.AVB1 : null) ? 'OK' : 'Diferente',
      TB1Igual: linha1.TB1 === (linha2Nome1 ? linha2Nome1.TB1 : null) ? 'OK' : 'Diferente',
      AP1Igual: linha1.AP1 === (linha2Nome1 ? linha2Nome1.AP1 : null) ? 'OK' : 'Diferente',
      nome2: linha1.nome2,
      AVM2Tabela1: linha1.AVM2,
      AVB2Tabela1: linha1.AVB2,
      TB2Tabela1: linha1.TB2,
      AP2Tabela1: linha1.AP2 || 'N/A',
      AVM2Tabela2: linha2Nome2 ? linha2Nome2.AVM2 : 'Não encontrado',
      AVB2Tabela2: linha2Nome2 ? linha2Nome2.AVB2 : 'Não encontrado',
      TB2Tabela2: linha2Nome2 ? linha2Nome2.TB2 : 'Não encontrado',
      AP2Tabela2: linha2Nome2 ? (linha2Nome2.AP2 || 'N/A') : 'Não encontrado',
      AVM2Igual: linha1.AVM2 === (linha2Nome2 ? linha2Nome2.AVM2 : null) ? 'OK' : 'Diferente',
      AVB2Igual: linha1.AVB2 === (linha2Nome2 ? linha2Nome2.AVB2 : null) ? 'OK' : 'Diferente',
      TB2Igual: linha1.TB2 === (linha2Nome2 ? linha2Nome2.TB2 : null) ? 'OK' : 'Diferente',
      AP2Igual: linha1.AP2 === (linha2Nome2 ? linha2Nome2.AP2 : null) ? 'OK' : 'Diferente',
    };

    comparacoes.push(comparacaoLinha);
  });

  return comparacoes;
}

// Carrega as tabelas a partir de arquivos Excel
const workbook1 = XLSX.readFile('tabela1.xlsx');
const worksheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
const tabela1 = XLSX.utils.sheet_to_json(worksheet1);

const workbook2 = XLSX.readFile('tabela2.xlsx');
const worksheet2 = workbook2.Sheets[workbook2.SheetNames[0]];
const tabela2 = XLSX.utils.sheet_to_json(worksheet2);

const resultadoComparacao = compararNotas(tabela1, tabela2);

// Inclui na tabela de resultados alunos não presentes na tabela 1
tabela2.forEach((linha2) => {
  const linha1Nome1 = tabela1.find((linha) => unidecode(linha.nome1.toLowerCase()) === unidecode(linha2.nome1.toLowerCase()));
  const linha1Nome2 = tabela1.find((linha) => unidecode(linha.nome2.toLowerCase()) === unidecode(linha2.nome2.toLowerCase()));

  if (!linha1Nome1) {
    const comparacaoLinha = {
      nome1: 'Não encontrado',
      AVM1Tabela1: 'Não encontrado',
      AVB1Tabela1: 'Não encontrado',
      TB1Tabela1: 'Não encontrado',
      AP1Tabela1: 'Não encontrado',
      AVM1Tabela2: linha2.AVM1,
      AVB1Tabela2: linha2.AVB1,
      TB1Tabela2: linha2.TB1,
      AP1Tabela2: linha2.AP1 || 'N/A',
      AVM1Igual: 'Não encontrado',
      AVB1Igual: 'Não encontrado',
      TB1Igual: 'Não encontrado',
      AP1Igual: 'Não encontrado',
      nome2: linha2.nome2,
      AVM2Tabela1: 'Não encontrado',
      AVB2Tabela1: 'Não encontrado',
      TB2Tabela1: 'Não encontrado',
      AP2Tabela1: 'Não encontrado',
      AVM2Tabela2: linha2.AVM2,
      AVB2Tabela2: linha2.AVB2,
      TB2Tabela2: linha2.TB2,
      AP2Tabela2: linha2.AP2 || 'N/A',
      AVM2Igual: 'Não encontrado',
      AVB2Igual: 'Não encontrado',
      TB2Igual: 'Não encontrado',
      AP2Igual: 'Não encontrado',
    };

    resultadoComparacao.push(comparacaoLinha);
  }
});

// Cria uma nova planilha para os resultados
const novaPlanilha = XLSX.utils.json_to_sheet(resultadoComparacao);

// Cria um novo arquivo Excel
const novoWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(novoWorkbook, novaPlanilha, 'Resultados');

// Salva o novo arquivo Excel
XLSX.writeFile(novoWorkbook, 'resultados.xlsx', { bookType: 'xlsx', type: 'binary' });
console.log('Arquivo "resultados.xlsx" gerado com os resultados.');
