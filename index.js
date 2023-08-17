const XLSX = require('xlsx');
const unidecode = require('unidecode');

// Função para comparar as notas das duas tabelas
function compararNotas(tabela1, tabela2) {
  const comparacoes = [];


  tabela1.forEach((linha1) => {
    const linha2 = tabela2.find((linha) => unidecode(linha.nome.toLowerCase()) === unidecode(linha1.nome.toLowerCase()));

    if (linha2) {
      const comparacaoLinha = {
        nome: linha1.nome,
        AVMTabela1: linha1.AVM,
        AVBTabela1: linha1.AVB,
        TBTabela1: linha1.TB,
        APTabela1: linha1.AP || 'N/A',
        AVMTabela2: linha2.AVM,
        AVBTabela2: linha2.AVB,
        TBTabela2: linha2.TB,
        APTabela2: linha2.AP || 'N/A',
        AVMIgual: linha1.AVM === linha2.AVM ? 'OK' : 'Diferente',
        AVBIgual: linha1.AVB === linha2.AVB ? 'OK' : 'Diferente',
        TBIgual: linha1.TB === linha2.TB ? 'OK' : 'Diferente',
        APIgual: linha1.AP === linha2.AP ? 'OK' : 'Diferente',
      };

      comparacoes.push(comparacaoLinha);
    }
  });

  // Percorre a tabela2 para incluir alunos não presentes na tabela1
  tabela2.forEach((linha2) => {
    const linha1 = tabela1.find((linha) => unidecode(linha.nome.toLowerCase()) === unidecode(linha2.nome.toLowerCase()));

    if (!linha1) {
      const comparacaoLinha = {
        nome: linha2.nome,
        AVMTabela1: 'Não encontrado',
        AVBTabela1: 'Não encontrado',
        TBTabela1: 'Não encontrado',
        APTabela1: 'Não encontrado',
        AVMTabela2: linha2.AVM,
        AVBTabela2: linha2.AVB,
        TBTabela2: linha2.TB,
        APTabela2: linha2.AP || 'N/A',
        AVMIgual: 'Não encontrado',
        AVBIgual: 'Não encontrado',
        TBIgual: 'Não encontrado',
        APIgual: 'Não encontrado',
      };

      comparacoes.push(comparacaoLinha);
    }
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

// Cria uma nova planilha para os resultados
const novaPlanilha = XLSX.utils.json_to_sheet(resultadoComparacao);

// Cria um novo arquivo Excel
const novoWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(novoWorkbook, novaPlanilha, 'Resultados');

// Salva o novo arquivo Excel
XLSX.writeFile(novoWorkbook, 'resultados.xlsx', { bookType: 'xlsx', type: 'binary' });
console.log('Arquivo "resultados.xlsx" gerado com os resultados.');
