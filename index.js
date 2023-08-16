const XLSX = require('xlsx');

// Função para comparar as notas das duas tabelas
function compararNotas(tabela1, tabela2) {
  const comparacoes = [];

  tabela1.forEach((linha1) => {
    const linha2 = tabela2.find((linha) => linha.nome.toLowerCase() === linha1.nome.toLowerCase());

    if (linha2) {
      const comparacaoLinha = {
        nome: linha1.nome,
        nota1Tabela1: linha1.nota1,
        nota2Tabela1: linha1.nota2,
        nota3Tabela1: linha1.nota3,
        nota4Tabela1: linha1.nota4 || 'N/A',
        nota1Tabela2: linha2.nota1,
        nota2Tabela2: linha2.nota2,
        nota3Tabela2: linha2.nota3,
        nota4Tabela2: linha2.nota4 || 'N/A',
        nota1Igual: linha1.nota1 === linha2.nota1 ? 'OK' : 'Diferente',
        nota2Igual: linha1.nota2 === linha2.nota2 ? 'OK' : 'Diferente',
        nota3Igual: linha1.nota3 === linha2.nota3 ? 'OK' : 'Diferente',
        nota4Igual: linha1.nota4 === linha2.nota4 ? 'OK' : 'Diferente',
      };

      comparacoes.push(comparacaoLinha);
    }
  });

  return comparacoes;
}

// Carregar tabelas a partir de arquivos Excel
const workbook1 = XLSX.readFile('tabela1.xlsx');
const worksheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
const tabela1 = XLSX.utils.sheet_to_json(worksheet1);

const workbook2 = XLSX.readFile('tabela2.xlsx');
const worksheet2 = workbook2.Sheets[workbook2.SheetNames[0]];
const tabela2 = XLSX.utils.sheet_to_json(worksheet2);

const resultadoComparacao = compararNotas(tabela1, tabela2);

// Criar uma nova planilha para os resultados
const novaPlanilha = XLSX.utils.json_to_sheet(resultadoComparacao);

// Criar um novo arquivo Excel
const novoWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(novoWorkbook, novaPlanilha, 'Resultados');

// Salvar o novo arquivo Excel
XLSX.writeFile(novoWorkbook, 'resultados.xlsx', { bookType: 'xlsx', type: 'binary' });
console.log('Arquivo "resultados.xlsx" gerado com os resultados.');
