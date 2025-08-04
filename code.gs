function doGet() {
  return HtmlService.createHtmlOutputFromFile('formulario')
    .setTitle('Consulta de Frequência');
}

// Retorna a lista de nomes dos alunos (coluna A, exceto cabeçalho)
function listarNomesAlunos() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dados = planilha.getRange(2, 1, planilha.getLastRow() - 1, 1).getValues();
  return dados.flat();
}


function verificarFrequencia(nome, emailDigitado) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abas = planilha.getSheets();

  let encontrouNome = false;

  for (let aba of abas) {
    const dados = aba.getDataRange().getValues();
    const cabecalhos = dados[0];

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const nomePlanilha = linha[0];
      const emailPlanilha = linha[1];

      if (nomePlanilha === nome) {
        encontrouNome = true;

        if (emailPlanilha.trim().toLowerCase() === emailDigitado.trim().toLowerCase()) {
          const situacoes = linha.slice(2, linha.length - 1);
          const percentual = linha[linha.length - 1];
          let classe = '';

          if (percentual >= 75) classe = 'frequencia-verde';
          else if (percentual >= 60) classe = 'frequencia-amarelo';
          else classe = 'frequencia-vermelho';

          let tabelaHTML = `<table class="table table-striped"><thead><tr>`;
          for (let j = 2; j < cabecalhos.length - 1; j++) {
            tabelaHTML += `<th>${cabecalhos[j]}</th>`;
          }
          tabelaHTML += `</tr></thead><tbody><tr>`;

          situacoes.forEach(situacao => {
            if (situacao === 'Presente') {
              tabelaHTML += `<td>✅</td>`;
            } else if (situacao === 'Ausente de Entrega') {
              tabelaHTML += `<td>⚠️</td>`;
            } else {
              tabelaHTML += `<td>❌</td>`;
            }
          });

          tabelaHTML += `</tr></tbody></table>`;
          tabelaHTML += `<strong class="mb-3 mt-3">Legenda da Planilha</strong><p>✅ Presente</p><p>⚠️ Ausência Justificada | Ausência de Atividade na Aula Indicada</p><p>❌ Falta</p>`;

          return {
            mensagem: `<div class="${classe}">Frequência Total (${aba.getName()}): ${percentual}%</div><br>${tabelaHTML}`,
            classe: ''
          };
        }
      }
    }
  }

  if (encontrouNome) {
    return {
      mensagem: `<div class="frequencia-vermelho">⚠️ Email incorreto para "${nome}". Verifique e tente novamente.</div>`,
      classe: ''
    };
  }

  return {
    mensagem: `<div class="frequencia-vermelho">Aluno "${nome}" não encontrado em nenhuma turma.</div>`,
    classe: ''
  };
}

