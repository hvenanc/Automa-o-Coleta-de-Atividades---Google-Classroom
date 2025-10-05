function getResponsesAsObject(
  formId,
  nomeQuestionTitle = "E-mail (atenção adicionar e-mail recebido na sua escola)",
  linkQuestionTitle = "Adicione aqui o link do miniprojeto"
) {
  const form = FormApp.openById(formId);
  const formResponses = form.getResponses();
  const studentData = [];

  for (const formResponse of formResponses) {
    const itemResponses = formResponse.getItemResponses();
    let nome = "";
    let link = "";

    for (const itemResponse of itemResponses) {
      //O Trim deve adicionado devido aos espaços deixados nas perguntas do Miniprojeto.
      const questionTitle = itemResponse.getItem().getTitle().trim();
      const response = itemResponse.getResponse();

      if (questionTitle === nomeQuestionTitle) {
        nome = formResponse.getRespondentEmail() || "";
      }
      if (questionTitle === linkQuestionTitle) {
        link = response || "";
      }
    }

    studentData.push({
      nome,
      link: link ? (verificaLinkPublico(link) ? "Presente" : "Link Privado") : "Sem link",
    });
  }

  // Ordena o array pelo nome de forma mais simples
  studentData.sort((a, b) => a.nome.localeCompare(b.nome));

  return studentData;
}


function getFormId(formLink) {
  // Extrai o ID do forms pela URL.
  const preId = formLink.substring(32);
  return preId.replace("/edit", "");
}

function verificaLinkPublico(url) {
  try {
    Logger.log(url)
    // Expressão regular genérica: captura o ID entre /d/ e /edit ou /view
    const regex = /\/d\/([a-zA-Z0-9_-]+)(?:\/|$)/;
    const match = url.match(regex);

    if (!match || match.length < 2) {
      // A URL não contém um ID válido
      return false;
    }

    const fileId = match[1];
    let file;

    try {
      file = DriveApp.getFileById(fileId);
    } catch (err) {
      return false;
    }

    const access = file.getSharingAccess();

    // Verifica se está público para qualquer pessoa com link ou totalmente público
    return access === DriveApp.Access.ANYONE_WITH_LINK || access === DriveApp.Access.ANYONE;

  } catch (e) {
    Logger.log('Erro ao verificar o link: ' + e.message);
    return false;
  }
}

function verificarAtividadesPorTurma() {

  const TURMAS = [
    { id: 'ID - TURMA 1', nome: 'Turma 01' },
    { id: 'ID - TURMA 2', nome: 'Turma 02' }
  ];

  //Ids das turmas do Classroom e Criação da Planilha com as Informações das Atividades.
  const PLANILHA = SpreadsheetApp.getActiveSpreadsheet();
  
  //Filtra as Atividades do Classroom
  const FILTROS = ['Miniprojeto', 'Desafio'];

  TURMAS.forEach(turma => {

    const nomeAbaAtualPlanilha = turma.nome;
    let abaAtual = PLANILHA.getSheetByName(nomeAbaAtualPlanilha);

    if(abaAtual) abaAtual.clear()
    else abaAtual = PLANILHA.insertSheet(nomeAbaAtualPlanilha);

    //Busca todos os temas da Turma ou retorna vazio, caso não haja nenhum tópico.
    let atividades = Classroom.Courses.CourseWork.list(turma.id).courseWork || [];
    atividades = atividades.filter(a =>
      FILTROS.some(f => a.title && a.title.toLowerCase().includes(f.toLowerCase())) && a.topicId
    );

    if (atividades.length === 0) {
      abaAtual.appendRow(['Nenhuma atividade encontrada com os filtros:', FILTROS.join(', ')]);
      return;
    }

    // Agrupa atividades por tema (topicId)
    const temas = {};
    atividades.forEach(a => {
      if (!temas[a.topicId]) temas[a.topicId] = [];
      temas[a.topicId].push(a);
    });

    // Mapa de nomes de tópicos
    const topicos = Classroom.Courses.Topics.list(turma.id).topic || [];
    const mapaTopicos = {};
    topicos.forEach(t => {
      mapaTopicos[t.topicId] = t.name;
    });
    
    // Monta cabeçalho: Nome, Email, para cada tema Dia 1 (Desafios) e Dia 2 (Miniprojeto)
    const cabecalho = ['Nome do Aluno', 'Email'];
    const listaTopicosOrdenada = Object.keys(temas).map(id => ({
      id,
      nome: mapaTopicos[id] || 'Sem Tema',
      atividades: temas[id]
    })).sort((a, b) => a.nome.localeCompare(b.nome));

    listaTopicosOrdenada.forEach(t => {
      cabecalho.push(`${t.nome} - Dia 1 (Desafios)`);
      cabecalho.push(`${t.nome} - Dia 2 (Miniprojeto)`);
    });
    cabecalho.push('Frequência (%)');
    abaAtual.appendRow(cabecalho);

    // Pré-busca: para cada atividade coleta formUrl (se houver) e lista de submissões (uma única chamada por atividade)
    listaTopicosOrdenada.forEach(topico => {
      topico.desafios = [];
      topico.miniprojetos = [];
      topico.atividades.forEach(atividade => {
        const formMaterial = atividade.materials && atividade.materials.find(m => m.form);
        const urlForm = formMaterial && formMaterial.form && formMaterial.form.formUrl;
        let submissions = [];
        try {
          submissions = Classroom.Courses.CourseWork.StudentSubmissions.list(turma.id, atividade.id).studentSubmissions || [];
        } catch (e) {
          Logger.log('Erro ao buscar StudentSubmissions para atividade ' + atividade.id + ': ' + e);
          submissions = [];
        }
        if (/desafio/i.test(atividade.title || '')) {
          topico.desafios.push({ atividade, urlForm, submissions });
        } else if (/miniprojeto/i.test(atividade.title || '')) {
          topico.miniprojetos.push({ atividade, urlForm, submissions });
        }
      });
    });

    // Obtem todos os estudantes do Classroom e seus respectivos E-mails.
    const alunos = Classroom.Courses.Students.list(turma.id).students || [];
    alunos.sort((a, b) => a.profile.name.fullName.localeCompare(b.profile.name.fullName));

    // Nesse cenário eu separei em dois tópicos por aula - Dia 1 (Desafios) | Dia 2 (Miniprojeto) = listaDeTópicos * 2;
    const totalDias = listaTopicosOrdenada.length * 2;

    alunos.forEach(aluno => {
      const linha = [aluno.profile.name.fullName, aluno.profile.emailAddress];
      let presencas = 0;

      listaTopicosOrdenada.forEach(topico => {
        // DIA 1: avaliar 3 desafios
        let totalDesafiosEntregues = 0;
        topico.desafios.forEach(desafio => {
          let entregue = false;
          if (desafio.urlForm) {
            // formulário -> verificar por email
            entregue = verificaRespostaFormulario(aluno.profile.emailAddress, desafio.urlForm);
          } else {
            const entregaAluno = desafio.submissions.find(s => s.userId === aluno.userId);
            if (entregaAluno && (entregaAluno.state === 'TURNED_IN' || entregaAluno.state === 'RETURNED')) {
              entregue = true;
            }
          }
          if (entregue) totalDesafiosEntregues++;
        });

        let situacaoDia1 = '';
        if (totalDesafiosEntregues >= 3) {
          situacaoDia1 = 'Presente';
          presencas++;
          //Possível Bug no &&
        } else if (totalDesafiosEntregues > 0 && totalDesafiosEntregues < 3) {
          situacaoDia1 = 'Ausente de Entrega';
        } else {
          situacaoDia1 = 'Falta';
        }
        linha.push(situacaoDia1);

        // DIA 2: avalia o miniprojeto  
        let situacaoDia2 = 'Falta'; // valor padrão
        topico.miniprojetos.forEach(miniprojeto => {
          if (situacaoDia2 !== 'Falta') return; // já encontrou situação válida

          if (miniprojeto.urlForm) {
            // Aqui recebo o Link do Forms do Miniprojeto.
            const respostas = getResponsesAsObject(getFormId(miniprojeto.urlForm));
            const respostaAluno = respostas.find(r => r.nome.toLowerCase() === aluno.profile.emailAddress.toLowerCase());

            if (respostaAluno) {
              if (respostaAluno.link === "Presente") {
                situacaoDia2 = "Presente";
                presencas++;
              } else if (respostaAluno.link === "Link Privado") {
                situacaoDia2 = "* LINK PRIVADO *";
              } else {
                situacaoDia2 = "Falta";
              }
            }
          } else {
            const entregaAluno = miniprojeto.submissions.find(s => s.userId === aluno.userId);
            if (entregaAluno && (entregaAluno.state === 'TURNED_IN' || entregaAluno.state === 'RETURNED')) {
              situacaoDia2 = "Presente";
              presencas++;
            }
          }
        });
        linha.push(situacaoDia2);
      });

      const frequencia = totalDias > 0 ? Math.round((presencas / totalDias) * 100) : 0;
      linha.push(frequencia);
      abaAtual.appendRow(linha);
    });
  });
}

// Verifica se o aluno respondeu ao formulário usando FormsApp
function verificaRespostaFormulario(email, urlForm) {
  try {
    const form = FormApp.openByUrl(urlForm);
    const respostas = form.getResponses();

    for (const resposta of respostas) {
        const respondente = resposta.getRespondentEmail();
        if (respondente && respondente.toLowerCase() === (email || '').toLowerCase()) {
          return true;
        }
    }
  } catch (e) {
    Logger.log("Erro ao acessar formulário: " + e);
  }
  return false;
}
