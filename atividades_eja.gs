function registrarSituacaoPorTemas() {
  const TURMAS = [
    { id: 'a', nome: 'Turma 01' },
    { id: 'b', nome: 'Turma 02' }
  ];

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  TURMAS.forEach(turma => {
    const nomeAba = turma.nome;
    let aba = planilha.getSheetByName(nomeAba);

    if (aba) aba.clear();
    else aba = planilha.insertSheet(nomeAba);

    // Busca todos os tópicos (temas) da turma
    const topicos = Classroom.Courses.Topics.list(turma.id).topic || [];
    const mapaTopicos = {};
    topicos.forEach(t => mapaTopicos[t.topicId] = t.name);

    // Busca todas as atividades da turma
    let atividades = Classroom.Courses.CourseWork.list(turma.id).courseWork || [];

    // Agrupa atividades por tema
    const temas = {};
    atividades.forEach(a => {
      if (a.topicId) {
        if (!temas[a.topicId]) temas[a.topicId] = [];
        temas[a.topicId].push(a);
      }
    });

    // Cabeçalho dinâmico
    const cabecalho = ['Nome do Aluno', 'Email'];
    Object.keys(temas).forEach(topicId => {
      const nomeTema = mapaTopicos[topicId] || 'Sem Tema';
      for (let dia = 1; dia <= 4; dia++) {
        cabecalho.push(`${nomeTema} - Dia ${dia}`);
      }
    });
    cabecalho.push('Frequência (%)');
    aba.appendRow(cabecalho);

    // Lista de alunos
    const alunos = Classroom.Courses.Students.list(turma.id).students || [];
    alunos.sort((a, b) => a.profile.name.fullName.localeCompare(b.profile.name.fullName));

    alunos.forEach(aluno => {
      const linha = [aluno.profile.name.fullName, aluno.profile.emailAddress];
      let totalPresencas = 0;
      let totalDias = 0;

      Object.keys(temas).forEach(topicId => {
        const atividadesTema = temas[topicId];
        for (let dia = 1; dia <= 4; dia++) {
          const situacao = calcularSituacaoDia(turma.id, aluno, atividadesTema, dia);
          linha.push(situacao);
          totalDias++;
          if (situacao === 'Presente') totalPresencas++;
        }
      });

      const frequencia = totalDias > 0 ? Math.round((totalPresencas / totalDias) * 100) : 0;
      linha.push(frequencia);
      aba.appendRow(linha);
    });
  });
}

// Aplica as regras de presença por dia para um conjunto de atividades de um tema
function calcularSituacaoDia(idTurma, aluno, atividades, dia) {
  let totalEnvios = 0;
  let atividadesFiltradas = [];

  // Filtra as atividades relevantes para cada dia
  if (dia === 1) {
    atividadesFiltradas = atividades.filter(a => a.title?.toLowerCase().includes('broadcast'));
  } else if (dia === 2) {
    atividadesFiltradas = atividades.filter(a => a.title?.toLowerCase().includes('desafio'));
  } else if (dia === 3 || dia === 4) {
    atividadesFiltradas = atividades.filter(a => a.title?.toLowerCase().includes('miniprojeto'));
  }

  atividadesFiltradas.forEach(atividade => {
    const entregas = Classroom.Courses.CourseWork.StudentSubmissions.list(idTurma, atividade.id).studentSubmissions || [];
    const entregaAluno = entregas.find(e => e.userId === aluno.userId);
    const urlForm = atividade.materials?.find(m => m.form)?.form?.formUrl;

    let enviado = false;
    if (urlForm) {
      enviado = alunoRespondeuFormularioFormsApp(aluno.profile.emailAddress, urlForm);
    } else if (entregaAluno && (entregaAluno.state === 'TURNED_IN' || entregaAluno.state === 'RETURNED')) {
      enviado = true;
    }

    if (enviado) totalEnvios++;
  });

  // Regras por dia
  if (dia === 1) {
    return (totalEnvios > 0) ? 'Presente' : 'Falta';
  } else if (dia === 2) {
    if (totalEnvios >= 3) return 'Presente';
    else if (totalEnvios > 0) return 'Ausência de Entrega';
    else return 'Falta';
  } else if (dia === 3 || dia === 4) {
    return (totalEnvios > 0) ? 'Presente' : 'Falta';
  }

  return '';
}

// Verifica se o aluno respondeu ao formulário
function alunoRespondeuFormularioFormsApp(email, urlForm) {
  try {
    const form = FormApp.openByUrl(urlForm);
    const respostas = form.getResponses();
    return respostas.some(r => r.getRespondentEmail() === email);
  } catch (e) {
    Logger.log("Erro ao acessar formulário: " + e);
    return false;
  }
}
