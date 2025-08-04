function registrarSituacaoPorTemas() {
  const TURMAS = [
    { id: '789996876690', nome: 'Turma 01' },
    { id: '700438412127', nome: 'Turma 02' }
  ];
  const FILTROS = ['Miniprojeto', 'Desafio'];

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  TURMAS.forEach(turma => {
    const nomeAba = turma.nome;
    let aba = planilha.getSheetByName(nomeAba);

    // Se a aba já existe, limpa. Se não, cria.
    if (aba) {
      aba.clear();
    } else {
      aba = planilha.insertSheet(nomeAba);
    }

    // Buscar atividades da turma
    let atividades = Classroom.Courses.CourseWork.list(turma.id).courseWork || [];
    atividades = atividades.filter(a =>
      FILTROS.some(f => a.title.toLowerCase().includes(f.toLowerCase())) && a.topicId
    );

    if (atividades.length === 0) {
      aba.appendRow(['Nenhuma atividade encontrada com os filtros:', FILTROS.join(', ')]);
      return;
    }

    // Agrupar por tema
    const temas = {};
    atividades.forEach(a => {
      if (!temas[a.topicId]) temas[a.topicId] = [];
      temas[a.topicId].push(a);
    });

    // Buscar nomes dos temas
    const topicos = Classroom.Courses.Topics.list(turma.id).topic || [];
    const mapaTopicos = {};
    topicos.forEach(t => {
      mapaTopicos[t.topicId] = t.name;
    });

    // Cabeçalho
    const cabecalho = ['Nome do Aluno', 'Email'];
    const listaTopicosOrdenada = Object.keys(temas).map(id => ({
      id,
      nome: mapaTopicos[id] || 'Sem Tema',
      atividades: temas[id]
    })).sort((a, b) => a.nome.localeCompare(b.nome));

    listaTopicosOrdenada.forEach(t => cabecalho.push(t.nome));
    cabecalho.push('Frequência (%)');
    aba.appendRow(cabecalho);

    // Buscar alunos
    const alunos = Classroom.Courses.Students.list(turma.id).students || [];
    alunos.sort((a, b) => a.profile.name.fullName.localeCompare(b.profile.name.fullName));

    const totalTopicos = listaTopicosOrdenada.length;

    // Para cada aluno
    alunos.forEach(aluno => {
      const linha = [aluno.profile.name.fullName, aluno.profile.emailAddress];
      let presencas = 0;

      listaTopicosOrdenada.forEach(topico => {
        let totalEnvios = 0;

        topico.atividades.forEach(atividade => {
          const entregas = Classroom.Courses.CourseWork.StudentSubmissions.list(turma.id, atividade.id).studentSubmissions || [];
          const entregaAluno = entregas.find(e => e.userId === aluno.userId);
          const enviado = entregaAluno && (entregaAluno.state === 'TURNED_IN' || entregaAluno.state === 'RETURNED');
          if (enviado) totalEnvios++;
        });

        let situacao = '';
        if (totalEnvios >= 3) {
          situacao = 'Presente';
          presencas++;
        } else if (totalEnvios > 0) {
          situacao = 'Ausente de Entrega';
        } else {
          situacao = 'Falta';
        }

        linha.push(situacao);
      });

      const frequencia = totalTopicos > 0 ? Math.round((presencas / totalTopicos) * 100) : 0;
      linha.push(frequencia);
      aba.appendRow(linha);
    });
  });
}
