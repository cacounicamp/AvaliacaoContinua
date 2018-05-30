/**
 * A função verifica na planilha configurada se há mensagens que possam ser enviadas seguindo:
 * Para cada linha não vazia e válida (marcada com "ok" ou "not ok") (enviamos um e-mail de aviso à gestão caso uma mensagem seja marcada com algo diferente disso):
 *   Se a crítica é para o docente,
 *     Se a crítica foi marcada com "not ok",
 *       Enviamos de volta ao autor da crítica com uma justificativa
 *     Se a crítica foi marcada com "ok",
 *       Enviamos ao docente
 *   Se a crítica é para uma secretaria de graduação,
 *     Se a crítica foi marcada com "not ok",
 *       Enviamos de volta ao autor da crítica com uma justificativa
 *     Se a crítica foi marcada com "ok",
 *       Procuramos a linha com o e-mail da instituição
 *       Enviamos o e-mail se encontrado
 *       Vamos para a próxima crítica se não encontrado (e-mail de aviso é enviado à gestão)
 * Marcamos a crítica como enviada se chegamos aqui
 *
 *
 * Considerações razoáveis para organização e estética de quem recebe este e-mail:
 * (a) Mensagens para a secretaria de graduação devem possuir mesmo assunto no caso de tratar de uma mesma disciplina,
 * (b) Mensagens para a/o mesma(o) docente devem possuir mesmo assunto independente da/do aluna(o) que envia a crítica,
 * (c) Críticas inválidas (não enviadas ao docente/à secretaria) DEVEM ser justificadas e enviadas à/ao aluna(o).
 */
function verificarCriticas() {
  // Recebemos a planilha configurada pelo seu ID do Google (os números aleatórios que estão no endereço no navegador)
  const planilha = SpreadsheetApp.openById(<ID DA PLANILHA>);
  // Definimos a planilha como ativa (necessário por algum motivo esquisito)
  SpreadsheetApp.setActiveSpreadsheet(planilha);
  
  // Da planilha de respostas, pegamos todas as respostas
  const respostas = planilha.getSheets()[0].getRange("A2:K");
  // Da planilha de endereços de e-mail das unidades, pegamos todos os pares (unidade, e-mail)
  const enderecos_unidades = planilha.getSheets()[1].getRange("A2:B");
  // Da planilha de endereços de e-mail para lembrete, pegamos a quantidade de dias desde último e-mail (assumo que esse script é executado diariamente)
  const dias_ultimo_lembrete = planilha.getSheets()[2].getRange("B1").getCell(1, 1);
  // Da planilha de endereços de e-mail para lembrete, pegamos todos os pares (unidade, e-mail)
  const enderecos_lembretes = planilha.getSheets()[2].getRange("A3:B");
  
  // Adicionamos mais uma execução ao número de dias desde o último lembrete
  dias_ultimo_lembrete.setValue(dias_ultimo_lembrete.getValue() + 1);
  Logger.log("Dias desde o último lembrete: " + dias_ultimo_lembrete.getValue());
  
  /*
  * Como a numeração das células em uma Range é relativa, deixo aqui anotado como está nossa planilha (coluna na planilha e posição relativa):
  * (A, 01) Carimbo de data/hora
  * (B, 02) Endereço de email
  * (C, 03) Você quer enviar a crítica à/ao docente?
  * (D, 04) Qual o e-mail da/do docente?
  * (E, 05) Descreva o seu problema com uma possível solução
  * (F, 06) De qual instituto é a disciplina oferecida?
  * (G, 07) Código da disciplina com turma
  * (H, 08) Descreva o seu problema com uma possível solução
  * (I, 09) Crítica válida
  * (J, 10) Crítica enviada
  * (K, 11) Justificativa para crítica inválida
  */
  
  Logger.log("Verificando as mensagens da planilha... Os números de linha serão relativas ao intervalo A2:K.");
  // Variável que indica se o e-mail de aviso sobre o status da crítica ter sido preenchido incorretamente já foi enviada (para evitarmos enviar mais de uma vez a mesma mensagem em um só dia)
  var avisoPreenchimentoIncorreto = false;
  // Variável que indica se a unidade (instituição/faculdade) da crítica não possui e-mail preenchido na segunda folha da planilha (para evitarmos que a mensagem seja enviada mais de uma vez por dia)
  var avisoUnidadeDesconhecida = false;
  // Variável que indica se é necessário o envio de e-mail lembrete para quem quer aprovar as críticas
  var mensagensNaoVisualizadas = 0;
  
  // Variável para eu não precisar refazer a planilha (já que a primeira resposta vai para a linha 5, assumimos que o script será executado com pelo menos uma resposta)
  var encontrouPrimeiro = false;
  for (var linha = 1; ; linha++) {
    // Verificamos se a linha atual está vazia através da data
    if (respostas.getCell(linha, 1).isBlank()) {
      if(encontrouPrimeiro) {
        Logger.log("Terminamos de verificar todas as respostas na linha " + linha + ".");
        break;
      } else {
        continue;
      }
    }
    
    // Encontramos a primeira linha não vazia
    encontrouPrimeiro = true;
    
    // Pegamos a célula que indica se a crítica é válida (e pode ser enviada)
    var criticaValida = respostas.getCell(linha, 9);
    
    // Conferimos se a crítica já foi aprovada por algum membro do CACo
    if (criticaValida.isBlank()) {
      // Ignoramos mensagens não avaliadas
      Logger.log("\tLinha " + linha + " ignorada por não ter sido aprovada. Um e-mail talvez seja enviado.");
      mensagensNaoVisualizadas++;
      continue;
    }
    
    // Status da crítica deve ser "ok" ou "not ok" ou vazio, caso a crítica não tenha sido avaliada (esse último caso já foi excluído deste trecho do código)
    var statusCritica = criticaValida.getValue().toString().toLowerCase();
    
    // Verificamos se o status da crítica foi preenchido incorretamente (nem "ok", nem "not ok")
    if (statusCritica != "ok" && statusCritica != "not ok") {
      Logger.log("\tLinha " + linha + " não foi preenchida corretamente. Enviando e-mail de aviso!");
      // Damos um shame por dia (não mandamos vários e-mails de uma vez)
      if (!avisoPreenchimentoIncorreto) {
        mandarEmailAviso("Há uma mensagem não aprovada que não respeita o formato desejado.\n\nLeia os comentários antes de preencher a planilha.");
        avisoPreenchimentoIncorreto = true;
      }
      continue;
    }
    
    // Pegamos a célula que indica se a crítica já foi enviada, será re-editada ao final da iteração caso não tenha sido
    var celulaEnviada = respostas.getCell(linha, 10);
    
    // Ignoramos críticas já enviadas
    if (!celulaEnviada.isBlank()) {
      Logger.log("Linha " + linha + " foi pulada por já ter sido enviada");
      continue;
    }
    
    // Agora que temos uma linha válida, distinguimos entre mensagem para docente, para secretaria ou para o remetente (quem fez a crítica)
    if (respostas.getCell(linha, 3).getValue() == "Sim") { // Respondeu "sim" para "quer enviar à/ao docente?"
      if (statusCritica == "not ok") {
        Logger.log("\tLinha " + linha + " será enviada à/ao autor(a) da crítica à/ao docente por ser inválida");
        mandarEmailMensagemInvalida(
          // E-mail da/do autor(a)
          respostas.getCell(linha, 2).getValue(),
          // Assunto do e-mail
          "Avaliação à/ao docente sobre o andamento de semestre",
          // Mensagem crítica
          respostas.getCell(linha, 5).getValue(),
          // Mensagem justificativa
          respostas.getCell(linha, 11).getValue()
        );
      } else {
        Logger.log("\tLinha " + linha + " será enviada à/ao docente");
        mandarEmail(
          // E-mail da/do docente
          respostas.getCell(linha, 4).getValue(),
          // E-mail da/do aluna(o)
          respostas.getCell(linha, 2).getValue(),
          // Assunto
          "Avaliação à/ao docente sobre o andamento de semestre",
          // Mensagem-crítica
          respostas.getCell(linha, 5).getValue()
        );
      }
    } else {
      // Respondeu "não" à pergunta "quer enviar a crítica à/ao docente?"
      if (statusCritica == "not ok") {
        Logger.log("\tLinha " + linha + " será enviada à/ao autor(a) da crítica à unidade de ensino por ser inválida");
        mandarEmailMensagemInvalida(
          // E-mail da/do autor(a)
          respostas.getCell(linha, 2).getValue(),
          // Assunto do e-mail com código da disciplina
          "Avaliação contínua sobre " + respostas.getCell(linha, 7).getValue(),
          // Mensagem crítica
          respostas.getCell(linha, 8).getValue(),
          // Mensagem justificativa
          respostas.getCell(linha, 11).getValue()
        );
      } else {
        Logger.log("\tLinha " + linha + " será enviada à secretaria de graduação");
        // Pegamos a célula que contém o texto do nome da instituição
        var celulaInstituto = respostas.getCell(linha, 6);
        
        // Verificamos se a unidade não foi encontrada, para usar o break e não marcar a linha com OK
        var naoEncontrado = false;
        
        // Percorremos a lista de unidades até encontrar o e-mail da unidade desejada
        for (var unidade = 1; ; unidade++) {
          // Verificamos se terminamos de verificar todas as unidades
          if (enderecos_unidades.getCell(unidade, 1).isBlank()) {
            Logger.log("\tLinha " + linha + " contém uma unidade (instituição/faculdade) *NÃO* reconhecida. Enviando e-mail de aviso!");
            naoEncontrado = true;
            if (!avisoUnidadeDesconhecida) {
              mandarEmailAviso("Há uma unidade (instituição, faculdade) que não possui endereço de e-mail na planilha!\n\nPor favor, corrigam a planilha assim que possível.");
              avisoUnidadeDesconhecida = true;
            }
            break;
          }
          
          // Caso encontramos a instituição, enviamos um e-mail
          if (enderecos_unidades.getCell(unidade, 1).getValue() == celulaInstituto.getValue()) {
            mandarEmail(
              // E-mail da unidade
              enderecos_unidades.getCell(unidade, 2).getValue(),
              // E-mail da/do aluna(o)
              respostas.getCell(linha, 2).getValue(),
              // Assunto com o código da disciplina
              "Avaliação contínua sobre " + respostas.getCell(linha, 7).getValue(),
              // Mensagem-crítica
              respostas.getCell(linha, 8).getValue()
              );
            break;
          }
        }
        
        // Não marcamos com "ok", continuamos à próxima crítica
        if (naoEncontrado) {
          Logger.log("\tLinha " + linha + " não foi marcada como ok :(");
          continue;
        }
      }
    }
    
    // Como tudo deu certo, marcamos a célula como enviada
    celulaEnviada.setValue("OK");
    Logger.log("\tLinha " + linha + " marcada como OK :)");
  }

  // Verificamos se é necessário enviar e-mail lembrete  
  if (mensagensNaoVisualizadas > 1 && dias_ultimo_lembrete.getValue() >= 4) {
    // Fazemos a lista de interessad*s
    var enderecos_cco = "";
    
    for (var linha = 1; ; linha++) {
      // Para cada linha, verificamos se há e-mail
      var email = enderecos_lembretes.getCell(linha, 2);
      
      if (email.isBlank()) {
        break;
      }
      
      // Adicionamos à lista
      enderecos_cco = email.getValue() + "," + enderecos_cco;
    }
    
    // Enviamos o lembrete para tod*s *s interessad*s
    mandarEmailAviso("Olá,\n\nHá " + mensagensNaoVisualizadas + " críticas a serem avaliadas na planilha do sistema de avaliação contínua de semestre.", enderecos_cco);
    
    // Definimos como zero o número de dias desde o último lembrete
    dias_ultimo_lembrete.setValue(0);
  }
}


/**
 * Função enviará e-mail para a graduação do instituto com cópia oculta a quem reclamou.
 */
function mandarEmail(enderecos, enderecos_cco, assunto, mensagem) {
  const DESCRICAO = "Olá,\n\nEsta mensagem vem de uma requisição feita ao centro acadêmico da computação (CACo) para ser passada anônimamente e/ou facilmente através de um sistema automatizado. O intuito da mensagem é criticar de forma *construtiva* alguma característica de alguma disciplina e oferecer uma solução factível. Para nos certificar disso, algum membro da gestão do CACo aprovou esta mensagem antes dela ter sido enviada.\n\nÉ importante notar que verificamos com frequência alta nossos e-mails. Caso acredite que isso seja um engano ou queira responder à crítica, nos avise através do endereço caco@ic.unicamp.br respondendo este e-mail, por exemplo!\n\nA mensagem enviada pela(o) aluna(o) de computação ou não foi:\n\n"
  // Como mensagem não conterá assinatura, assinamos o e-mail com uma rápida descrição do sistema
  const ASSINATURA = "\n\n\n--\nCACo - Centro Acadêmico da Computação da Unicamp\n\nMensagem autorizada por membros ativos e enviada por um sistema automatizado.\n";
  
  // Enviamos o e-mail de caco.gestao@gmail.com (com um nome bonitinho) para 'enderecos', com cópia oculta para 'enderecos_cco'
  MailApp.sendEmail({
    name: "Centro Acadêmico da Computação",
    to: enderecos,
    cc: "caco@ic.unicamp.br",
    bcc: enderecos_cco,
    replyTo: "caco@ic.unicamp.br",
    subject: assunto,
    body: DESCRICAO + mensagem + ASSINATURA
  });
  
  Logger.log("\t\tE-mail enviado a \"" + enderecos + "\" com CC para o CACo e CCo para \"" + enderecos_cco + "\"");
}


function mandarEmailMensagemInvalida(enderecos, assunto, mensagem_critica, mensagem_justificativa) {
  const DESCRICAO = "Olá,\n\nO centro acadêmico da computação não enviou a sua mensagem à secretaria de graduação ou à/ao docente por não satisfazer os critérios que descrevemos no formulário.\n\nNo entanto, encorajamos que as críticas necessárias sejam enviadas, então enviaremos aqui a mensagem original (caso queira corrigí-la e re-enviá-la) junto com a justificativa pela qual deixamos de enviar sua crítica.\n\n"
  const INTRO_CRITICA = "Crítica original:\n\n";
  const INTRO_JUSTIFICATIVA = "\n\nJustificativa dos membros ativos para não enviar a crítica:\n\n";
  const ASSINATURA = "\n\n\n--\nCACo - Centro Acadêmico da Computação da Unicamp\n\nMensagem autorizada por membros ativos e enviada por um sistema automatizado.\n"
  
   MailApp.sendEmail({
    name: "Centro Acadêmico da Computação",
    to: enderecos,
    cc: "caco@ic.unicamp.br",
    replyTo: "caco@ic.unicamp.br",
    subject: assunto,
    body: DESCRICAO + INTRO_CRITICA + mensagem_critica + INTRO_JUSTIFICATIVA + mensagem_justificativa + ASSINATURA
  });
  
  Logger.log("\t\tE-mail enviado a \"" + enderecos + "\" com CC para o CACo justificando o não-envio da crítica"); 
}

/**
 * Função manda um e-mail de aviso à gestão do CACo por algum motivo (aviso, shame).
 */
function mandarEmailAviso(mensagem) {
  mandarEmailAviso(mensagem, "");
}

/**
 * Função manda um e-mail de aviso à gestão do CACo e uma cópia oculta à lista de endereços por algum motivo (aviso, shame).
 */
function mandarEmailAviso(mensagem, enderecos_cco) {
  // Como mensagem não conterá assinatura, assinamos o e-mail como o sistema
  const ASSINATURA = "\n\n--\nCACo - Centro Acadêmico da Computação da Unicamp\n\nMensagem enviada por um sistema automatizado.\n";
  
  // Enviamos o e-mail de caco.gestao@gmail.com (com um nome bonitinho) para 'enderecos', com cópia oculta para 'enderecos_cco'
  MailApp.sendEmail({
    name: "Centro Acadêmico da Computação",
    to: "caco@ic.unicamp.br",
    bcc: enderecos_cco,
    subject: "Sistema de avaliação contínua de andamento de semestre",
    body: mensagem + ASSINATURA
  });
  
  Logger.log("\t\tE-mail enviado a \"caco@ic.unicamp.br\" com avisos");
}

