/**
 * resumo-kpis.gs
 * 
 * Script para automatizar o resumo de KPIs em uma única aba no Google Sheets.
 * Todos os nomes sensíveis foram anonimizados para uso genérico e seguro.
 * 
 * Estrutura:
 * - Limpeza da aba de resumo
 * - Cálculo e consolidação dos dados nacionais
 * - Cálculo e consolidação dos dados por regionais genéricos
 * 
 * Autor: Guilherme Piovezan (exemplo)
 * Data: 2025
 */

/**
 * Função principal que executa o resumo básico e regional.
 */
function resumoBasicoERegional2025() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaResumo = ss.getSheetByName("Resumo_KPIs");
  if (!abaResumo) {
    SpreadsheetApp.getUi().alert("Aba Resumo_KPIs não encontrada.");
    return;
  }

  // Limpa um intervalo amplo para reiniciar os dados da aba de resumo
  abaResumo.getRange(1, 1, 1500, 50).clearContent();

  // Define o título principal na célula A1
  abaResumo.getRange(1, 1).setValue("Controle e Gestão de Quadro - Projeto Genérico");

  // Configuração genérica dos meses e suas respectivas abas
  var meses = [
    { nome: "Jan", aba: "Dados_202501" },
    { nome: "Fev", aba: "Dados_202502" },
    { nome: "Mar", aba: "Dados_202503" },
    { nome: "Abr", aba: "Dados_202504" },
    { nome: "Mai", aba: "Dados_202505" },
    { nome: "Jun", aba: "Dados_202506" },
    { nome: "Jul", aba: "Dados_202507" },
    { nome: "Ago", aba: "Dados_202508" },
    { nome: "Set", aba: "Dados_202509" },
    { nome: "Out", aba: "Dados_202510" },
    { nome: "Nov", aba: "Dados_202511" },
    { nome: "Dez", aba: "Dados_202512" }
  ];

  // Tipos de movimentos monitorados
  var movimentos = [
    "Desligamento",
    "Mudança de HUB",
    "Licença (Saúde/Maternidade)",
    "Férias",
    "Retorno (Afastamento)",
    "Entradas"
  ];

  // Mapeamento genérico de cargos para quadros
  var cargosMapeados = {
    "ANALYST JR": "Executivo 2",
    "ANALYST SSR": "Executivo 3",
    "SALES EXECUTIVE": "Executivo 1",
    "TEAM LEADER": "Team Leader"
  };

  // Lista de quadros considerados
  var quadros = ["Executivo 1", "Executivo 2", "Executivo 3", "Team Leader"];

  // Lista genérica de regionais usadas no relatório
  var regionaisPadrao = listarRegionais();

  /**
   * Normaliza variações de nomes de regionais para nomes padronizados.
   * Todos os nomes foram anonimizados para uso genérico.
   */
  function normalizarRegional(nome, idxMes) {
    if (!nome) return "";
    nome = nome.trim().toUpperCase();

    // Mapeamento de variações para nomes genéricos
    const mapa = {
      "REG A": "Regional 1",
      "REG AA": "Regional 1",
      "REG B": "Regional 2",
      "REG BB": "Regional 2",
      "REG C": "Regional 3",
      "REG CC": "Regional 3",
      "REG D": "Regional 4",
      "REG DD": "Regional 4",
      "REG E": "Regional 5",
      "REG EE": "Regional 5",
      "REG F": "Regional 6",
      "REG FF": "Regional 6",
      "REG G": "Regional 7",
      "REG GG": "Regional 7",
      "REG H": "Regional 8",
      "REG HH": "Regional 8",
      "REG I": "Regional 9",
      "REG II": "Regional 9",
      "REG J": "Regional 10",
      "REG JJ": "Regional 10",
      // Regra específica para Regional 10 em meses maiores que Fevereiro (idxMes >= 2)
      "CENTRO-NORTE": idxMes >= 2 ? "Regional 7" : "Regional 10"
    };

    return mapa[nome] || nome;
  }

  /**
   * Retorna lista genérica de regionais para o projeto.
   */
  function listarRegionais() {
    return [
      "Regional 1", "Regional 2", "Regional 3", "Regional 4", "Regional 5",
      "Regional 6", "Regional 7", "Regional 8", "Regional 9", "Regional 10"
    ];
  }

  // --- 1. RESUMO NACIONAL ---
  var resultados = [ ["Movimento"].concat(meses.map(m => m.nome)) ];
  var dadosMov = {};
  movimentos.forEach(mv => { dadosMov[mv] = Array(meses.length).fill(0); });
  var ativosQuadros = {};
  quadros.forEach(q => { ativosQuadros[q] = Array(meses.length).fill(0); });

  meses.forEach(function(mes, idx) {
    var abaMes = ss.getSheetByName(mes.aba);
    if (!abaMes) return;

    var dados = abaMes.getDataRange().getValues();
    var head = dados[0];

    var idxStatus = head.indexOf("STATUS");
    var idxAdmissao = head.indexOf("DATA_ADMISSAO");
    var idxMesRef = head.indexOf("MES_REF");
    var idxCargoQuadro = head.indexOf("CARGO_QUADRO");
    var idxCargo = head.indexOf("CARGO");
    var idxHub = head.indexOf("HUB");
    var idxRegional = head.indexOf("REGIONAL");
    var idxUserId = head.indexOf("USER_ID");

    // Mapas para HUB/trocas entre meses
    var mapaHubAntigo = {};
    var mapaStatusAntigo = {};

    if (idxUserId !== -1 && idxHub !== -1 && idxStatus !== -1) {
      var mesAntAba = (idx === 0) ? ss.getSheetByName("Dados_202412") : ss.getSheetByName(meses[idx-1].aba);
      if (mesAntAba) {
        var dadosAntigos = mesAntAba.getDataRange().getValues();
        var headAnt = dadosAntigos[0];
        var idxIdAntigo = headAnt.indexOf("USER_ID");
        var idxHubAntigo = headAnt.indexOf("HUB");
        var idxStatusAnt = headAnt.indexOf("STATUS");

        if (idxIdAntigo !== -1 && idxHubAntigo !== -1 && idxStatusAnt !== -1) {
          for (var k = 1; k < dadosAntigos.length; k++) {
            var idAnt = (dadosAntigos[k][idxIdAntigo] || "").toString().trim();
            var hubAnt = (dadosAntigos[k][idxHubAntigo] || "").toString().trim();
            var statusAnt = (dadosAntigos[k][idxStatusAnt] || "").toString().trim();
            mapaHubAntigo[idAnt] = hubAnt;
            mapaStatusAntigo[idAnt] = statusAnt;
          }
        }
      }
    }

    // Processa dados linha a linha
    dados.slice(1).forEach(function(row) {
      var status = row[idxStatus];
      var userId = idxUserId !== -1 ? (row[idxUserId] || "").toString().trim() : "";
      var cargoNome = idxCargoQuadro !== -1 ? row[idxCargoQuadro] : (idxCargo !== -1 ? row[idxCargo] : null);

      // Conta desligamentos
      if (status && status.indexOf("Desligamento") !== -1) dadosMov["Desligamento"][idx]++;

      // Conta férias
      if (status === "Férias") dadosMov["Férias"][idx]++;

      // Conta licenças (doença/maternidade)
      if (status === "Inativo - Licença Doença" || status === "Inativo - Licença Maternidade") dadosMov["Licença (Saúde/Maternidade)"][idx]++;

      // Conta entradas (novas admissões no mês)
      var dtAdmissao = row[idxAdmissao];
      var mesRef = row[idxMesRef];
      if (dtAdmissao && mesRef) {
        var data = new Date(dtAdmissao);
        var mesAdmissao = (data.getMonth() + 1).toString().padStart(2, '0');
        var anoAdmissao = data.getFullYear().toString();
        mesRef = mesRef.toString();
        if (mesRef.slice(4) === mesAdmissao && mesRef.slice(0,4) === anoAdmissao) dadosMov["Entradas"][idx]++;
      }

      // Conta mudanças de HUB
      var hubAtual = idxHub !== -1 ? row[idxHub] : "";
      if (userId && mapaHubAntigo[userId] && hubAtual && status && status.indexOf("Ativo") !== -1) {
        if (mapaHubAntigo[userId] !== hubAtual) {
          dadosMov["Mudança de HUB"][idx]++;
        }
      }

      // Conta retornos de afastamento
      if (userId && (mapaStatusAntigo[userId] === "Inativo - Licença Doença" || mapaStatusAntigo[userId] === "Inativo - Licença Maternidade") && status === "Ativo") {
        dadosMov["Retorno (Afastamento)"][idx]++;
      }

      // Conta ativos por quadro/cargo
      if (status && status.indexOf("Ativo") !== -1 && cargoNome) {
        var cargoPadrao = String(cargoNome).trim().toUpperCase();
        var quadro = cargosMapeados[cargoPadrao];
        if (quadro) ativosQuadros[quadro][idx]++;
      }
    });
  });

  // Monta bloco Nacional - Movimentos
  movimentos.forEach(function(mv) {
    resultados.push([mv].concat(dadosMov[mv]));
  });

  // Linha Total - soma de movimentos por mês
  var totalLinha = ["TOTAL"];
  for (var col = 0; col < meses.length; col++) {
    var soma = 0;
    movimentos.forEach(function(mv) {
      soma += dadosMov[mv][col];
    });
    totalLinha.push(soma);
  }
  resultados.push(totalLinha);

  // Linha Total Ativo - total de ativos menos desligamentos, licenças e férias
  var totalAtivosArray = ["Total Ativo"];
  for (var col = 0; col < meses.length; col++) {
    var totalAtivosMes = 0;
    quadros.forEach(function(q) { totalAtivosMes += ativosQuadros[q][col]; });
    var totalAtivoCalc = totalAtivosMes - dadosMov["Desligamento"][col] - dadosMov["Licença (Saúde/Maternidade)"][col] - dadosMov["Férias"][col];
    totalAtivosArray.push(totalAtivoCalc);
  }
  resultados.push(totalAtivosArray);

  // Monta bloco Nacional - Quadros/Cargos
  var resultadosQuadros = [["Níveis e Cargos"].concat(meses.map(m => m.nome))];
  quadros.forEach(function(q) {
    resultadosQuadros.push([q].concat(ativosQuadros[q]));
  });
  var totalAtivos = ["Total Ativos"];
  for (var col = 0; col < meses.length; col++) {
    var soma = 0;
    quadros.forEach(function(q) { soma += ativosQuadros[q][col]; });
    totalAtivos.push(soma);
  }
  resultadosQuadros.push(totalAtivos);

  // Escreve blocos nacionais na aba resumo
  abaResumo.getRange(2, 1, resultados.length, resultados[0].length).setValues(resultados);
  var linhaQuadrosNac = 2 + resultados.length + 1;
  abaResumo.getRange(linhaQuadrosNac, 1, resultadosQuadros.length, resultadosQuadros[0].length).setValues(resultadosQuadros);

  var outputLin = linhaQuadrosNac + resultadosQuadros.length + 2;

  // --- 2. RESUMO POR REGIONAL ---
  regionaisPadrao.forEach(function(regionalAtiva) {
    // Filtra condições específicas para regional "Regional 10"
    if (regionalAtiva == "Regional 10") {
      var temDados = false;
      for (var idxMes = 0; idxMes <= 1; idxMes++) {
        var abaMes = ss.getSheetByName(meses[idxMes].aba);
        if (!abaMes) continue;
        var dados = abaMes.getDataRange().getValues();
        var head = dados[0];
        var idxRegional = head.indexOf("REGIONAL");
        if (idxRegional === -1) continue;
        var count = 0;
        dados.slice(1).forEach(function(row) {
          var regionalNome = row[idxRegional];
          var regionalNorm = normalizarRegional(regionalNome, idxMes);
          if (regionalNorm === "Regional 10")
            count++;
        });
        if (count > 0) temDados = true;
      }
      if (!temDados) return;
    }
    // Filtra condições para regional "Regional 7" a partir de Março
    if (regionalAtiva == "Regional 7") {
      var temAlguem = false;
      for (var idxMes = 2; idxMes < meses.length; idxMes++) {
        var abaMes = ss.getSheetByName(meses[idxMes].aba);
        if (!abaMes) continue;
        var dados = abaMes.getDataRange().getValues();
        var head = dados[0];
        var idxRegional = head.indexOf("REGIONAL");
        if (idxRegional === -1) continue;
        var count = 0;
        dados.slice(1).forEach(function(row) {
          var regionalNome = row[idxRegional];
          var regionalNorm = normalizarRegional(regionalNome, idxMes);
          if (regionalNorm === "Regional 7")
            count++;
        });
        if (count > 0) temAlguem = true;
      }
      if (!temAlguem) return;
    }

    abaResumo.getRange(outputLin, 1).setValue("Regional: " + regionalAtiva);
    outputLin++;

    // Inicializa dados e quadros para a regional atual
    var dadosMov = {};
    movimentos.forEach(function(mv) { dadosMov[mv] = Array(meses.length).fill(0); });
    var ativosQuadros = {};
    quadros.forEach(function(q) { ativosQuadros[q] = Array(meses.length).fill(0); });

    meses.forEach(function(mes, idx) {
      var abaMes = ss.getSheetByName(mes.aba);
      if (!abaMes) return;
      var dados = abaMes.getDataRange().getValues();
      var head = dados[0];
      var idxStatus = head.indexOf("STATUS");
      var idxAdmissao = head.indexOf("DATA_ADMISSAO");
      var idxMesRef = head.indexOf("MES_REF");
      var idxCargoQuadro = head.indexOf("CARGO_QUADRO");
      var idxCargo = head.indexOf("CARGO");
      var idxHub = head.indexOf("HUB");
      var idxUserId = head.indexOf("USER_ID");
      var idxRegional = head.indexOf("REGIONAL");
      if (idxStatus === -1) return;

      var mapaHubAntigo = {}, mapaStatusAntigo = {};
      if (idxUserId !== -1 && idxHub !== -1 && idxStatus !== -1) {
        var mesAntAba = (idx === 0) ? ss.getSheetByName("Dados_202412") : ss.getSheetByName(meses[idx - 1].aba);
        if (mesAntAba) {
          var dadosAntigos = mesAntAba.getDataRange().getValues();
          var headAnt = dadosAntigos[0];
          var idxIdAntigo = headAnt.indexOf("USER_ID");
          var idxHubAntigo = headAnt.indexOf("HUB");
          var idxStatusAnt = headAnt.indexOf("STATUS");
          var idxRegionalAnt = headAnt.indexOf("REGIONAL");
          if (idxIdAntigo !== -1 && idxHubAntigo !== -1 && idxStatusAnt !== -1) {
            for (var k = 1; k < dadosAntigos.length; k++) {
              var idAnt = (dadosAntigos[k][idxIdAntigo] || "").toString().trim();
              var hubAnt = (dadosAntigos[k][idxHubAntigo] || "").toString().trim();
              var statusAnt = (dadosAntigos[k][idxStatusAnt] || "").toString().trim();
              var regionalAnt = normalizarRegional(dadosAntigos[k][idxRegionalAnt], idx - 1);
              mapaHubAntigo[idAnt] = { hub: hubAnt, status: statusAnt, regional: regionalAnt };
              mapaStatusAntigo[idAnt] = statusAnt;
            }
          }
        }
      }

      dados.slice(1).forEach(function(row) {
        var regionalNome = idxRegional !== -1 ? row[idxRegional] : "";
        var status = row[idxStatus];
        var userId = idxUserId !== -1 ? (row[idxUserId] || "").toString().trim() : "";
        var cargoNome = idxCargoQuadro !== -1 ? row[idxCargoQuadro] : (idxCargo !== -1 ? row[idxCargo] : null);
        var regionalNorm = normalizarRegional(regionalNome, idx);
        if (regionalNorm !== regionalAtiva) return;

        // Conta desligamentos
        if (status && status.indexOf("Desligamento") !== -1) dadosMov["Desligamento"][idx]++;
        // Conta férias
        if (status === "Férias") dadosMov["Férias"][idx]++;
        // Conta licenças
        if (status === "Inativo - Licença Doença" || status === "Inativo - Lic
