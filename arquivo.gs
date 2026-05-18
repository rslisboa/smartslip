/****************************************************
 * SMARTSLIP - API para Google AI Studio / Gemini
 * 
 * Funções:
 * - consultar_info_loja
 * - salvar_dados_comprovante_planilha
 *
 * Não interfere no fluxo BaseClara atual.
 ****************************************************/

const SMARTSLIP_TOKEN_PROPERTY = "SMARTSLIP_TOKEN";

// Planilha antiga onde está a aba Info_limites.
// ATENÇÃO: coloque aqui o ID exato da planilha Capta_Clara/Info_limites.
const SMARTSLIP_INFO_LIMITES_SPREADSHEET_ID = "1_XW0IqbYjiCPpqtwdEi1xPxDlIP2MSkMrLGbeinLIeI";

// Nova planilha SmartSlip, onde ficam BASE_SMARTSLIP e SMARTSLIP_FILA.
const SMARTSLIP_DB_SPREADSHEET_ID = "145m9bu6mg6n4Oeh0QPDaiKZIE6MhZv0vux0W1xR9dpY";

const SMARTSLIP_ABA_INFO_LIMITES = "Info_limites";
const SMARTSLIP_ABA_DESTINO = "BASE_SMARTSLIP";
const SMARTSLIP_ABA_FILA = "SMARTSLIP_FILA";
const SMARTSLIP_ABA_VEKTOR_EMAILS = "VEKTOR_EMAILS";
const SMARTSLIP_VEKTOR_EMAILS_SPREADSHEET_ID = SMARTSLIP_INFO_LIMITES_SPREADSHEET_ID;
const SMARTSLIP_ABA_USUARIOS = "SMARTSLIP_USUARIOS";

function smartSlipGetUsuarioAtual() {
  const email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();

  const primeiroNome = smartSlipExtrairPrimeiroNome(email);

    let lojaPadrao = "";
    let empresaPadrao = "";

    let acessoUsuario = {
      role: "Usuário",
      is_admin: false,
      is_analista_pro: false,
      can_comp_hub: false
    };

    try {
      acessoUsuario = smartSlipGetAcessoUsuario(email);
    } catch (errAcesso) {
      Logger.log("Erro ao verificar acesso do usuário: " + String(errAcesso && errAcesso.message ? errAcesso.message : errAcesso));
    }

  try {
    const prefs = smartSlipObterPreferenciasUsuario(email);

    lojaPadrao = prefs.loja_padrao || "";
    empresaPadrao = prefs.empresa_padrao || "";

    if (!lojaPadrao) {
      lojaPadrao = smartSlipObterUltimaLojaUsuario(email);

      if (lojaPadrao) {
        const infoLoja = smartSlipConsultarInfoLoja_(lojaPadrao);

        if (infoLoja.ok) {
          empresaPadrao = infoLoja.empresa;
        }
      }
    }

  } catch (errLoja) {
    Logger.log("Erro ao obter loja padrão do usuário: " + String(errLoja && errLoja.message ? errLoja.message : errLoja));
    lojaPadrao = "";
    empresaPadrao = "";
  }

    return {
    email: email,
    primeiro_nome: primeiroNome,
    perfil: acessoUsuario.role || "Usuário",
    is_admin: acessoUsuario.is_admin === true,
    is_analista_pro: acessoUsuario.is_analista_pro === true,
    can_comp_hub: acessoUsuario.can_comp_hub === true,
    loja_padrao: lojaPadrao,
    empresa_padrao: empresaPadrao
  };
}

function smartSlipExtrairPrimeiroNome(email) {
  if (!email) return "Usuário";

  const parte = String(email).split("@")[0] || "";
  const primeiro = parte.split(/[._-]/)[0] || "Usuário";

  return primeiro.charAt(0).toUpperCase() + primeiro.slice(1).toLowerCase();
}

function smartSlipNormalizarPerfilSmartSlip(valor) {
  return String(valor || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function smartSlipGetAcessoUsuario(email) {
  try {
    email = String(email || "").trim().toLowerCase();

    const acessoPadrao = {
      role: "Usuário",
      role_norm: "USUARIO",
      is_admin: false,
      is_analista_pro: false,
      can_comp_hub: false
    };

    if (!email) {
      return acessoPadrao;
    }

    const ss = SpreadsheetApp.openById(SMARTSLIP_VEKTOR_EMAILS_SPREADSHEET_ID);
    const sh = ss.getSheetByName(SMARTSLIP_ABA_VEKTOR_EMAILS);

    if (!sh) {
      Logger.log("SmartSlip Acesso: aba VEKTOR_EMAILS não encontrada.");
      return acessoPadrao;
    }

    const values = sh.getDataRange().getValues();

    if (values.length < 2) {
      return acessoPadrao;
    }

    const headers = values[0].map(function(h) {
      return String(h || "").trim();
    });

    const idxEmail = smartSlipEncontrarIndiceHeader_(headers, [
      "Email",
      "EMAIL",
      "E-mail",
      "E_MAIL"
    ]);

    const idxRole = smartSlipEncontrarIndiceHeader_(headers, [
      "ROLE",
      "Role",
      "Perfil",
      "PERFIL"
    ]);

    const idxAtivo = smartSlipEncontrarIndiceHeader_(headers, [
      "ATIVO",
      "Ativo",
      "Status",
      "STATUS"
    ]);

    if (idxEmail < 0 || idxRole < 0) {
      Logger.log("SmartSlip Acesso: colunas Email/ROLE não encontradas na VEKTOR_EMAILS.");
      Logger.log(JSON.stringify(headers));
      return acessoPadrao;
    }

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const emailRow = String(row[idxEmail] || "").trim().toLowerCase();
      const roleRaw = String(row[idxRole] || "").trim();
      const roleNorm = smartSlipNormalizarPerfilSmartSlip(roleRaw);

      const ativoRaw = idxAtivo >= 0
        ? smartSlipNormalizarPerfilSmartSlip(row[idxAtivo])
        : "SIM";

      const estaAtivo = ![
        "NAO",
        "FALSE",
        "0",
        "INATIVO",
        "N"
      ].includes(ativoRaw);

      if (emailRow !== email || !estaAtivo) {
        continue;
      }

      const isAdmin =
        roleNorm.includes("ADMINISTRADOR") ||
        roleNorm.includes("ADMIN");

      const isAnalistaPro =
        roleNorm.includes("ANALISTA PRO") ||
        roleNorm.includes("ANALISTA_PRO") ||
        roleNorm.includes("ANALISTA-PRO") ||
        (roleNorm.includes("ANALISTA") && roleNorm.includes("PRO"));

      return {
        role: roleRaw || "Usuário",
        role_norm: roleNorm || "USUARIO",
        is_admin: isAdmin,
        is_analista_pro: isAnalistaPro,
        can_comp_hub: isAdmin || isAnalistaPro
      };
    }

    return acessoPadrao;

  } catch (err) {
    Logger.log("Erro smartSlipGetAcessoUsuario: " + String(err && err.message ? err.message : err));

    return {
      role: "Usuário",
      role_norm: "USUARIO",
      is_admin: false,
      is_analista_pro: false,
      can_comp_hub: false
    };
  }
}

function smartSlipUsuarioEhAdmin(email) {
  return smartSlipGetAcessoUsuario(email).is_admin === true;
}


function smartSlipGetToken() {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty(SMARTSLIP_TOKEN_PROPERTY);

  if (!token) {
    throw new Error("SMARTSLIP_TOKEN não configurado nas propriedades do script.");
  }

  return token;
}

/**
 * Web App usado pelo Gemini / AI Studio.
 *
 * Recebe:
 * {
 *   token: "...",
 *   acao: "consultar_info_loja" ou "salvar_dados_comprovante_planilha",
 *   args: {...}
 * }
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");

    const tokenEsperado = smartSlipGetToken();

    if (body.token !== tokenEsperado) {
      return smartSlipJson_({
        ok: false,
        erro: "Token inválido."
      });
    }

    const acao = String(body.acao || "").trim();
    const args = body.args || {};

    if (acao === "consultar_info_loja") {
      return smartSlipJson_(smartSlipConsultarInfoLoja_(args.loja_informada));
    }

    if (acao === "salvar_dados_comprovante_planilha") {
      return smartSlipJson_(smartSlipSalvarDadosComprovantePlanilha_(args));
    }

    return smartSlipJson_({
      ok: false,
      erro: "Ação não reconhecida: " + acao
    });

  } catch (err) {
    return smartSlipJson_({
      ok: false,
      erro: smartSlipSafeErr_(err)
    });
  }
}

/**
 * Consulta a aba Info_limites para localizar a loja e empresa.
 */
function smartSlipConsultarInfoLoja_(lojaInformada) {
  if (!lojaInformada) {
    return {
      ok: false,
      erro: "loja_informada não recebida."
    };
  }

  const lojaDigits = String(lojaInformada).replace(/\D/g, "");

  if (!lojaDigits) {
    return {
      ok: false,
      erro: "Número da loja inválido."
    };
  }

  const lojaSemZeros = String(Number(lojaDigits));
  const loja4 = lojaDigits.padStart(4, "0").slice(-4);

  const ss = SpreadsheetApp.openById(SMARTSLIP_INFO_LIMITES_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SMARTSLIP_ABA_INFO_LIMITES);

  if (!sh) {
    throw new Error("Aba '" + SMARTSLIP_ABA_INFO_LIMITES + "' não encontrada.");
  }

  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error("Aba '" + SMARTSLIP_ABA_INFO_LIMITES + "' está vazia.");
  }

  const headers = values[0].map(function(h) {
    return String(h || "").trim();
  });

  const idxNormLoja = smartSlipEncontrarIndiceHeader_(headers, [
    "Norm_Loja",
    "NORM_LOJA",
    "Norm Loja",
    "Loja",
    "LOJA"
  ]);

  let idxEmpresa = smartSlipEncontrarIndiceHeader_(headers, [
    "EMPRESA0",
    "EMPRESA",
    "Empresa"
  ]);

  // Fallback: você comentou que EMPRESA0 está na coluna G.
  // Coluna G = índice 6.
  if (idxEmpresa < 0) {
    idxEmpresa = 6;
  }

  if (idxNormLoja < 0) {
    throw new Error("Coluna Norm_Loja não encontrada na aba Info_limites.");
  }

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const normLojaRaw = row[idxNormLoja];
    const normLojaDigits = String(normLojaRaw || "").replace(/\D/g, "");

    if (!normLojaDigits) {
      continue;
    }

    const normLojaSemZeros = String(Number(normLojaDigits));
    const normLoja4 = normLojaDigits.padStart(4, "0").slice(-4);

    const bateu =
      normLojaSemZeros === lojaSemZeros ||
      normLoja4 === loja4;

    if (bateu) {
      const empresaRaw = String(row[idxEmpresa] || "").trim();
      const empresaNorm = smartSlipNormalizarEmpresa_(empresaRaw);

      return {
        ok: true,
        loja_informada: String(lojaInformada),
        loja_normalizada: normLoja4,
        empresa: empresaNorm,
        empresa_original: empresaRaw,
        linha_info_limites: i + 1
      };
    }
  }

  return {
    ok: false,
    erro: "Loja não encontrada na Info_limites.",
    loja_informada: String(lojaInformada),
    loja_normalizada_tentativa: loja4
  };
}

/**
 * Salva o comprovante validado na aba BASE_SMARTSLIP.
 */
function smartSlipSalvarDadosComprovantePlanilha_(dados) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

   const ss = SpreadsheetApp.openById(SMARTSLIP_DB_SPREADSHEET_ID);
    let sh = ss.getSheetByName(SMARTSLIP_ABA_DESTINO);

    if (!sh) {
      sh = ss.insertSheet(SMARTSLIP_ABA_DESTINO);
      smartSlipCriarCabecalho_(sh);
    }

    smartSlipGarantirCabecalho_(sh);
    smartSlipValidarDadosComprovante_(dados);

      sh.appendRow([
      new Date(),
      dados.id_comprovante || "",
      dados.hash_comprovante || "",
      dados.protocolo || "",
      dados.indice_comprovante || "",
      dados.loja || "",
      dados.empresa || "",
      dados.tipo_documento || "",
      dados.data_deposito || "",
      dados.valor_deposito || 0,
      dados.banco || "",
      dados.data_movimento || "",
      dados.houve_retirada === true ? "Sim" : "Não",
      dados.valor_retirada || 0,
      dados.motivo_retirada || "",
      dados.data_geracao_documento || "",
      dados.codigo_autenticacao || "",
      dados.link_comprovante || "",
      dados.status_processamento || "",
      (dados.pendencias || []).join(" | "),
      (dados.divergencias || []).join(" | "),
      dados.confianca_geral || "",
      JSON.stringify(dados)
    ]);

    return {
      ok: true,
      mensagem: "Dados salvos com sucesso na planilha.",
      aba_destino: SMARTSLIP_ABA_DESTINO
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/**
 * Cria cabeçalho da aba BASE_SMARTSLIP.
 */
function smartSlipCriarCabecalho_(sh) {
  sh.appendRow([
    "Data Registro",
    "ID Comprovante",
    "Hash Comprovante",
    "Protocolo",
    "Índice Comprovante",
    "Loja",
    "Empresa",
    "Tipo Documento",
    "Data Depósito",
    "Valor Depósito",
    "Banco",
    "Data Movimento",
    "Houve Retirada",
    "Valor Retirada",
    "Motivo Retirada",
    "Data Geração Documento",
    "Código Autenticação",
    "Link Comprovante",
    "Status Processamento",
    "Pendências",
    "Divergências",
    "Confiança Geral",
    "JSON Original"
  ]);
}

/**
 * Garante que a aba tem cabeçalho.
 */
function smartSlipGarantirCabecalho_(sh) {
  if (sh.getLastRow() === 0) {
    smartSlipCriarCabecalho_(sh);
    return;
  }

  const primeiraCelula = String(sh.getRange(1, 1).getValue() || "").trim();

  if (!primeiraCelula) {
    smartSlipCriarCabecalho_(sh);
  }
}

/**
 * Valida campos obrigatórios antes de salvar.
 */
function smartSlipValidarDadosComprovante_(dados) {
  const obrigatorios = [
    "loja",
    "empresa",
    "tipo_documento",
    "data_deposito",
    "valor_deposito",
    "data_movimento",
    "houve_retirada",
    "valor_retirada",
    "motivo_retirada",
    "data_geracao_documento",
    "status_processamento",
    "pendencias",
    "divergencias",
    "confianca_geral"
  ];

  const faltantes = [];

  obrigatorios.forEach(function(campo) {
    const valor = dados[campo];

    const vazio =
      valor === undefined ||
      valor === null ||
      valor === "";

    if (vazio) {
      faltantes.push(campo);
    }
  });

  if (!Array.isArray(dados.pendencias)) {
    faltantes.push("pendencias deve ser array");
  }

  if (!Array.isArray(dados.divergencias)) {
    faltantes.push("divergencias deve ser array");
  }

  if (faltantes.length) {
    throw new Error("Campos obrigatórios ausentes/inválidos: " + faltantes.join(", "));
  }
}

/**
 * Localiza índice de cabeçalho por nomes possíveis.
 */
function smartSlipEncontrarIndiceHeader_(headers, nomesPossiveis) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").trim().toLowerCase();

    for (let j = 0; j < nomesPossiveis.length; j++) {
      const alvo = String(nomesPossiveis[j] || "").trim().toLowerCase();

      if (h === alvo) {
        return i;
      }
    }
  }

  return -1;
}

/**
 * Normaliza nome da empresa.
 */
function smartSlipNormalizarEmpresa_(empresaRaw) {
  const txt = String(empresaRaw || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  if (txt.includes("FISIA") || txt.includes("NIKE")) {
    return "Fisia";
  }

  if (txt.includes("CENTAURO")) {
    return "Centauro";
  }

  return "Nao identificado";
}

/**
 * Resposta JSON padrão.
 */
function smartSlipJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Erro seguro.
 */
function smartSlipSafeErr_(err) {
  try {
    if (!err) return "erro desconhecido";
    if (err.message) return err.message;
    return String(err);
  } catch (e) {
    return "erro ao serializar erro";
  }
}

/**
 * Teste manual da consulta de loja.
 * Rode pelo Apps Script antes de conectar no AI Studio.
 */
function smartSlipTesteConsultarLoja() {
  const r = smartSlipConsultarInfoLoja_("241");
  Logger.log(JSON.stringify(r, null, 2));
}

/**
 * Teste manual de salvamento.
 * Rode pelo Apps Script antes de conectar no AI Studio.
 */
function smartSlipTesteSalvar() {
  const r = smartSlipSalvarDadosComprovantePlanilha_({
    loja: "0241",
    empresa: "Centauro",
    tipo_documento: "deposito_caixa_eletronico",
    data_deposito: "2026-05-08",
    valor_deposito: 37200.46,
    banco: "Itaú",
    data_movimento: "2026-05-07",
    houve_retirada: false,
    valor_retirada: 0,
    motivo_retirada: "Não houve retirada",
    data_geracao_documento: "2026-05-08",
    codigo_autenticacao: "TESTE123",
    link_comprovante: "",
    status_processamento: "PRONTO_PARA_SALVAR",
    pendencias: ["sem_pendencias"],
    divergencias: ["sem_divergencias"],
    confianca_geral: 0.95
  });

  Logger.log(JSON.stringify(r, null, 2));
}

function smartSlipTesteDoPostFake() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        token: smartSlipGetToken(),
        acao: "consultar_info_loja",
        args: {
          loja_informada: "241"
        }
      })
    }
  };

  const r = doPost(fakeEvent);
  Logger.log(r.getContent());
}

function smartSlipTesteGeminiPing() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    throw new Error("GEMINI_API_KEY não configurada nas propriedades do script.");
  }

  const model = "gemini-3-flash-preview";

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    model +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);

  const payload = {
    contents: [
      {
        role: "user",
        parts: [
          {
            text: "Responda somente este JSON: {\"ok\":true,\"mensagem\":\"Gemini funcionando no Apps Script\"}"
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      response_mime_type: "application/json"
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode = resp.getResponseCode();
  const txt = resp.getContentText();

  Logger.log("HTTP STATUS: " + statusCode);
  Logger.log(txt);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("Erro Gemini API HTTP " + statusCode + ": " + txt);
  }

  const json = JSON.parse(txt);
  const respostaTexto = json.candidates[0].content.parts[0].text;

  Logger.log("RESPOSTA GEMINI:");
  Logger.log(respostaTexto);

  return respostaTexto;
}

function smartSlipTesteGeminiComArquivoDrive() {
  const fileId = "1DtSSYlpbdpjDubLd05wue-XKgaosjcRN";

  const file = DriveApp.getFileById(fileId);
  const linkComprovante = file.getUrl();
  const blob = file.getBlob();

  const base64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

const respostasUsuario = {
  loja: "0255",
  data_movimento_inicio: "01/05/2026",
  data_movimento_fim: "04/05/2026",
  houve_retirada: false,
  valor_retirada: 0,
  motivo_retirada: "Não houve retirada",
  mais_comprovantes: false
};

  const resultado = smartSlipChamarGemini(base64, mimeType, respostasUsuario);

  Logger.log(JSON.stringify(resultado, null, 2));

  if (resultado.dados_comprovante && resultado.dados_comprovante.loja) {
    const infoLoja = smartSlipConsultarInfoLoja_(resultado.dados_comprovante.loja);

    Logger.log("INFO LOJA:");
    Logger.log(JSON.stringify(infoLoja, null, 2));

    if (infoLoja.ok) {
      resultado.dados_comprovante.loja = infoLoja.loja_normalizada;
      resultado.dados_comprovante.empresa = infoLoja.empresa;
    }
  }

  Logger.log("RESULTADO FINAL:");
  Logger.log(JSON.stringify(resultado, null, 2));
}

function smartSlipChamarGemini(base64Arquivo, mimeType, respostasUsuario) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    throw new Error("GEMINI_API_KEY não configurada nas propriedades do script.");
  }

  const model = "gemini-2.5-flash";

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    model +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);

  const prompt = `
Você é o Agente SmartSlip, especialista em leitura, extração, validação e interpretação semântica de comprovantes de depósito numerário para lojas Centauro e Fisia/Nike.

Objetivo:
Ler o comprovante anexado, extrair dados essenciais e retornar SOMENTE JSON válido.

Regras obrigatórias:
- Nunca invente número de loja.
- Nunca use CNPJ, código de barras, número de documento, autenticação, agência, conta ou identificador bancário como número de loja.
- Se a loja não estiver clara no comprovante e também não estiver em respostasUsuario.loja, retorne loja como null.
- Se respostasUsuario.loja estiver preenchido, use esse valor como loja informada pelo usuário.
- A empresa só pode ser "Centauro", "Fisia" ou "Nao identificado".
- Nunca use razão social como empresa. Exemplo proibido: "SBF COMERCIO DE PRODUTOS ESPORTIVOS S/A".
- Se a empresa não foi consultada na base Info_limites, retorne empresa como "Nao identificado".
- Se o comprovante estiver ilegível, retorne status = "INELEGIVEL".
- Se faltar dado obrigatório, retorne status = "PRECISA_COMPLEMENTO".
- Banco/origem deve ser lido exclusivamente do comprovante.
- Se banco/origem não for identificado, retorne banco como "Nao identificado".
- Banco/origem não é campo obrigatório e não deve gerar pergunta complementar.
- Nunca pergunte banco/origem para a loja.
- A data de movimento vem das respostas do usuário, não do comprovante.
- A data de movimento pode ser uma data única ou um período com data inicial e data final.
- Se houver data_movimento_inicio e data_movimento_fim, use ambas.
- Se houver apenas data_movimento, trate como data única.
- A data final do movimento nunca pode ser maior que a data do depósito/documento.
- Se a data do movimento for maior que a data do depósito/documento, retorne status = "DIVERGENCIA".
- Se todos os dados estiverem completos, retorne status = "PRONTO_PARA_SALVAR".
- Datas devem estar no formato dd/mm/aaaa.
- Se o comprovante trouxer data em outro formato, converta para dd/mm/aaaa.
- Valor deve ser número decimal. Exemplo: 37200.46.
- Não envie nada para Slack.
- Não escreva explicações fora do JSON.
- Para cada item em pendencias, crie uma pergunta correspondente em perguntas_complementares.
- Se loja estiver ausente, perguntas_complementares deve conter: "Qual o número da loja?"
- Se data_movimento estiver ausente, perguntas_complementares deve conter: "Qual a data do movimento desse depósito?"
- Se houve_retirada estiver null, perguntas_complementares deve conter: "Houve alguma retirada desse movimento para estorno, reembolso no ato ou algo do tipo?"
- Se houver múltiplos comprovantes na mesma imagem/PDF, tente identificar e processar todos os comprovantes fisicamente visíveis e legíveis.
- Retorne um array chamado "comprovantes", com um item para cada comprovante físico claramente separado na imagem/PDF.
- qtd_comprovantes_detectados deve representar somente a quantidade de comprovantes físicos visíveis, não a quantidade de códigos, valores, linhas ou blocos de texto.
- Não crie comprovantes adicionais por inferência.
- Não continue sequência de código de autenticação, controle, terminal, envelope ou qualquer outro identificador.
- Não replique comprovantes parecidos como se fossem novos.
- Não deduza valores ausentes com base em padrão visual, repetição ou sequência.
- Não use valores de conta, terminal, envelope, autenticação, telefone, controle ou identificadores como valor de depósito.
- O campo valor_deposito só pode ser preenchido quando houver evidência textual clara no próprio comprovante.
- O valor_deposito deve ser extraído preferencialmente de linhas como "VALOR TOTAL EM DINHEIRO", "VALOR DO DEPÓSITO", "TOTAL", "VALOR RECEBIDO" ou expressão equivalente.
- Cada item do array deve conter: status, indice_comprovante, tipo_documento, data_deposito, valor_deposito, banco, data_geracao_documento, codigo_autenticacao, observacoes, confianca_item, posicao_visual_aproximada e evidencias.
- Para cada comprovante, retorne evidencias com os trechos textuais que sustentam os campos extraídos.
- O objeto evidencias deve conter: valor_deposito_texto, data_deposito_texto, codigo_autenticacao_texto, controle_texto e bloco_textual_comprovante.
- valor_deposito_texto deve conter o trecho exato onde o valor foi lido. Se o valor não estiver claro, retorne valor_deposito como null.
- data_deposito_texto deve conter o trecho exato onde a data foi lida. Se a data não estiver clara, retorne data_deposito como null.
- codigo_autenticacao_texto deve conter o trecho exato onde o código foi lido. Se não estiver claro, retorne codigo_autenticacao como string vazia.
- bloco_textual_comprovante deve conter um resumo textual do próprio comprovante físico usado para aquele item, sem misturar dados de outro comprovante.
- posicao_visual_aproximada deve indicar onde o comprovante aparece na imagem/PDF, por exemplo: "linha 1 coluna 1", "linha 2 coluna 3", "parte superior direita".
- Se algum comprovante estiver cortado, ilegível ou parcialmente visível, inclua esse item com status = "INELEGIVEL", não preencha valor_deposito sem evidência e explique em observacoes.
- Se houver dúvida sobre valor, data ou código, marque o item como "PENDENCIA" ou "INELEGIVEL"; nunca complete por suposição.
- Não misture valores de comprovantes diferentes.
- Não some os valores dos comprovantes.
- Não invente valores, datas, banco, controle ou código de autenticação.
- Quando houver boleto bancário de fundo e recibo/comprovante de Lotéricas CAIXA por cima, trate o conjunto visual como um único comprovante de pagamento, usando o recibo da Lotéricas CAIXA como fonte principal do valor pago.
- Não conte o boleto bancário de fundo como um segundo comprovante separado se ele estiver parcialmente coberto pelo recibo de pagamento.
- Não replique o valor de um comprovante para outro.
- Se dois comprovantes tiverem o mesmo valor, mesma data e mesmo banco, só retorne ambos se houver evidência documental distinta para cada um, com bloco_textual_comprovante e posicao_visual_aproximada claramente diferentes.
- Se houver dúvida se dois itens são comprovantes distintos ou duplicação da mesma leitura, marque o item como PENDENCIA e explique em observacoes.
- Valores iguais podem existir em comprovantes diferentes. Não trate valor repetido como duplicidade por si só.
- Se dois comprovantes tiverem o mesmo valor, retorne ambos somente se houver evidência documental distinta para cada um.
- Para valores repetidos, a posicao_visual_aproximada e o bloco_textual_comprovante precisam indicar claramente comprovantes físicos diferentes.
- Se não for possível distinguir se o valor repetido pertence a dois comprovantes diferentes ou se é duplicação da mesma leitura, marque o item como PENDENCIA e explique em observacoes.
- A loja, empresa, data_movimento, houve_retirada, valor_retirada e motivo_retirada vêm das respostas do usuário e se aplicam ao arquivo inteiro.
- codigo_autenticacao não é campo bloqueante. Se não for identificado, retorne como string vazia e registre em observacoes apenas como observação.

Respostas complementares já informadas pelo usuário:
${JSON.stringify(respostasUsuario || {}, null, 2)}

Retorne exatamente neste formato JSON:

{
  "status": "PRONTO_PARA_SALVAR | PRECISA_COMPLEMENTO | DIVERGENCIA | PENDENCIA | INELEGIVEL",
  "qtd_comprovantes_detectados": 0,
  "alerta_qualidade_imagem": "",
  "criterio_contagem": "",
  "dados_arquivo": {
    "loja": null,
    "empresa": "Nao identificado",
    "data_movimento": null,
    "data_movimento_inicio": null,
    "data_movimento_fim": null,
    "houve_retirada": null,
    "valor_retirada": null,
    "motivo_retirada": null,
    "link_comprovante": ""
  },
  "comprovantes": [
  {
    "status": "PRONTO_PARA_SALVAR | INELEGIVEL | PENDENCIA",
    "indice_comprovante": 1,
    "tipo_documento": "boleto_bancario | deposito_caixa_eletronico | comprovante_carro_forte | outro_comprovante_deposito | nao_identificado",
    "data_deposito": null,
    "valor_deposito": null,
    "banco": "Nao identificado",
    "data_geracao_documento": null,
    "codigo_autenticacao": "",
    "confianca_item": 0,
    "posicao_visual_aproximada": "",
    "evidencias": {
      "valor_deposito_texto": "",
      "data_deposito_texto": "",
      "codigo_autenticacao_texto": "",
      "controle_texto": "",
      "bloco_textual_comprovante": ""
    },
    "observacoes": []
  }
  ],
  "pendencias": [],
  "divergencias": [],
  "perguntas_complementares": [],
  "confianca_geral": 0
}
`;

  const payload = {
    contents: [
      {
        role: "user",
        parts: [
          {
            inline_data: {
              mime_type: mimeType || "application/pdf",
              data: base64Arquivo
            }
          },
          {
            text: prompt
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.2,
      response_mime_type: "application/json"
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode = resp.getResponseCode();
  const txt = resp.getContentText();

  Logger.log("HTTP GEMINI: " + statusCode);
  Logger.log(txt);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("Erro Gemini API HTTP " + statusCode + ": " + txt);
  }

  const json = JSON.parse(txt);

  const respostaTexto =
    json &&
    json.candidates &&
    json.candidates[0] &&
    json.candidates[0].content &&
    json.candidates[0].content.parts &&
    json.candidates[0].content.parts[0] &&
    json.candidates[0].content.parts[0].text;

  if (!respostaTexto) {
    throw new Error("Gemini não retornou texto válido: " + txt);
  }

  const resultado = JSON.parse(respostaTexto);
  return smartSlipAplicarRegrasComplementares_(resultado, respostasUsuario);
}

function smartSlipAplicarRegrasComplementares_(resultado, respostasUsuario) {
  respostasUsuario = respostasUsuario || {};
  resultado = resultado || {};

  if (!resultado.dados_comprovante) {
    resultado.dados_comprovante = {};
  }

  if (!Array.isArray(resultado.pendencias)) {
    resultado.pendencias = [];
  }

  if (!Array.isArray(resultado.divergencias)) {
    resultado.divergencias = [];
  }

  if (!Array.isArray(resultado.perguntas_complementares)) {
    resultado.perguntas_complementares = [];
  }

  const dados = resultado.dados_comprovante;

  smartSlipNormalizarPeriodoMovimento(respostasUsuario, dados);

  // Se a loja foi respondida pelo usuário, usa essa loja.
  if (!dados.loja && respostasUsuario.loja) {
    dados.loja = smartSlipNormalizarLoja4(respostasUsuario.loja);
  }

  // Não aceitar razão social como empresa.
  if (
    dados.empresa &&
    dados.empresa !== "Centauro" &&
    dados.empresa !== "Fisia" &&
    dados.empresa !== "Nao identificado"
  ) {
    dados.empresa = "Nao identificado";
  }

  if (!dados.empresa) {
    dados.empresa = "Nao identificado";
  }

  // Código de autenticação não deve bloquear o fluxo.
  if (dados.codigo_autenticacao === null || dados.codigo_autenticacao === undefined) {
    dados.codigo_autenticacao = "";
  }

  if (dados.link_comprovante === null || dados.link_comprovante === undefined) {
    dados.link_comprovante = "";
  }

  // Data de movimento respondida pelo usuário.
  if (!dados.data_movimento && respostasUsuario.data_movimento) {
    dados.data_movimento = respostasUsuario.data_movimento;
  }

  // Informação de retirada respondida pelo usuário.
  if (
    (dados.houve_retirada === null || dados.houve_retirada === undefined) &&
    typeof respostasUsuario.houve_retirada === "boolean"
  ) {
    dados.houve_retirada = respostasUsuario.houve_retirada;
  }

  if (
    (dados.valor_retirada === null || dados.valor_retirada === undefined) &&
    respostasUsuario.valor_retirada !== undefined
  ) {
    dados.valor_retirada = Number(respostasUsuario.valor_retirada || 0);
  }

  if (!dados.motivo_retirada && respostasUsuario.motivo_retirada) {
    dados.motivo_retirada = respostasUsuario.motivo_retirada;
  }

  // Se não houve retirada, padroniza os campos.
  if (dados.houve_retirada === false) {
    dados.valor_retirada = 0;
    dados.motivo_retirada = dados.motivo_retirada || "Não houve retirada";
  }

  // Regras de pendência/pergunta.
  smartSlipGarantirPergunta_(
    !dados.loja,
    resultado,
    "Loja não informada.",
    "Qual o número da loja?"
  );

  smartSlipGarantirPergunta_(
    !dados.data_movimento,
    resultado,
    "Data de movimento não informada.",
    "Qual a data do movimento desse depósito?"
  );

  smartSlipGarantirPergunta_(
    dados.houve_retirada === null || dados.houve_retirada === undefined,
    resultado,
    "Informação sobre retirada não informada.",
    "Houve alguma retirada desse movimento para estorno, reembolso no ato ou algo do tipo?"
  );

  smartSlipGarantirPergunta_(
    dados.houve_retirada === true && (!dados.valor_retirada || !dados.motivo_retirada),
    resultado,
    "Retirada informada, mas valor ou motivo da retirada está ausente.",
    "Informe o valor e o motivo da retirada."
  );

  smartSlipGarantirPergunta_(
    !dados.data_deposito,
    resultado,
    "Data do depósito não identificada.",
    "Qual a data do depósito?"
  );

  smartSlipGarantirPergunta_(
    dados.valor_deposito === null || dados.valor_deposito === undefined || dados.valor_deposito === "",
    resultado,
    "Valor do depósito não identificado.",
    "Qual o valor do depósito?"
  );

  smartSlipValidarGuardrailMovimento(resultado, dados);

  // Se houver perguntas, o status precisa ser PRECISA_COMPLEMENTO.
  if (resultado.perguntas_complementares.length > 0) {
    resultado.status = "PRECISA_COMPLEMENTO";
  }

  // Se não houver perguntas nem divergências, pode ficar pronto para salvar.
  if (
    resultado.status !== "INELEGIVEL" &&
    resultado.divergencias.length === 0 &&
    resultado.perguntas_complementares.length === 0
  ) {
    resultado.status = "PRONTO_PARA_SALVAR";
  }

  return resultado;
}

function smartSlipGarantirPergunta_(condicao, resultado, pendencia, pergunta) {
  if (!condicao) return;

  if (resultado.pendencias.indexOf(pendencia) === -1) {
    resultado.pendencias.push(pendencia);
  }

  if (resultado.perguntas_complementares.indexOf(pergunta) === -1) {
    resultado.perguntas_complementares.push(pergunta);
  }
}

function smartSlipNormalizarLoja4(loja) {
  const digits = String(loja || "").replace(/\D/g, "");

  if (!digits) {
    return "";
  }

  return digits.padStart(4, "0").slice(-4);
}

function smartSlipTesteSalvarCompleto() {
  const fileId = "1DtSSYlpbdpjDubLd05wue-XKgaosjcRN";

const respostasUsuario = {
  loja: "0241",
  data_movimento_inicio: "01/05/2026",
   data_movimento_fim: "04/05/2026",
  houve_retirada: false,
  valor_retirada: 0,
  motivo_retirada: "Não houve retirada",
  mais_comprovantes: false
};

const protocolo = "SS-TESTE-SALVAR-COMPLETO";

const arquivoOrganizado = smartSlipCopiarArquivoParaPastaEstruturada(
  fileId,
  respostasUsuario,
  protocolo
);

  const file = DriveApp.getFileById(arquivoOrganizado.file_id);
  const linkComprovante = arquivoOrganizado.link_comprovante;
  const blob = file.getBlob();

  const base64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  const resultado = smartSlipChamarGemini(base64, mimeType, respostasUsuario);

  Logger.log("RESULTADO GEMINI:");
  Logger.log(JSON.stringify(resultado, null, 2));

  const dados = resultado.dados_comprovante || {};

  // Usa respostas complementares do usuário como fonte de verdade quando o comprovante não traz o dado.
if (!dados.banco && respostasUsuario.banco) {
  dados.banco = respostasUsuario.banco;
}

if (!dados.data_movimento && respostasUsuario.data_movimento) {
  dados.data_movimento = respostasUsuario.data_movimento;
}

if (
  (dados.houve_retirada === null || dados.houve_retirada === undefined) &&
  typeof respostasUsuario.houve_retirada === "boolean"
) {
  dados.houve_retirada = respostasUsuario.houve_retirada;
}

if (
  (dados.valor_retirada === null || dados.valor_retirada === undefined) &&
  respostasUsuario.valor_retirada !== undefined
) {
  dados.valor_retirada = Number(respostasUsuario.valor_retirada || 0);
}

if (!dados.motivo_retirada && respostasUsuario.motivo_retirada) {
  dados.motivo_retirada = respostasUsuario.motivo_retirada;
}

if (dados.codigo_autenticacao === null || dados.codigo_autenticacao === undefined) {
  dados.codigo_autenticacao = "";
}

// 1. A loja informada pelo usuário é a fonte de verdade.
// Não confiar em número de loja inferido do comprovante.
if (respostasUsuario.loja) {
  dados.loja = String(respostasUsuario.loja).replace(/\D/g, "").padStart(4, "0").slice(-4);

  resultado.divergencias = (resultado.divergencias || []).filter(function(d) {
    const txt = String(d || "").toLowerCase();

    return !(
      txt.includes("número da loja no comprovante") ||
      txt.includes("numero da loja no comprovante") ||
      txt.includes("difere do número de loja informado") ||
      txt.includes("difere do numero de loja informado")
    );
  });
}

  // 2. Consulta Info_limites para obter empresa correta.
  if (dados.loja) {
    const infoLoja = smartSlipConsultarInfoLoja_(dados.loja);

    Logger.log("INFO LOJA:");
    Logger.log(JSON.stringify(infoLoja, null, 2));

    if (infoLoja.ok) {
      dados.loja = infoLoja.loja_normalizada;
      dados.empresa = infoLoja.empresa;
    } else {
      resultado.status = "PRECISA_COMPLEMENTO";
      resultado.pendencias = resultado.pendencias || [];
      resultado.perguntas_complementares = resultado.perguntas_complementares || [];

      resultado.pendencias.push("Loja não encontrada na Info_limites.");
      resultado.perguntas_complementares.push("Confirme o número correto da loja.");

      Logger.log("NÃO SALVOU: loja não encontrada na Info_limites.");
      Logger.log(JSON.stringify(resultado, null, 2));
      return;
    }
  }

  // Código de autenticação não bloqueia salvamento.
resultado.pendencias = (resultado.pendencias || []).filter(function(p) {
  const txt = String(p || "").toLowerCase();

  return !(
    txt.includes("código de autenticação") ||
    txt.includes("codigo de autenticacao")
  );
});

  // 3. Valida campos obrigatórios antes de salvar.
  const pendenciasBloqueantes = [];

  if (!dados.loja) pendenciasBloqueantes.push("Loja não informada.");
  if (!dados.empresa || dados.empresa === "Nao identificado") pendenciasBloqueantes.push("Empresa não identificada pela Info_limites.");
  if (!resultado.tipo_documento || resultado.tipo_documento === "nao_identificado") pendenciasBloqueantes.push("Tipo de documento não identificado.");
  if (!dados.data_deposito) pendenciasBloqueantes.push("Data do depósito não identificada.");
  if (dados.valor_deposito === null || dados.valor_deposito === undefined || dados.valor_deposito === "") pendenciasBloqueantes.push("Valor do depósito não identificado.");
  if (!dados.data_movimento) pendenciasBloqueantes.push("Data de movimento não informada.");
  if (dados.houve_retirada === null || dados.houve_retirada === undefined) pendenciasBloqueantes.push("Informação sobre retirada não informada.");
  if (!dados.data_geracao_documento) pendenciasBloqueantes.push("Data de geração do documento não identificada.");

  if (dados.houve_retirada === true) {
    if (!dados.valor_retirada) pendenciasBloqueantes.push("Valor da retirada não informado.");
    if (!dados.motivo_retirada) pendenciasBloqueantes.push("Motivo da retirada não informado.");
  }

  if (pendenciasBloqueantes.length > 0) {
    resultado.status = "PRECISA_COMPLEMENTO";
    resultado.pendencias = pendenciasBloqueantes;
    resultado.perguntas_complementares = smartSlipMontarPerguntasComplementares(pendenciasBloqueantes);

    Logger.log("NÃO SALVOU: existem pendências bloqueantes.");
    Logger.log(JSON.stringify(resultado, null, 2));
    return;
  }

  // 4. Monta payload final para salvar.
  const payloadSalvar = {
    loja: dados.loja,
    empresa: dados.empresa,
    tipo_documento: resultado.tipo_documento,
    data_deposito: dados.data_deposito,
    valor_deposito: Number(dados.valor_deposito || 0),
    banco: dados.banco || "Nao identificado",
    data_movimento: dados.data_movimento,
    houve_retirada: dados.houve_retirada === true,
    valor_retirada: Number(dados.valor_retirada || 0),
    motivo_retirada: dados.houve_retirada === true
      ? dados.motivo_retirada
      : "Não houve retirada",
    data_geracao_documento: dados.data_geracao_documento,
    codigo_autenticacao: dados.codigo_autenticacao || "",
    link_comprovante: linkComprovante,
    status_processamento: "PRONTO_PARA_SALVAR",
    pendencias: [],
    divergencias: resultado.divergencias || [],
    confianca_geral: Number(resultado.confianca_geral || 0)
  };

  Logger.log("PAYLOAD PARA SALVAR:");
  Logger.log(JSON.stringify(payloadSalvar, null, 2));

  // 5. Salva na BASE_SMARTSLIP.
  const retornoSalvar = smartSlipSalvarDadosComprovantePlanilha_(payloadSalvar);

  Logger.log("RETORNO SALVAMENTO:");
  Logger.log(JSON.stringify(retornoSalvar, null, 2));

  Logger.log("PROCESSO CONCLUÍDO. Verifique a aba BASE_SMARTSLIP.");
}

function smartSlipMontarPerguntasComplementares(pendencias) {
  const perguntas = [];

  pendencias.forEach(function(p) {
    const txt = String(p || "").toLowerCase();

    if (txt.includes("loja")) {
      perguntas.push("Qual o número da loja?");
    } else if (txt.includes("empresa")) {
      perguntas.push("Confirme o número da loja para consultar a empresa correta.");
    } else if (txt.includes("tipo de documento")) {
      perguntas.push("Qual é o tipo do comprovante?");
    } else if (txt.includes("data do depósito")) {
      perguntas.push("Qual a data do depósito?");
    } else if (txt.includes("valor do depósito")) {
      perguntas.push("Qual o valor do depósito?");
    } else if (txt.includes("banco")) {
      perguntas.push("Qual o banco do comprovante?");
    } else if (txt.includes("data de movimento")) {
      perguntas.push("Qual a data do movimento desse depósito?");
    } else if (txt.includes("retirada")) {
      perguntas.push("Houve alguma retirada desse movimento para estorno, reembolso no ato ou algo do tipo?");
    } else if (txt.includes("data de geração")) {
      perguntas.push("Qual a data de geração do documento?");
    }
  });

  return [...new Set(perguntas)];
}

function smartSlipGarantirAbaFila() {
  const ss = SpreadsheetApp.openById(SMARTSLIP_DB_SPREADSHEET_ID);
  let sh = ss.getSheetByName(SMARTSLIP_ABA_FILA);

  if (!sh) {
    sh = ss.insertSheet(SMARTSLIP_ABA_FILA);
    sh.appendRow([
      "Data Registro",
      "Protocolo",
      "Email Usuário",
      "Loja Informada",
      "Banco Informado",
      "Data Movimento",
      "Houve Retirada",
      "Valor Retirada",
      "Motivo Retirada",
      "Mais Comprovantes",
      "File ID",
      "Link Comprovante",
      "Nome Arquivo",
      "Status Fila",
      "Mensagem",
      "JSON Respostas",
      "JSON Resultado",
      "Data Processamento"
    ]);
  }

  return sh;
}

function smartSlipTesteFila() {
  const sh = smartSlipGarantirCabecalhoFila();

  const respostasUsuario = {
    loja: "0025",
    banco: "CARRO FORTE",
    data_movimento: "04/05/2026",
    houve_retirada: false,
    valor_retirada: 0,
    motivo_retirada: "Não houve retirada",
    mais_comprovantes: false
  };

  const dataMovimentoTexto = smartSlipMontarDataMovimentoTexto(respostasUsuario);

  sh.appendRow([
    new Date(),
    "SS-TESTE-001",
    Session.getActiveUser().getEmail() || "",
    smartSlipNormalizarLoja4(respostasUsuario.loja),
    respostasUsuario.banco || "",
    dataMovimentoTexto,
    respostasUsuario.houve_retirada ? "Sim" : "Não",
    respostasUsuario.valor_retirada,
    respostasUsuario.motivo_retirada,
    respostasUsuario.mais_comprovantes ? "Sim" : "Não",
    "FILE_ID_TESTE",
    "https://drive.google.com/teste",
    "comprovante_teste.pdf",
    "RECEBIDO",
    "Linha de teste criada na fila.",
    JSON.stringify(respostasUsuario),
    "",
    ""
  ]);

  Logger.log("Linha de teste criada na SMARTSLIP_FILA.");
}

function smartSlipGetScriptProperty(nome) {
  const valor = PropertiesService
    .getScriptProperties()
    .getProperty(nome);

  if (!valor) {
    throw new Error("Propriedade do script não configurada: " + nome);
  }

  return valor;
}

function smartSlipGetRootFolderComprovantes() {
  const folderId = smartSlipGetScriptProperty("SMARTSLIP_FOLDER_ID");
  return DriveApp.getFolderById(folderId);
}

function smartSlipGetOrCreateFolder(parentFolder, folderName) {
  const nome = smartSlipSanitizarNomePasta(folderName);
  const folders = parentFolder.getFoldersByName(nome);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(nome);
}

function smartSlipSanitizarNomePasta(nome) {
  return String(nome || "")
    .trim()
    .replace(/[\\\/:*?"<>|#%{}~&]/g, "-")
    .replace(/\s+/g, " ")
    .substring(0, 120);
}

function smartSlipNormalizarLojaParaPasta(loja) {
  const digits = String(loja || "").replace(/\D/g, "");

  if (!digits) {
    throw new Error("Loja inválida para criação de pasta.");
  }

  return digits.padStart(4, "0").slice(-4);
}

function smartSlipObterAnoMes(dataMovimento) {
  const data = smartSlipParseDateBR(dataMovimento);

  if (data) {
    const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    return {
      ano: Utilities.formatDate(data, tz, "yyyy"),
      mes: Utilities.formatDate(data, tz, "MM")
    };
  }

  const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  const hoje = new Date();

  return {
    ano: Utilities.formatDate(hoje, tz, "yyyy"),
    mes: Utilities.formatDate(hoje, tz, "MM")
  };
}

function smartSlipObterPastaAnoMesLoja(dataMovimento, loja) {
  const root = smartSlipGetRootFolderComprovantes();
  const anoMes = smartSlipObterAnoMes(dataMovimento);
  const loja4 = smartSlipNormalizarLojaParaPasta(loja);

  const pastaAno = smartSlipGetOrCreateFolder(root, anoMes.ano);
  const pastaMes = smartSlipGetOrCreateFolder(pastaAno, anoMes.mes);
  const pastaLoja = smartSlipGetOrCreateFolder(pastaMes, loja4);

  return pastaLoja;
}

function smartSlipMontarNomeArquivoComprovante(protocolo, loja, dataMovimento, nomeOriginal) {
  const loja4 = smartSlipNormalizarLojaParaPasta(loja);
  const data = String(dataMovimento || "sem-data").trim();
  const nome = String(nomeOriginal || "comprovante.pdf").trim();

  const nomeLimpo = nome
    .replace(/[\\\/:*?"<>|#%{}~&]/g, "-")
    .replace(/\s+/g, "_")
    .substring(0, 120);

  return [
    protocolo || "SS-SEM-PROTOCOLO",
    loja4,
    data,
    nomeLimpo
  ].join("_");
}

function smartSlipSalvarUploadEmPastaEstruturada(form, respostasUsuario, protocolo) {
  if (!form || !form.arquivoBase64) {
    throw new Error("Arquivo não recebido.");
  }

  respostasUsuario = respostasUsuario || {};

  const loja = respostasUsuario.loja || form.loja;

  const dataMovimentoTexto = smartSlipMontarDataMovimentoTexto(respostasUsuario);

  const dataMovimentoReferencia =
    respostasUsuario.data_movimento_inicio ||
    respostasUsuario.data_movimento ||
    respostasUsuario.data_movimento_fim ||
    form.data_movimento ||
    "";

  if (!loja) {
    throw new Error("Loja não informada para salvar comprovante.");
  }

  if (!dataMovimentoTexto) {
    throw new Error("Data de movimento não informada para salvar comprovante.");
  }

  if (!dataMovimentoReferencia) {
    throw new Error("Data de referência do movimento não informada para salvar comprovante.");
  }

  const pastaDestino = smartSlipObterPastaAnoMesLoja(dataMovimentoReferencia, loja);

  const base64Limpo = String(form.arquivoBase64).split(",").pop();
  const bytes = Utilities.base64Decode(base64Limpo);

  const nomeArquivo = smartSlipMontarNomeArquivoComprovante(
    protocolo,
    loja,
    dataMovimentoTexto,
    form.nomeArquivo || "comprovante.pdf"
  );

  const mimeType = form.mimeType || "application/pdf";
  const blob = Utilities.newBlob(bytes, mimeType, nomeArquivo);

  const file = pastaDestino.createFile(blob);

  return {
    file_id: file.getId(),
    link_comprovante: file.getUrl(),
    nome_arquivo: file.getName(),
    pasta_id: pastaDestino.getId(),
    pasta_url: pastaDestino.getUrl()
  };
}

function smartSlipCopiarArquivoParaPastaEstruturada(fileId, respostasUsuario, protocolo) {
  respostasUsuario = respostasUsuario || {};

  const loja = respostasUsuario.loja;

  const dataMovimentoTexto = smartSlipMontarDataMovimentoTexto(respostasUsuario);

  const dataMovimentoReferencia =
    respostasUsuario.data_movimento_inicio ||
    respostasUsuario.data_movimento ||
    respostasUsuario.data_movimento_fim ||
    "";

  if (!fileId) {
    throw new Error("fileId não informado.");
  }

  if (!loja) {
    throw new Error("Loja não informada.");
  }

  if (!dataMovimentoTexto) {
    throw new Error("Data de movimento não informada.");
  }

  if (!dataMovimentoReferencia) {
    throw new Error("Data de referência do movimento não informada.");
  }

  const fileOriginal = DriveApp.getFileById(fileId);

  // A pasta ano/mês/loja usa uma data única de referência.
  // Se for período, usa a data inicial.
  const pastaDestino = smartSlipObterPastaAnoMesLoja(dataMovimentoReferencia, loja);

  const nomeArquivo = smartSlipMontarNomeArquivoComprovante(
    protocolo,
    loja,
    dataMovimentoTexto,
    fileOriginal.getName()
  );

  const novoArquivo = fileOriginal.makeCopy(nomeArquivo, pastaDestino);

  return {
    file_id: novoArquivo.getId(),
    link_comprovante: novoArquivo.getUrl(),
    nome_arquivo: novoArquivo.getName(),
    pasta_id: pastaDestino.getId(),
    pasta_url: pastaDestino.getUrl()
  };
}

function smartSlipTesteCriarPastasComprovante() {
  const fileId = "1DtSSYlpbdpjDubLd05wue-XKgaosjcRN";

  const respostasUsuario = {
    loja: "0178",
    data_movimento_inicio: "01/05/2026",
    data_movimento_fim: "04/05/2026",
    houve_retirada: false,
    valor_retirada: 0,
    motivo_retirada: "Não houve retirada",
    mais_comprovantes: false
  };

  const protocolo = "SS-TESTE-PASTAS";

  const arquivo = smartSlipCopiarArquivoParaPastaEstruturada(
    fileId,
    respostasUsuario,
    protocolo
  );

  Logger.log(JSON.stringify(arquivo, null, 2));
}

function smartSlipParseDateBR(valor) {
  if (!valor) return null;

  const txt = String(valor).trim();

  // Aceita dd/mm/aaaa
  let m = txt.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  }

  // Aceita yyyy-mm-dd também, para compatibilidade com input type="date"
  m = txt.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  return null;
}

function smartSlipFormatDateBR(date) {
  const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  return Utilities.formatDate(date, tz, "dd/MM/yyyy");
}

function smartSlipNormalizarPeriodoMovimento(respostasUsuario, dados) {
  respostasUsuario = respostasUsuario || {};
  dados = dados || {};

  let inicio =
    respostasUsuario.data_movimento_inicio ||
    respostasUsuario.data_movimento ||
    dados.data_movimento_inicio ||
    dados.data_movimento ||
    "";

  let fim =
    respostasUsuario.data_movimento_fim ||
    dados.data_movimento_fim ||
    inicio;

  dados.data_movimento_inicio = inicio || "";
  dados.data_movimento_fim = fim || "";

  if (inicio && fim && inicio === fim) {
    dados.data_movimento = inicio;
  } else if (inicio && fim) {
    dados.data_movimento = inicio + " a " + fim;
  } else {
    dados.data_movimento = inicio || "";
  }

  return dados;
}

function smartSlipValidarGuardrailMovimento(resultado, dados) {
  resultado = resultado || {};
  dados = dados || {};

  if (!Array.isArray(resultado.divergencias)) {
    resultado.divergencias = [];
  }

  const dataDeposito = smartSlipParseDateBR(dados.data_deposito);
  const movInicio = smartSlipParseDateBR(dados.data_movimento_inicio);
  const movFim = smartSlipParseDateBR(dados.data_movimento_fim);

  if (!dataDeposito || !movInicio || !movFim) {
    return resultado;
  }

  if (movInicio > movFim) {
    resultado.status = "DIVERGENCIA";
    resultado.divergencias.push(
      "Data inicial do movimento maior que a data final do movimento."
    );
  }

  if (movFim > dataDeposito) {
    resultado.status = "DIVERGENCIA";
    resultado.divergencias.push(
      "Data de movimento maior que a data do depósito/documento. Movimento: " +
      smartSlipFormatDateBR(movFim) +
      " | Depósito/documento: " +
      smartSlipFormatDateBR(dataDeposito)
    );
  }

  return resultado;
}

function smartSlipGarantirCabecalhoFila() {
  const ss = SpreadsheetApp.openById(SMARTSLIP_DB_SPREADSHEET_ID);
  let sh = ss.getSheetByName(SMARTSLIP_ABA_FILA);

  if (!sh) {
    sh = ss.insertSheet(SMARTSLIP_ABA_FILA);
  }

  const headers = [
    "Data Registro",
    "Protocolo",
    "Email Usuário",
    "Loja Informada",
    "Banco",
    "Data Movimento",
    "Houve Retirada",
    "Valor Retirada",
    "Motivo Retirada",
    "Mais Comprovantes",
    "File ID",
    "Link Comprovante",
    "Nome Arquivo",
    "Status Fila",
    "Mensagem",
    "JSON Respostas",
    "JSON Resultado",
    "Data Processamento"
  ];

  const primeiraCelula = String(sh.getRange(1, 1).getValue() || "").trim();

  if (primeiraCelula !== "Data Registro") {
    sh.insertRowBefore(1);
  }

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  return sh;
}

function smartSlipTesteCabecalhoFila() {
  smartSlipGarantirCabecalhoFila();
  Logger.log("Cabeçalho da SMARTSLIP_FILA criado/verificado.");
}

function smartSlipMontarDataMovimentoTexto(respostasUsuario) {
  respostasUsuario = respostasUsuario || {};

  const inicio = respostasUsuario.data_movimento_inicio || respostasUsuario.data_movimento || "";
  const fim = respostasUsuario.data_movimento_fim || "";

  if (inicio && fim && inicio !== fim) {
    return inicio + " a " + fim;
  }

  return inicio || fim || "";
}

function smartSlipValidarValoresDiariosMovimento_(payload, dataInicioBR, dataFimBR) {
  const inicio = smartSlipParseDateBR(dataInicioBR);
  const fim = smartSlipParseDateBR(dataFimBR);

  if (!inicio || !fim) {
    throw new Error("Período de movimento inválido para validação dos valores diários.");
  }

  if (fim < inicio) {
    throw new Error("Data movimento final menor que a data movimento inicial.");
  }

  const datasEsperadas = smartSlipGerarRangeDatasBR_(inicio, fim);

  if (datasEsperadas.length > 31) {
    throw new Error("O intervalo de movimento não pode passar de 31 dias.");
  }

  if (typeof payload === "string") {
    try {
      payload = JSON.parse(payload || "{}");
    } catch (err) {
      throw new Error("Valores diários do movimento em formato inválido.");
    }
  }

  payload = payload || {};

  const itens = Array.isArray(payload.itens)
    ? payload.itens
    : Array.isArray(payload)
      ? payload
      : [];

  if (itens.length !== datasEsperadas.length) {
    throw new Error(
      "Quantidade de valores diários diferente do período informado. Esperado: " +
      datasEsperadas.length +
      " dia(s). Recebido: " +
      itens.length +
      "."
    );
  }

  const mapa = {};
  let total = 0;

  itens.forEach(function(item) {
    const data = String(item.data || "").trim();
    const valor = Number(item.valor || 0);

    if (!data) {
      throw new Error("Existe valor diário sem data informada.");
    }

    if (isNaN(valor) || valor <= 0) {
      throw new Error("Valor diário inválido para " + data + ". Informe valor maior que zero.");
    }

    if (valor < 0) {
      throw new Error("Valor diário negativo não é permitido para " + data + ".");
    }

    mapa[data] = valor;
    total += valor;
  });

  datasEsperadas.forEach(function(data) {
    if (!mapa[data]) {
      throw new Error("Valor diário não informado para " + data + ".");
    }
  });

  return {
    periodo_inicio: dataInicioBR,
    periodo_fim: dataFimBR,
    itens: datasEsperadas.map(function(data) {
      return {
        data: data,
        valor: mapa[data]
      };
    }),
    total: total
  };
}

function smartSlipGerarRangeDatasBR_(inicio, fim) {
  const datas = [];
  const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

  const atual = new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate());
  const limite = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate());

  while (atual <= limite) {
    datas.push(Utilities.formatDate(atual, tz, "dd/MM/yyyy"));
    atual.setDate(atual.getDate() + 1);
  }

  return datas;
}

function smartSlipTesteFilaPeriodo() {
  const sh = smartSlipGarantirCabecalhoFila();

  const respostasUsuario = {
    loja: "0171",
    banco: "",
    data_movimento_inicio: "01/05/2026",
    data_movimento_fim: "04/05/2026",
    houve_retirada: false,
    valor_retirada: 0,
    motivo_retirada: "Não houve retirada",
    mais_comprovantes: false
  };

  const dataMovimentoTexto = smartSlipMontarDataMovimentoTexto(respostasUsuario);

  sh.appendRow([
    new Date(),
    "SS-TESTE-PERIODO-001",
    Session.getActiveUser().getEmail() || "",
    smartSlipNormalizarLoja4(respostasUsuario.loja),
    respostasUsuario.banco || "",
    dataMovimentoTexto,
    respostasUsuario.houve_retirada ? "Sim" : "Não",
    respostasUsuario.valor_retirada,
    respostasUsuario.motivo_retirada,
    respostasUsuario.mais_comprovantes ? "Sim" : "Não",
    "FILE_ID_TESTE",
    "https://drive.google.com/teste",
    "comprovante_teste_periodo.pdf",
    "RECEBIDO",
    "Linha de teste com período criada na fila.",
    JSON.stringify(respostasUsuario),
    "",
    ""
  ]);

  Logger.log("Linha de teste com período criada na SMARTSLIP_FILA.");
  Logger.log("Data Movimento gravada: " + dataMovimentoTexto);
}

function doGet(e) {
  const faviconUrl = "https://raw.githubusercontent.com/rslisboa/smartslip/051387a22e6d310c60e8855e95f9cf3ed8ded2e6/alfabeto.png";

  return HtmlService
    .createHtmlOutputFromFile("SmartSlipApp")
    .setTitle("SmartSlip - Envio de Comprovantes")
    .setFaviconUrl(faviconUrl)
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function smartSlipEnviarComprovanteApp(form) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!form) {
      throw new Error("Formulário não recebido.");
    }

    if (!form.arquivoBase64) {
      throw new Error("Arquivo não recebido.");
    }

    if (!form.loja) {
      throw new Error("Número da loja não informado.");
    }

    if (!form.data_movimento_inicio) {
      throw new Error("Data de movimento inicial não informada.");
    }

    if (!form.data_movimento_fim) {
      throw new Error("Data de movimento final não informada. Se for uma única data, repita a data inicial.");
    }

    const dataInicioObj = smartSlipParseDateBR(form.data_movimento_inicio);
    const dataFimObj = smartSlipParseDateBR(form.data_movimento_fim);

    if (!dataInicioObj) {
      throw new Error("Data de movimento inicial inválida.");
    }

    if (!dataFimObj) {
      throw new Error("Data de movimento final inválida.");
    }

    dataInicioObj.setHours(0, 0, 0, 0);
    dataFimObj.setHours(0, 0, 0, 0);

    if (dataFimObj < dataInicioObj) {
      throw new Error("A data movimento final não pode ser menor que a data movimento inicial.");
    }

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    if (dataFimObj > hoje) {
      throw new Error("A data movimento final não pode ser maior que a data atual.");
    }

    const valoresDiariosMovimento = smartSlipValidarValoresDiariosMovimento_(
      form.valores_diarios_movimento,
      form.data_movimento_inicio,
      form.data_movimento_fim
    );

    const houveRetiradaValidacao = String(form.houve_retirada) === "true";
    const valorRetiradaValidacao = Number(form.valor_retirada || 0);
    const motivoRetiradaValidacao = String(form.motivo_retirada || "").trim();

    if (valorRetiradaValidacao < 0) {
      throw new Error("O valor da retirada não pode ser negativo.");
    }

    if (houveRetiradaValidacao && (!valorRetiradaValidacao || valorRetiradaValidacao <= 0)) {
      throw new Error("Informe o valor da retirada quando 'Houve retirada?' estiver marcado como Sim.");
    }

    if (houveRetiradaValidacao && !motivoRetiradaValidacao) {
      throw new Error("Informe o motivo da retirada quando 'Houve retirada?' estiver marcado como Sim.");
    }

    if (!houveRetiradaValidacao && valorRetiradaValidacao > 0) {
      throw new Error("Se houver valor de retirada, marque 'Houve retirada?' como Sim.");
    }

    const usuarioAtual = smartSlipGetUsuarioAtual();

    const lojaForm = smartSlipNormalizarLoja4(form.loja);
    const lojaPadraoUsuario = smartSlipNormalizarLoja4(usuarioAtual.loja_padrao || "");

    if (!usuarioAtual.is_admin) {
      if (!lojaPadraoUsuario) {
        throw new Error("Configure sua loja padrão antes de enviar comprovantes.");
      }

      if (lojaForm && lojaForm !== lojaPadraoUsuario) {
        throw new Error("A loja enviada não corresponde à loja padrão configurada para seu usuário.");
      }
    }

    const protocolo = smartSlipGerarProtocolo();
    const loja4 = usuarioAtual.is_admin ? lojaForm : lojaPadraoUsuario;

    if (!loja4) {
      throw new Error("Loja inválida para envio.");
    }

    const houveRetiradaResposta = String(form.houve_retirada) === "true";

    const respostasUsuario = {
      loja: loja4,
      data_movimento_inicio: form.data_movimento_inicio,
      data_movimento_fim: form.data_movimento_fim,
      houve_retirada: houveRetiradaResposta,
      valor_retirada: houveRetiradaResposta ? Number(form.valor_retirada || 0) : 0,
      motivo_retirada: houveRetiradaResposta
        ? String(form.motivo_retirada || "").trim()
        : "Não houve retirada",
      valores_diarios_movimento: valoresDiariosMovimento,
      mais_comprovantes: String(form.mais_comprovantes) === "true"
    };

    const dataMovimentoTexto = smartSlipMontarDataMovimentoTexto(respostasUsuario);

    const arquivo = smartSlipSalvarUploadEmPastaEstruturada(
      form,
      respostasUsuario,
      protocolo
    );

    const sh = smartSlipGarantirCabecalhoFila();

    const emailUsuario = Session.getActiveUser().getEmail() || "";

    sh.appendRow([
      new Date(),
      protocolo,
      emailUsuario,
      loja4,
      "",
      dataMovimentoTexto,
      respostasUsuario.houve_retirada ? "Sim" : "Não",
      respostasUsuario.valor_retirada,
      respostasUsuario.motivo_retirada,
      respostasUsuario.mais_comprovantes ? "Sim" : "Não",
      arquivo.file_id,
      arquivo.link_comprovante,
      arquivo.nome_arquivo,
      "RECEBIDO",
      "Comprovante recebido e aguardando processamento.",
      JSON.stringify(respostasUsuario),
      "",
      ""
    ]);

    return {
      ok: true,
      protocolo: protocolo,
      status: "RECEBIDO",
      mensagem: "Comprovante recebido com sucesso."
    };

  } catch (err) {
    return {
      ok: false,
      erro: String(err && err.message ? err.message : err)
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function smartSlipGerarProtocolo() {
  const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  const agora = new Date();
  const data = Utilities.formatDate(agora, tz, "yyyyMMdd-HHmmss");
  const rand = Math.floor(Math.random() * 9000) + 1000;

  return "SS-" + data + "-" + rand;
}

function smartSlipProcessarFila() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    return;
  }

  try {
    const sh = smartSlipGarantirCabecalhoFila();
    const lastRow = sh.getLastRow();

    if (lastRow < 2) {
      return;
    }

    const values = sh.getRange(2, 1, lastRow - 1, 18).getValues();
    const limitePorExecucao = 10;
    let processados = 0;

    for (let i = 0; i < values.length; i++) {
      if (processados >= limitePorExecucao) {
        break;
      }

      const rowIndex = i + 2;
      const row = values[i];
      const protocolo = row[1];
      const fileId = row[10];
      const statusFila = row[13];
      const jsonRespostas = row[15];

      if (statusFila !== "RECEBIDO") {
        continue;
      }

      sh.getRange(rowIndex, 14).setValue("PROCESSANDO");
      sh.getRange(rowIndex, 15).setValue("Processando comprovante via Gemini...");

      try {
        const respostasUsuario = JSON.parse(jsonRespostas || "{}");
        const resultado = smartSlipProcessarComprovanteFila(fileId, respostasUsuario, protocolo);

        sh.getRange(rowIndex, 14).setValue(resultado.status_fila);
        sh.getRange(rowIndex, 15).setValue(resultado.mensagem || "");
        sh.getRange(rowIndex, 17).setValue(JSON.stringify(resultado));
        sh.getRange(rowIndex, 18).setValue(new Date());

        processados++;

      } catch (err) {
        sh.getRange(rowIndex, 14).setValue("ERRO_PROCESSAMENTO");
        sh.getRange(rowIndex, 15).setValue(String(err && err.message ? err.message : err));
        sh.getRange(rowIndex, 18).setValue(new Date());
      }
    }

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function smartSlipProcessarComprovanteFila(fileId, respostasUsuario, protocolo) {
  const file = DriveApp.getFileById(fileId);
  const linkComprovante = file.getUrl();
  const blob = file.getBlob();

  const base64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  const resultado = smartSlipChamarGemini(base64, mimeType, respostasUsuario);

  const loja4 = smartSlipNormalizarLoja4(respostasUsuario.loja);
  const dataMovimentoTexto = smartSlipMontarDataMovimentoTexto(respostasUsuario);
  const valoresDiariosMovimento = respostasUsuario.valores_diarios_movimento || null;

  const infoLoja = smartSlipConsultarInfoLoja_(loja4);

  if (!infoLoja.ok) {
    return {
      ok: false,
      status_fila: "PENDENTE_INTERNO",
      mensagem: "Loja não encontrada na Info_limites.",
      resultado: resultado
    };
  }

  const empresa = infoLoja.empresa;

  const houveRetirada = respostasUsuario.houve_retirada === true;
  const valorRetirada = Number(respostasUsuario.valor_retirada || 0);
  const motivoRetirada = houveRetirada
    ? respostasUsuario.motivo_retirada || ""
    : "Não houve retirada";

  const comprovantes = smartSlipExtrairComprovantesDoResultado(resultado);

  if (!comprovantes.length) {
    return {
      ok: false,
      status_fila: "PENDENTE_INTERNO",
      mensagem: "Nenhum comprovante legível foi identificado no arquivo.",
      resultado: resultado
    };
  }

  if (comprovantes.length > 10) {
  return {
    ok: false,
    status_fila: "PENDENTE_INTERNO",
    mensagem:
      "Validação anti-alucinação bloqueou o lote. A IA retornou " +
      comprovantes.length +
      " comprovantes, acima do limite operacional de 10 por arquivo. Nenhum item foi salvo na BASE_SMARTSLIP.",
    resultado: resultado
  };
}

const validacaoDuplicidadeEvidencia = smartSlipValidarDuplicidadeEvidenciaLote_(comprovantes);

if (!validacaoDuplicidadeEvidencia.ok) {
  return {
    ok: false,
    status_fila: "PENDENTE_INTERNO",
    mensagem: validacaoDuplicidadeEvidencia.mensagem + " Nenhum item foi salvo na BASE_SMARTSLIP.",
    validacao_duplicidade_evidencia: validacaoDuplicidadeEvidencia,
    resultado: resultado
  };
}

  let salvos = 0;
  let pendentes = 0;
  const mensagens = [];

  let retiradaJaAplicada = false;

  comprovantes.forEach(function(item) {
    const pendenciasItem = [];

    if (!item.tipo_documento || item.tipo_documento === "nao_identificado") {
      pendenciasItem.push("Tipo de documento não identificado.");
    }

    if (!item.data_deposito) {
      pendenciasItem.push("Data do depósito não identificada.");
    }

    if (item.valor_deposito === null || item.valor_deposito === undefined || item.valor_deposito === "") {
      pendenciasItem.push("Valor do depósito não identificado.");
    }

    if (!item.data_geracao_documento) {
      pendenciasItem.push("Data de geração do documento não identificada.");
    }

    const dadosGuardrail = {
      data_deposito: item.data_deposito,
      data_movimento: dataMovimentoTexto,
      data_movimento_inicio: respostasUsuario.data_movimento_inicio || respostasUsuario.data_movimento || "",
      data_movimento_fim: respostasUsuario.data_movimento_fim || respostasUsuario.data_movimento_inicio || respostasUsuario.data_movimento || ""
    };

    const resultadoGuardrail = {
      status: "PRONTO_PARA_SALVAR",
      divergencias: []
    };

    smartSlipValidarGuardrailMovimento(resultadoGuardrail, dadosGuardrail);

    if (resultadoGuardrail.divergencias.length > 0) {
      pendenciasItem.push(resultadoGuardrail.divergencias.join(" | "));
    }

    if (pendenciasItem.length > 0 || item.status === "INELEGIVEL") {
      pendentes++;
      mensagens.push(
        "Comprovante " +
        (item.indice_comprovante || "?") +
        ": " +
        pendenciasItem.join(" | ")
      );
      return;
    }

    const validacaoEvidencia = smartSlipItemTemInferenciaOuBaixaEvidencia_(item);

    if (validacaoEvidencia.bloquear) {
      pendentes++;

      mensagens.push(
        "Comprovante " +
        (item.indice_comprovante || "?") +
        ": bloqueado por validação anti-alucinação. " +
        validacaoEvidencia.motivo
      );

      return;
    }

    const aplicarRetiradaNestaLinha = houveRetirada && !retiradaJaAplicada;
    const indiceComprovante = item.indice_comprovante || (salvos + pendentes + 1);
    const idComprovante = smartSlipGerarIdComprovante(
      protocolo,
      indiceComprovante
    );

    const payloadSalvar = {
      id_comprovante: idComprovante,
      hash_comprovante: "",
      protocolo: protocolo,
      indice_comprovante: indiceComprovante,
      loja: loja4,
      empresa: empresa,
      tipo_documento: item.tipo_documento || "nao_identificado",
      data_deposito: item.data_deposito || "",
      valor_deposito: Number(item.valor_deposito || 0),
      banco: item.banco || "Nao identificado",
      data_movimento: dataMovimentoTexto,
      valores_diarios_movimento: valoresDiariosMovimento,
      houve_retirada: aplicarRetiradaNestaLinha,
      valor_retirada: aplicarRetiradaNestaLinha ? valorRetirada : 0,
      motivo_retirada: aplicarRetiradaNestaLinha ? motivoRetirada : "Não houve retirada",
      data_geracao_documento: item.data_geracao_documento || "",
      codigo_autenticacao: item.codigo_autenticacao || "",
      link_comprovante: linkComprovante,
      status_processamento: "PRONTO_PARA_SALVAR",
      pendencias: item.observacoes || [],
      divergencias: [],
      confianca_geral: Number(resultado.confianca_geral || 0)
    };

    payloadSalvar.hash_comprovante = smartSlipGerarHashComprovante(payloadSalvar);

    smartSlipSalvarDadosComprovantePlanilha_(payloadSalvar);
    if (aplicarRetiradaNestaLinha) {
        retiradaJaAplicada = true;
      }
    salvos++;
  });

  if (salvos > 0 && pendentes === 0) {
    return {
      ok: true,
      status_fila: "SALVO",
      mensagem: salvos + " comprovante(s) processado(s) e salvo(s) na BASE_SMARTSLIP.",
      qtd_salvos: salvos,
      qtd_pendentes: pendentes,
      resultado: resultado
    };
  }

  if (salvos > 0 && pendentes > 0) {
    return {
      ok: true,
      status_fila: "SALVO_PARCIAL",
      mensagem: salvos + " comprovante(s) salvo(s). " + pendentes + " comprovante(s) com pendência.",
      qtd_salvos: salvos,
      qtd_pendentes: pendentes,
      pendencias: mensagens,
      resultado: resultado
    };
  }

  return {
    ok: false,
    status_fila: "PENDENTE_INTERNO",
    mensagem: "Nenhum comprovante foi salvo. " + mensagens.join(" | "),
    qtd_salvos: salvos,
    qtd_pendentes: pendentes,
    pendencias: mensagens,
    resultado: resultado
  };
}

function smartSlipValidarBloqueiosFila(resultado, dados) {
  const pendencias = [];

  if (!dados.loja) pendencias.push("Loja não informada.");
  if (!dados.empresa || dados.empresa === "Nao identificado") pendencias.push("Empresa não identificada pela Info_limites.");
  if (!resultado.tipo_documento || resultado.tipo_documento === "nao_identificado") pendencias.push("Tipo de documento não identificado.");
  if (!dados.data_deposito) pendencias.push("Data do depósito não identificada.");
  if (dados.valor_deposito === null || dados.valor_deposito === undefined || dados.valor_deposito === "") pendencias.push("Valor do depósito não identificado.");
  if (!dados.data_movimento) pendencias.push("Data de movimento não informada.");
  if (dados.houve_retirada === null || dados.houve_retirada === undefined) pendencias.push("Informação sobre retirada não informada.");
  if (!dados.data_geracao_documento) pendencias.push("Data de geração do documento não identificada.");

  if (dados.houve_retirada === true) {
    if (!dados.valor_retirada) pendencias.push("Valor da retirada não informado.");
    if (!dados.motivo_retirada) pendencias.push("Motivo da retirada não informado.");
  }

  return pendencias;
}

function smartSlipInstalarTriggerFila() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "smartSlipProcessarFila") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("smartSlipProcessarFila")
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log("Trigger smartSlipProcessarFila instalado para rodar a cada minuto.");
}

function smartSlipExtrairListaComprovantes(resultado) {
  resultado = resultado || {};

  if (Array.isArray(resultado.comprovantes) && resultado.comprovantes.length > 0) {
    return resultado.comprovantes;
  }

  // Compatibilidade com o formato antigo, de 1 comprovante só.
  if (resultado.dados_comprovante) {
    return [{
      status: resultado.status || "PRONTO_PARA_SALVAR",
      indice_comprovante: 1,
      tipo_documento: resultado.tipo_documento || "nao_identificado",
      data_deposito: resultado.dados_comprovante.data_deposito || null,
      valor_deposito: resultado.dados_comprovante.valor_deposito || null,
      banco: resultado.dados_comprovante.banco || "Nao identificado",
      data_geracao_documento: resultado.dados_comprovante.data_geracao_documento || null,
      codigo_autenticacao: resultado.dados_comprovante.codigo_autenticacao || "",
      observacoes: resultado.pendencias || []
    }];
  }

  return [];
}

function smartSlipExtrairComprovantesDoResultado(resultado) {
  resultado = resultado || {};

  if (Array.isArray(resultado.comprovantes) && resultado.comprovantes.length > 0) {
    return resultado.comprovantes;
  }

  if (resultado.dados_comprovante) {
    return [{
      status: resultado.status || "PRONTO_PARA_SALVAR",
      indice_comprovante: 1,
      tipo_documento: resultado.tipo_documento || "nao_identificado",
      data_deposito: resultado.dados_comprovante.data_deposito || null,
      valor_deposito: resultado.dados_comprovante.valor_deposito || null,
      banco: resultado.dados_comprovante.banco || "Nao identificado",
      data_geracao_documento: resultado.dados_comprovante.data_geracao_documento || null,
      codigo_autenticacao: resultado.dados_comprovante.codigo_autenticacao || "",
      observacoes: resultado.pendencias || []
    }];
  }

  return [];
}

function smartSlipItemTemInferenciaOuBaixaEvidencia_(item) {
  item = item || {};

  const obs = Array.isArray(item.observacoes)
    ? item.observacoes.join(" | ")
    : String(item.observacoes || "");

  const obsNorm = obs
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  const evid = item.evidencias || {};

  const valorTexto = String(evid.valor_deposito_texto || "").trim();
  const blocoTexto = String(evid.bloco_textual_comprovante || "").trim();

  const valor = Number(item.valor_deposito || 0);

  const temTextoInferencia =
    obsNorm.includes("inferid") ||
    obsNorm.includes("parcialmente visivel") ||
    obsNorm.includes("parcialmente visíveis") ||
    obsNorm.includes("trechos parcialmente") ||
    obsNorm.includes("nao esta claro") ||
    obsNorm.includes("não está claro") ||
    obsNorm.includes("duvida") ||
    obsNorm.includes("dúvida");

  if (temTextoInferencia) {
    return {
      bloquear: true,
      motivo: "Item possui observação de inferência ou leitura parcial."
    };
  }

  if (!valor || valor <= 0) {
    return {
      bloquear: true,
      motivo: "Valor do depósito ausente ou inválido."
    };
  }

  if (!valorTexto) {
    return {
      bloquear: true,
      motivo: "Sem evidência textual do valor do depósito."
    };
  }

  if (!blocoTexto || blocoTexto.length < 40) {
    return {
      bloquear: true,
      motivo: "Sem bloco textual suficiente do comprovante."
    };
  }

  return {
    bloquear: false,
    motivo: ""
  };
}

function smartSlipValidarDuplicidadeEvidenciaLote_(comprovantes) {
  comprovantes = Array.isArray(comprovantes) ? comprovantes : [];

  const mapaBloco = {};
  const mapaPosicaoValor = {};
  const duplicidades = [];
  const repeticoesSemEvidenciaDistinta = [];

  comprovantes.forEach(function(item) {
    item = item || {};

    const evid = item.evidencias || {};
    const indice = item.indice_comprovante || "?";

    const valorCentavos = smartSlipValorCentavos_(item.valor_deposito);
    const dataDeposito = smartSlipTextoChave_(item.data_deposito || "");
    const banco = smartSlipTextoChave_(item.banco || "");
    const tipo = smartSlipTextoChave_(item.tipo_documento || "");

    const posicao = smartSlipTextoChave_(item.posicao_visual_aproximada || "");
    const valorTexto = smartSlipTextoChave_(evid.valor_deposito_texto || "");
    const blocoTexto = smartSlipTextoChave_(evid.bloco_textual_comprovante || "");
    const controleTexto = smartSlipTextoChave_(
      evid.codigo_autenticacao_texto ||
      evid.controle_texto ||
      item.codigo_autenticacao ||
      ""
    );

    /*
      1) Mesmo valor NÃO é duplicidade.
      A duplicidade só é considerada quando a IA usa o mesmo bloco/evidência
      para mais de um comprovante.
    */

    if (blocoTexto && blocoTexto.length >= 40) {
      const chaveBloco = [
        tipo,
        banco,
        dataDeposito,
        valorCentavos,
        blocoTexto.substring(0, 320)
      ].join("|");

      if (!mapaBloco[chaveBloco]) {
        mapaBloco[chaveBloco] = [];
      }

      mapaBloco[chaveBloco].push(indice);
    }

    /*
      2) Se a IA informou a mesma posição visual para o mesmo valor/data/banco,
      isso é suspeito, porque dois comprovantes reais iguais deveriam aparecer
      em posições diferentes na foto.
    */
    if (posicao && valorCentavos) {
      const chavePosicaoValor = [
        tipo,
        banco,
        dataDeposito,
        valorCentavos,
        posicao
      ].join("|");

      if (!mapaPosicaoValor[chavePosicaoValor]) {
        mapaPosicaoValor[chavePosicaoValor] = [];
      }

      mapaPosicaoValor[chavePosicaoValor].push(indice);
    }

    /*
      3) Se existe valor repetido, mas a IA não entregou evidência textual mínima,
      não é bloqueio por valor repetido; é bloqueio por falta de evidência distinta.
      A validação individual também deve pegar isso, mas aqui reforçamos no lote.
    */
    if (valorCentavos && (!valorTexto || !blocoTexto || blocoTexto.length < 40)) {
      repeticoesSemEvidenciaDistinta.push({
        indice: indice,
        valor: item.valor_deposito,
        motivo: "valor informado sem evidência textual/bloco OCR suficiente"
      });
    }
  });

  Object.keys(mapaBloco).forEach(function(chave) {
    if (mapaBloco[chave].length > 1) {
      duplicidades.push({
        tipo: "MESMO_BLOCO_TEXTUAL",
        indices: mapaBloco[chave]
      });
    }
  });

  Object.keys(mapaPosicaoValor).forEach(function(chave) {
    if (mapaPosicaoValor[chave].length > 1) {
      duplicidades.push({
        tipo: "MESMA_POSICAO_VISUAL_MESMO_VALOR",
        indices: mapaPosicaoValor[chave]
      });
    }
  });

  if (duplicidades.length) {
    return {
      ok: false,
      motivo: "DUPLICIDADE_EVIDENCIA_DOCUMENTAL",
      mensagem:
        "Validação anti-alucinação bloqueou o lote. A IA retornou comprovantes usando a mesma evidência documental ou a mesma posição visual para o mesmo valor. Índices envolvidos: " +
        duplicidades.map(function(d) {
          return d.indices.join(", ");
        }).join(" | "),
      duplicidades: duplicidades
    };
  }

  return {
    ok: true,
    motivo: "",
    mensagem: "",
    repeticoes_sem_evidencia_distinta: repeticoesSemEvidenciaDistinta
  };
}

function smartSlipValorCentavos_(valor) {
  if (valor === null || valor === undefined || valor === "") {
    return 0;
  }

  let n = 0;

  if (typeof valor === "number") {
    n = valor;
  } else {
    const txt = String(valor)
      .replace(/[R$\s]/g, "")
      .replace(/\./g, "")
      .replace(",", ".")
      .trim();

    n = Number(txt);
  }

  if (isNaN(n)) {
    return 0;
  }

  return Math.round(n * 100);
}

function smartSlipTextoChave_(valor) {
  return String(valor || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .replace(/[^\w\s.,:/-]/g, "")
    .trim();
}

function smartSlipGerarCodigoUnicoComprovante(protocolo, indiceComprovante) {
  const indice = String(indiceComprovante || 1).padStart(3, "0");
  const base = String(protocolo || "SS-SEM-PROTOCOLO").trim();

  return base + "-C" + indice;
}

function smartSlipGerarIdComprovante(protocolo, indiceComprovante) {
  const protocoloLimpo = String(protocolo || "SS-SEM-PROTOCOLO").trim();
  const indice = String(indiceComprovante || 1).padStart(3, "0");

  return protocoloLimpo + "-C" + indice;
}

function smartSlipGerarHashComprovante(dados) {
  const base = [
    dados.protocolo || "",
    dados.indice_comprovante || "",
    dados.loja || "",
    dados.empresa || "",
    dados.tipo_documento || "",
    dados.data_deposito || "",
    dados.valor_deposito || "",
    dados.banco || "",
    dados.data_movimento || "",
    dados.link_comprovante || ""
  ].join("|");

  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    base,
    Utilities.Charset.UTF_8
  );

  return bytes.map(function(b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

function smartSlipObterUltimaLojaUsuario(email) {
  email = String(email || "").trim().toLowerCase();

  if (!email) return "";

  const ss = SpreadsheetApp.openById(SMARTSLIP_DB_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SMARTSLIP_ABA_FILA);

  if (!sh || sh.getLastRow() < 2) {
    return "";
  }

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];

    const emailRow = String(row[2] || "").trim().toLowerCase();
    const loja = String(row[3] || "").trim();

    if (emailRow === email && loja) {
      return loja;
    }
  }

  return "";
}

function smartSlipGetBootstrapApp() {
  try {
    const usuario = smartSlipGetUsuarioAtual();

    let resumoMensal = [];
    let historico = [];
    let resumoDiario = [];
    let resumoStatus = [];

    try {
      resumoMensal = smartSlipGetResumoMensal(usuario);
    } catch (errResumo) {
      Logger.log("Erro ao carregar resumo mensal: " + String(errResumo && errResumo.message ? errResumo.message : errResumo));
      resumoMensal = [];
    }

    try {
      resumoDiario = smartSlipGetResumoDiario(usuario);
    } catch (errDiario) {
      Logger.log("Erro ao carregar resumo diário: " + String(errDiario && errDiario.message ? errDiario.message : errDiario));
      resumoDiario = [];
    }

    try {
      historico = smartSlipGetHistorico(usuario);
    } catch (errHist) {
      Logger.log("Erro ao carregar histórico: " + String(errHist && errHist.message ? errHist.message : errHist));
      historico = [];
    }

    try {
      resumoStatus = smartSlipGetResumoStatus(usuario);
    } catch (errStatus) {
      Logger.log("Erro ao carregar resumo de status: " + String(errStatus && errStatus.message ? errStatus.message : errStatus));
      resumoStatus = [];
    }

    return {
      ok: true,
      usuario: usuario,
      lojas: smartSlipListarLojasInfoLimites(usuario),
      resumo_mensal: resumoMensal,
      resumo_diario: resumoDiario,
      resumo_status: resumoStatus,
      historico: historico
    };

  } catch (err) {
    Logger.log("Erro smartSlipGetBootstrapApp: " + String(err && err.message ? err.message : err));

    return {
      ok: true,
      erro_bootstrap: String(err && err.message ? err.message : err),
      usuario: {
        email: "",
        primeiro_nome: "Usuário",
        is_admin: false,
        loja_padrao: ""
      },
      lojas: [],
      resumo_mensal: [],
      resumo_diario: [],
      resumo_status: [],
      historico: []
    };
  }
}

function smartSlipDiagnosticoBootstrap() {
  const resp = smartSlipGetBootstrapApp();
  Logger.log(JSON.stringify(resp, null, 2));
  return resp;
}

function smartSlipGetBootstrapAppJson() {
  try {
    const resp = smartSlipGetBootstrapApp();

    return JSON.stringify({
      ok: true,
      payload: resp
    });

  } catch (err) {
    return JSON.stringify({
      ok: false,
      erro: String(err && err.message ? err.message : err)
    });
  }
}

function smartSlipGetCompHubJson(optionsJson) {
  try {
    const usuario = smartSlipGetUsuarioAtual();

    if (!usuario.can_comp_hub) {
      return JSON.stringify({
        ok: false,
        erro: "Acesso restrito. O Comp Hub está disponível apenas para Administrador ou Analista Pro."
      });
    }

    let options = {};

    try {
      options = optionsJson ? JSON.parse(optionsJson) : {};
    } catch (e) {
      options = {};
    }

    const dias = Number(options.dias || 60);

    const payload = smartSlipGetCompHub(usuario, {
      dias: dias
    });

    return JSON.stringify({
      ok: true,
      payload: payload
    });

  } catch (err) {
    return JSON.stringify({
      ok: false,
      erro: String(err && err.message ? err.message : err)
    });
  }
}

function smartSlipExtrairValoresDiariosCompHub_(jsonOriginal) {
  if (!jsonOriginal) {
    return null;
  }

  try {
    const obj = JSON.parse(String(jsonOriginal || "{}"));
    return obj.valores_diarios_movimento || null;
  } catch (err) {
    return null;
  }
}

function smartSlipMontarMapaFilaPorProtocolo_(ss) {
  const mapa = {};

  const shFila = ss.getSheetByName(SMARTSLIP_ABA_FILA);

  if (!shFila || shFila.getLastRow() < 2) {
    return mapa;
  }

  const values = shFila.getRange(2, 1, shFila.getLastRow() - 1, 18).getValues();

  values.forEach(function(row) {
    const protocolo = String(row[1] || "").trim();

    if (!protocolo) {
      return;
    }

    mapa[protocolo] = {
      dataRegistroFila: smartSlipFormatarDataHora(row[0]),
      emailUsuarioFila: String(row[2] || ""),
      lojaInformadaFila: String(row[3] || ""),
      statusFila: String(row[13] || "").trim(),
      mensagemFila: String(row[14] || "").trim(),
      dataProcessamentoFila: smartSlipFormatarDataHora(row[17])
    };
  });

  return mapa;
}

function smartSlipGetCompHub(usuario, options) {
  usuario = usuario || smartSlipGetUsuarioAtual();
  options = options || {};

  const dias = Number(options.dias || 60);

  const limiteData = new Date();
  limiteData.setDate(limiteData.getDate() - dias);
  limiteData.setHours(0, 0, 0, 0);

  if (!usuario.can_comp_hub) {
    throw new Error("Acesso restrito ao Comp Hub.");
  }

  const ss = SpreadsheetApp.openById(SMARTSLIP_DB_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SMARTSLIP_ABA_DESTINO);
  const mapaFilaPorProtocolo = smartSlipMontarMapaFilaPorProtocolo_(ss);

  if (!sh || sh.getLastRow() < 2) {
    return {
      rows: [],
      atualizado_em: smartSlipFormatarDataHora(new Date()),
      range_dias: dias
    };
  }

  const values = sh.getDataRange().getValues();

  const headers = values[0].map(function(h) {
    return String(h || "").trim();
  });

  const linhas = [];

  for (let i = values.length - 1; i >= 1; i--) {
    const row = values[i];

    const dataRegistroRaw = smartSlipGetValorHeaderCompHub(row, headers, ["Data Registro"]);
    const dataRegistroObj = smartSlipConverterDataCompHub_(dataRegistroRaw);

    if (dataRegistroObj && dataRegistroObj < limiteData) {
      continue;
    }

    const protocolo = String(
      smartSlipGetValorHeaderCompHub(row, headers, ["Protocolo"]) || ""
    ).trim();

    const filaInfo = mapaFilaPorProtocolo[protocolo] || {};

    const statusBase = String(
      smartSlipGetValorHeaderCompHub(row, headers, ["Status Processamento"]) || ""
    ).trim();

    const statusOperacional = String(
      filaInfo.statusFila || statusBase || ""
    ).trim();

    const jsonOriginalRaw = String(
  smartSlipGetValorHeaderCompHub(row, headers, ["JSON Original"]) || ""
);

const valoresDiarios = smartSlipExtrairValoresDiariosCompHub_(jsonOriginalRaw);

    const item = {
      dataRegistro: smartSlipFormatarDataHora(
        smartSlipGetValorHeaderCompHub(row, headers, ["Data Registro"])
      ),

      dataRegistroInput: smartSlipFormatarDataInputCompHub(
        smartSlipGetValorHeaderCompHub(row, headers, ["Data Registro"])
      ),

      idComprovante: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["ID Comprovante"]) || ""
      ),

      hashComprovante: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Hash Comprovante"]) || ""
      ),

      protocolo: protocolo,

      indiceComprovante: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Índice Comprovante", "Indice Comprovante"]) || ""
      ),

      loja: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Loja"]) || ""
      ),

      empresa: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Empresa"]) || ""
      ),

      tipoDocumento: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Tipo Documento"]) || ""
      ),

      dataDeposito: smartSlipFormatarDataMovimentoHistorico(
        smartSlipGetValorHeaderCompHub(row, headers, ["Data Depósito", "Data Deposito"])
      ),

      valorDeposito: smartSlipNumeroCompHub(
        smartSlipGetValorHeaderCompHub(row, headers, ["Valor Depósito", "Valor Deposito"])
      ),

      banco: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Banco"]) || ""
      ),

      dataMovimento: smartSlipFormatarDataMovimentoHistorico(
        smartSlipGetValorHeaderCompHub(row, headers, ["Data Movimento"])
      ),

      houveRetirada: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Houve Retirada"]) || ""
      ),

      valorRetirada: smartSlipNumeroCompHub(
        smartSlipGetValorHeaderCompHub(row, headers, ["Valor Retirada"])
      ),

      motivoRetirada: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Motivo Retirada"]) || ""
      ),

      dataGeracaoDocumento: smartSlipFormatarDataMovimentoHistorico(
        smartSlipGetValorHeaderCompHub(row, headers, ["Data Geração Documento", "Data Geracao Documento"])
      ),

      codigoAutenticacao: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Código Autenticação", "Codigo Autenticacao"]) || ""
      ),

      linkComprovante: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Link Comprovante"]) || ""
      ),

      // Status operacional vindo da SMARTSLIP_FILA.
      // Se não encontrar protocolo na fila, usa o status da BASE_SMARTSLIP como fallback.
      statusProcessamento: statusOperacional,
      statusBaseProcessamento: statusBase,
      statusFila: filaInfo.statusFila || "",
      mensagemFila: filaInfo.mensagemFila || "",
      dataProcessamentoFila: filaInfo.dataProcessamentoFila || "",

      pendencias: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Pendências", "Pendencias"]) || ""
      ),

      divergencias: String(
        smartSlipGetValorHeaderCompHub(row, headers, ["Divergências", "Divergencias"]) || ""
      ),

      confiancaGeral: smartSlipNumeroCompHub(
        smartSlipGetValorHeaderCompHub(row, headers, ["Confiança Geral", "Confianca Geral"])
      ),

      valoresDiarios: valoresDiarios,
      jsonOriginal: jsonOriginalRaw
    };

    linhas.push(item);
  }

  return {
    rows: linhas,
    atualizado_em: smartSlipFormatarDataHora(new Date()),
    range_dias: dias
  };
}

function smartSlipGetValorHeaderCompHub(row, headers, nomes) {
  const idx = smartSlipEncontrarIndiceHeader_(headers, nomes);

  if (idx < 0) {
    return "";
  }

  return row[idx];
}

function smartSlipNumeroCompHub(valor) {
  if (valor === null || valor === undefined || valor === "") {
    return 0;
  }

  if (typeof valor === "number") {
    return valor;
  }

  const txt = String(valor)
    .trim()
    .replace(/\./g, "")
    .replace(",", ".");

  const n = Number(txt);

  return isNaN(n) ? 0 : n;
}

function smartSlipFormatarDataInputCompHub(valor) {
  if (!valor) return "";

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    return Utilities.formatDate(valor, tz, "yyyy-MM-dd");
  }

  const txt = String(valor).trim();

  let m = txt.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) {
    return m[3] + "-" + m[2] + "-" + m[1];
  }

  m = txt.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    return m[1] + "-" + m[2] + "-" + m[3];
  }

  return "";
}

function smartSlipConverterDataCompHub_(valor) {
  if (!valor) return null;

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    return valor;
  }

  const txt = String(valor).trim();

  let m = txt.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) {
    return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  }

  m = txt.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  return null;
}

function smartSlipGetHistorico(usuario) {
  usuario = usuario || smartSlipGetUsuarioAtual();

  const ss = SpreadsheetApp.openById(SMARTSLIP_DB_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SMARTSLIP_ABA_FILA);

  if (!sh || sh.getLastRow() < 2) {
    return [];
  }

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();
  const historico = [];

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];

    const emailRow = String(row[2] || "").trim().toLowerCase();

    if (!usuario.is_admin && emailRow !== usuario.email) {
      continue;
    }

    const statusFila = String(row[13] || "").trim();
    const camposReenvio = smartSlipMontarCamposReenvioHistorico_(row);

    historico.push({
      data_registro: smartSlipFormatarDataHora(row[0]),
      protocolo: row[1] || "",
      email_usuario: row[2] || "",
      loja: row[3] || "",
      data_movimento: smartSlipFormatarDataMovimentoHistorico(row[5]),
      link_comprovante: row[11] || "",
      nome_arquivo: row[12] || "",
      status_fila: statusFila,
      mensagem: row[14] || "",

      // Controle do botão REENVIAR?
      reenviar_permitido: smartSlipStatusPermiteReenvio_(statusFila),

      // Campos usados para preencher automaticamente a tela de Envio
      campos_reenvio: camposReenvio
    });

    if (historico.length >= 50) break;
  }

  return historico;
}

function smartSlipStatusPermiteReenvio_(status) {
  status = String(status || "").trim().toUpperCase();

  return [
    "PENDENTE_INTERNO",
    "SALVO_PARCIAL",
    "ERRO_PROCESSAMENTO",
    "INELEGIVEL",
    "DIVERGENCIA",
    "PRECISA_COMPLEMENTO",
    "PENDENCIA"
  ].includes(status);
}

function smartSlipMontarCamposReenvioHistorico_(row) {
  row = row || [];

  let respostas = {};

  try {
    respostas = JSON.parse(String(row[15] || "{}"));
  } catch (err) {
    respostas = {};
  }

  const movimentoRaw = row[5];

  const inicioRaw =
    respostas.data_movimento_inicio ||
    respostas.data_movimento ||
    smartSlipExtrairInicioMovimentoHistorico_(movimentoRaw);

  const fimRaw =
    respostas.data_movimento_fim ||
    respostas.data_movimento ||
    smartSlipExtrairFimMovimentoHistorico_(movimentoRaw) ||
    inicioRaw;

  const houveRetirada = typeof respostas.houve_retirada === "boolean"
    ? respostas.houve_retirada
    : String(row[6] || "").trim().toLowerCase() === "sim";

  const valorRetirada = respostas.valor_retirada !== undefined
    ? Number(respostas.valor_retirada || 0)
    : Number(row[7] || 0);

  const motivoRetirada = respostas.motivo_retirada !== undefined
    ? String(respostas.motivo_retirada || "").trim()
    : String(row[8] || "").trim();

  return {
    loja: smartSlipNormalizarLoja4(respostas.loja || row[3] || ""),
    data_movimento_inicio_input: smartSlipConverterDataParaInputHistorico_(inicioRaw),
    data_movimento_fim_input: smartSlipConverterDataParaInputHistorico_(fimRaw),

    houve_retirada: houveRetirada,
    valor_retirada: valorRetirada,
    motivo_retirada: houveRetirada
      ? motivoRetirada
      : "Não houve retirada",

    // Novo: valores informados no envio original.
    // Para envios antigos que ainda não tinham essa estrutura, ficará null.
    valores_diarios_movimento: respostas.valores_diarios_movimento || null
  };
}

function smartSlipExtrairInicioMovimentoHistorico_(valor) {
  if (!valor) return "";

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    return valor;
  }

  const txt = String(valor || "").trim();

  if (txt.indexOf(" a ") > -1) {
    return txt.split(" a ")[0].trim();
  }

  return txt;
}

function smartSlipExtrairFimMovimentoHistorico_(valor) {
  if (!valor) return "";

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    return valor;
  }

  const txt = String(valor || "").trim();

  if (txt.indexOf(" a ") > -1) {
    return txt.split(" a ")[1].trim();
  }

  return txt;
}

function smartSlipConverterDataParaInputHistorico_(valor) {
  if (!valor) return "";

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    return Utilities.formatDate(valor, tz, "yyyy-MM-dd");
  }

  const txt = String(valor || "").trim();

  let m = txt.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    return m[3] + "-" + m[2] + "-" + m[1];
  }

  m = txt.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    return m[1] + "-" + m[2] + "-" + m[3];
  }

  return "";
}

function smartSlipGetResumoStatus(usuario) {
  usuario = usuario || smartSlipGetUsuarioAtual();

  const ss = SpreadsheetApp.openById(SMARTSLIP_DB_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SMARTSLIP_ABA_FILA);

  if (!sh || sh.getLastRow() < 2) {
    return [];
  }

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();
  const mapa = {};
  let total = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const emailRow = String(row[2] || "").trim().toLowerCase();

    if (!usuario.is_admin && emailRow !== usuario.email) {
      continue;
    }

    const status = String(row[13] || "").trim() || "SEM_STATUS";

    mapa[status] = (mapa[status] || 0) + 1;
    total++;
  }

  return Object.keys(mapa).sort().map(function(status) {
    const quantidade = mapa[status] || 0;

    return {
      status: status,
      quantidade: quantidade,
      percentual: total > 0
        ? Number(((quantidade / total) * 100).toFixed(2))
        : 0
    };
  });
}

function smartSlipFormatarDataHora(valor) {
  if (!valor) return "";

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    return Utilities.formatDate(valor, tz, "dd/MM/yyyy HH:mm:ss");
  }

  return String(valor);
}

function smartSlipFormatarDataMovimentoHistorico(valor) {
  if (!valor) return "";

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    return Utilities.formatDate(valor, tz, "dd/MM/yyyy");
  }

  const txt = String(valor).trim();

  if (!txt) return "";

  if (txt.indexOf(" a ") > -1) {
    return txt
      .split(" a ")
      .map(function(parte) {
        return smartSlipFormatarDataMovimentoItem(parte);
      })
      .join(" a ");
  }

  return smartSlipFormatarDataMovimentoItem(txt);
}

function smartSlipFormatarDataMovimentoItem(valor) {
  const txt = String(valor || "").trim();

  if (!txt) return "";

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(txt)) {
    return txt;
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(txt) || /^\d{4}-\d{2}-\d{2}T/.test(txt)) {
    const d = new Date(txt);

    if (!isNaN(d.getTime())) {
      const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
      return Utilities.formatDate(d, tz, "dd/MM/yyyy");
    }
  }

  return txt;
}

function smartSlipGetResumoMensal(usuario) {
  try {
    usuario = usuario || smartSlipGetUsuarioAtual();

    const ss = SpreadsheetApp.openById(SMARTSLIP_DB_SPREADSHEET_ID);
    const shBase = ss.getSheetByName(SMARTSLIP_ABA_DESTINO);
    const shFila = ss.getSheetByName(SMARTSLIP_ABA_FILA);

    if (!shBase || shBase.getLastRow() < 2 || !shFila || shFila.getLastRow() < 2) {
      return [];
    }

    const filaValues = shFila.getRange(2, 1, shFila.getLastRow() - 1, 18).getValues();
    const protocolosPermitidos = {};

    filaValues.forEach(function(row) {
      const protocolo = String(row[1] || "").trim();
      const emailRow = String(row[2] || "").trim().toLowerCase();

      if (!protocolo) return;

      if (usuario.is_admin || emailRow === usuario.email) {
        protocolosPermitidos[protocolo] = true;
      }
    });

    const baseValues = shBase.getDataRange().getValues();

    if (baseValues.length < 2) {
      return [];
    }

    const headers = baseValues[0].map(function(h) {
      return String(h || "").trim();
    });

    const idxDataRegistro = smartSlipEncontrarIndiceHeader_(headers, ["Data Registro"]);
    const idxProtocolo = smartSlipEncontrarIndiceHeader_(headers, ["Protocolo"]);
    const idxLoja = smartSlipEncontrarIndiceHeader_(headers, ["Loja"]);
    const idxValor = smartSlipEncontrarIndiceHeader_(headers, ["Valor Depósito"]);

    if (idxDataRegistro < 0 || idxProtocolo < 0 || idxValor < 0) {
      Logger.log("Resumo mensal: colunas necessárias não encontradas.");
      Logger.log(JSON.stringify(headers));
      return [];
    }

    const mapa = {};

    for (let i = 1; i < baseValues.length; i++) {
      const row = baseValues[i];

      const protocolo = String(row[idxProtocolo] || "").trim();

      if (!protocolosPermitidos[protocolo]) {
        continue;
      }

      const loja = idxLoja >= 0 ? String(row[idxLoja] || "").trim() : "";

      if (!usuario.is_admin && usuario.loja_padrao && loja !== usuario.loja_padrao) {
        continue;
      }

      const mes = smartSlipMesAno(row[idxDataRegistro]);
      const valor = smartSlipNumero(row[idxValor]);

      if (!mapa[mes]) {
        mapa[mes] = {
          mes: mes,
          quantidade: 0,
          valor_total: 0
        };
      }

      mapa[mes].quantidade++;
      mapa[mes].valor_total += valor;
    }

    return Object.keys(mapa)
      .sort()
      .map(function(k) {
        return mapa[k];
      });

  } catch (err) {
    Logger.log("Erro smartSlipGetResumoMensal: " + String(err && err.message ? err.message : err));
    return [];
  }
}

function smartSlipGetResumoDiario(usuario) {
  try {
    usuario = usuario || smartSlipGetUsuarioAtual();

    const ss = SpreadsheetApp.openById(SMARTSLIP_DB_SPREADSHEET_ID);
    const shBase = ss.getSheetByName(SMARTSLIP_ABA_DESTINO);
    const shFila = ss.getSheetByName(SMARTSLIP_ABA_FILA);

    if (!shBase || shBase.getLastRow() < 2 || !shFila || shFila.getLastRow() < 2) {
      return [];
    }

    const filaValues = shFila.getRange(2, 1, shFila.getLastRow() - 1, 18).getValues();
    const protocolosPermitidos = {};

    filaValues.forEach(function(row) {
      const protocolo = String(row[1] || "").trim();
      const emailRow = String(row[2] || "").trim().toLowerCase();

      if (!protocolo) return;

      // Admin vê todos. Usuário normal vê apenas os envios dele.
      if (usuario.is_admin || emailRow === usuario.email) {
        protocolosPermitidos[protocolo] = true;
      }
    });

    const baseValues = shBase.getDataRange().getValues();

    if (baseValues.length < 2) {
      return [];
    }

    const headers = baseValues[0].map(function(h) {
      return String(h || "").trim();
    });

    const idxDataRegistro = smartSlipEncontrarIndiceHeader_(headers, ["Data Registro"]);
    const idxProtocolo = smartSlipEncontrarIndiceHeader_(headers, ["Protocolo"]);
    const idxValor = smartSlipEncontrarIndiceHeader_(headers, ["Valor Depósito"]);

    if (idxDataRegistro < 0 || idxProtocolo < 0 || idxValor < 0) {
      Logger.log("Resumo diário: colunas necessárias não encontradas.");
      Logger.log(JSON.stringify(headers));
      return [];
    }

    const mapa = {};

    for (let i = 1; i < baseValues.length; i++) {
      const row = baseValues[i];

      const protocolo = String(row[idxProtocolo] || "").trim();

      if (!protocolosPermitidos[protocolo]) {
        continue;
      }

      const dia = smartSlipDiaIso(row[idxDataRegistro]);
      const valor = smartSlipNumero(row[idxValor]);

      if (!dia) {
        continue;
      }

      if (!mapa[dia]) {
        mapa[dia] = {
          dia: dia,
          quantidade: 0,
          valor_total: 0
        };
      }

      mapa[dia].quantidade++;
      mapa[dia].valor_total += valor;
    }

    return Object.keys(mapa)
      .sort()
      .map(function(k) {
        return mapa[k];
      });

  } catch (err) {
    Logger.log("Erro smartSlipGetResumoDiario: " + String(err && err.message ? err.message : err));
    return [];
  }
}

function smartSlipDiaIso(valor) {
  let d = valor;

  if (Object.prototype.toString.call(d) !== "[object Date]") {
    d = new Date(valor);
  }

  if (isNaN(d.getTime())) {
    return "";
  }

  const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  return Utilities.formatDate(d, tz, "yyyy-MM-dd");
}

function smartSlipMesAno(valor) {
  let d = valor;

  if (Object.prototype.toString.call(d) !== "[object Date]") {
    d = new Date(valor);
  }

  if (isNaN(d.getTime())) {
    return "Sem data";
  }

  const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  return Utilities.formatDate(d, tz, "yyyy-MM");
}

function smartSlipNumero(valor) {
  if (typeof valor === "number") return valor;

  const txt = String(valor || "")
    .replace(/[R$\s.]/g, "")
    .replace(",", ".");

  const n = Number(txt);

  return isNaN(n) ? 0 : n;
}

function smartSlipGarantirAbaUsuarios() {
  const ss = SpreadsheetApp.openById(SMARTSLIP_DB_SPREADSHEET_ID);
  let sh = ss.getSheetByName(SMARTSLIP_ABA_USUARIOS);

  if (!sh) {
    sh = ss.insertSheet(SMARTSLIP_ABA_USUARIOS);
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "Email",
      "Loja Padrao",
      "Empresa",
      "Data Atualizacao"
    ]);
  }

  return sh;
}

function smartSlipObterPreferenciasUsuario(email) {
  email = String(email || "").trim().toLowerCase();

  if (!email) {
    return {
      loja_padrao: "",
      empresa_padrao: ""
    };
  }

  const sh = smartSlipGarantirAbaUsuarios();
  const lastRow = sh.getLastRow();

  if (lastRow < 2) {
    return {
      loja_padrao: "",
      empresa_padrao: ""
    };
  }

  const values = sh.getRange(2, 1, lastRow - 1, 4).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const emailRow = String(row[0] || "").trim().toLowerCase();

    if (emailRow === email) {
      return {
        loja_padrao: String(row[1] || "").trim(),
        empresa_padrao: String(row[2] || "").trim()
      };
    }
  }

  return {
    loja_padrao: "",
    empresa_padrao: ""
  };
}

function smartSlipSalvarPreferenciasUsuario(lojaInformada) {
  const email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();

  if (!email) {
    throw new Error("Não foi possível identificar o e-mail do usuário.");
  }

  const infoLoja = smartSlipConsultarInfoLoja_(lojaInformada);

  if (!infoLoja.ok) {
    throw new Error(infoLoja.erro || "Loja não encontrada na Info_limites.");
  }

  const usuarioAtual = smartSlipGetUsuarioAtual();
  const empresaPermitida = usuarioAtual.is_admin
    ? ""
    : smartSlipObterEmpresaPermitidaPorEmail(usuarioAtual.email);

  if (empresaPermitida && infoLoja.empresa !== empresaPermitida) {
    throw new Error(
      "Essa loja pertence à empresa " +
      infoLoja.empresa +
      ", mas seu perfil permite apenas lojas da empresa " +
      empresaPermitida +
      "."
    );
  }

  const loja4 = infoLoja.loja_normalizada;
  const empresa = infoLoja.empresa;

  const sh = smartSlipGarantirAbaUsuarios();
  const lastRow = sh.getLastRow();

  if (lastRow >= 2) {
    const values = sh.getRange(2, 1, lastRow - 1, 4).getValues();

    for (let i = 0; i < values.length; i++) {
      const rowIndex = i + 2;
      const emailRow = String(values[i][0] || "").trim().toLowerCase();

      if (emailRow === email) {
        sh.getRange(rowIndex, 2).setValue(loja4);
        sh.getRange(rowIndex, 3).setValue(empresa);
        sh.getRange(rowIndex, 4).setValue(new Date());

        return {
          ok: true,
          email: email,
          loja_padrao: loja4,
          empresa_padrao: empresa
        };
      }
    }
  }

  sh.appendRow([
    email,
    loja4,
    empresa,
    new Date()
  ]);

  return {
    ok: true,
    email: email,
    loja_padrao: loja4,
    empresa_padrao: empresa
  };
}

function smartSlipSalvarPreferenciasUsuarioJson(lojaInformada) {
  try {
    return JSON.stringify({
      ok: true,
      payload: smartSlipSalvarPreferenciasUsuario(lojaInformada)
    });

  } catch (err) {
    return JSON.stringify({
      ok: false,
      erro: String(err && err.message ? err.message : err)
    });
  }
}

function smartSlipListarLojasInfoLimites(usuario) {
  usuario = usuario || smartSlipGetUsuarioAtual();

  const empresaPermitida = usuario.is_admin
    ? ""
    : smartSlipObterEmpresaPermitidaPorEmail(usuario.email);

  const ss = SpreadsheetApp.openById(SMARTSLIP_INFO_LIMITES_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SMARTSLIP_ABA_INFO_LIMITES);

  if (!sh) {
    throw new Error("Aba Info_limites não encontrada.");
  }

  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(function(h) {
    return String(h || "").trim();
  });

  const idxNormLoja = smartSlipEncontrarIndiceHeader_(headers, [
    "Norm_Loja",
    "NORM_LOJA",
    "Norm Loja",
    "Loja",
    "LOJA"
  ]);

  let idxEmpresa = smartSlipEncontrarIndiceHeader_(headers, [
    "EMPRESA0",
    "EMPRESA",
    "Empresa"
  ]);

  if (idxEmpresa < 0) {
    idxEmpresa = 6;
  }

  if (idxNormLoja < 0) {
    throw new Error("Coluna Norm_Loja não encontrada na Info_limites.");
  }

  const mapa = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const lojaRaw = row[idxNormLoja];
    const lojaDigits = String(lojaRaw || "").replace(/\D/g, "");

    if (!lojaDigits) {
      continue;
    }

    const loja4 = lojaDigits.padStart(4, "0").slice(-4);
    const empresaRaw = String(row[idxEmpresa] || "").trim();
    const empresa = smartSlipNormalizarEmpresa_(empresaRaw);

    if (empresaPermitida && empresa !== empresaPermitida) {
      continue;
    }

    if (!mapa[loja4]) {
      mapa[loja4] = {
        loja: loja4,
        empresa: empresa,
        empresa_original: empresaRaw
      };
    }
  }

  return Object.keys(mapa)
    .sort()
    .map(function(k) {
      return mapa[k];
    });
}

function smartSlipObterEmpresaPermitidaPorEmail(email) {
  const txt = String(email || "").trim().toLowerCase();

  if (!txt) {
    return "";
  }

  if (
    txt.includes("@fisia") ||
    txt.includes(".fisia") ||
    txt.includes("+fisia") ||
    txt.includes("fisia")
  ) {
    return "Fisia";
  }

  if (
    txt.includes("@centauro") ||
    txt.includes(".centauro") ||
    txt.includes("+centauro") ||
    txt.includes("centauro")
  ) {
    return "Centauro";
  }

  return "";
}

/****************************************************
 * SMARTSLIP MOBILE - PATCH BACKEND PARA APPS SCRIPT
 *
 * Cole este bloco no Code.gs, abaixo das funções SmartSlip
 * que você já tem no projeto.
 *
 * Depende destas funções já existentes no seu backend:
 * - smartSlipChamarGemini(base64Arquivo, mimeType, respostasUsuario)
 * - smartSlipConsultarInfoLoja_(lojaInformada)
 * - smartSlipSalvarDadosComprovantePlanilha_(dados)
 * - smartSlipNormalizarLoja4(loja)
 * - smartSlipSafeErr_(err)
 ****************************************************/

function smartSlipMobileProcessarArquivo(payload) {
  try {
    payload = payload || {};

    var base64 = String(payload.base64 || "").trim();
    base64 = base64.replace(/^data:[^;]+;base64,/, "");

    if (!base64) {
      throw new Error("Arquivo não recebido em base64.");
    }

    var mimeType = String(payload.mimeType || "application/pdf").trim();
    var respostasUsuario = payload.respostasUsuario || {};

    respostasUsuario = smartSlipMobileNormalizarRespostasUsuario_(respostasUsuario);

    var resultado = smartSlipChamarGemini(base64, mimeType, respostasUsuario);
    resultado = resultado || {};
    resultado.nome_arquivo = String(payload.filename || "");

    smartSlipMobileAplicarRespostasUsuario_(resultado, respostasUsuario);
    smartSlipMobileConsultarEmpresaPorLoja_(resultado);
    smartSlipMobileNormalizarResultado_(resultado);

    return {
      ok: true,
      resultado: resultado,
      arquivo: {
        nome: String(payload.filename || ""),
        mimeType: mimeType
      }
    };

  } catch (err) {
    return {
      ok: false,
      erro: smartSlipSafeErr_(err)
    };
  }
}

function smartSlipMobileSalvarResultado(payload) {
  try {
    payload = payload || {};
    var resultado = payload.resultado || {};

    smartSlipMobileConsultarEmpresaPorLoja_(resultado);
    smartSlipMobileNormalizarResultado_(resultado);

    if (resultado.status !== "PRONTO_PARA_SALVAR") {
      return {
        ok: false,
        erro: "Resultado ainda não está pronto para salvar. Pendências: " + (resultado.pendencias || []).join(" | "),
        resultado: resultado
      };
    }

    var dados = resultado.dados_comprovante || {};

    var payloadSalvar = {
      loja: dados.loja || "",
      empresa: dados.empresa || "",
      tipo_documento: resultado.tipo_documento || "",
      data_deposito: dados.data_deposito || "",
      valor_deposito: Number(dados.valor_deposito || 0),
      banco: dados.banco || "",
      data_movimento: dados.data_movimento || "",
      houve_retirada: dados.houve_retirada === true,
      valor_retirada: Number(dados.valor_retirada || 0),
      motivo_retirada: dados.houve_retirada === true
        ? (dados.motivo_retirada || "")
        : "Não houve retirada",
      data_geracao_documento: dados.data_geracao_documento || "",
      codigo_autenticacao: dados.codigo_autenticacao || "",
      link_comprovante: dados.link_comprovante || payload.link_comprovante || "",
      status_processamento: "PRONTO_PARA_SALVAR",
      pendencias: [],
      divergencias: resultado.divergencias || [],
      confianca_geral: Number(resultado.confianca_geral || 0),
      nome_arquivo: resultado.nome_arquivo || ""
    };

    var retorno = smartSlipSalvarDadosComprovantePlanilha_(payloadSalvar);

    return {
      ok: true,
      retorno: retorno,
      payload_salvo: payloadSalvar
    };

  } catch (err) {
    return {
      ok: false,
      erro: smartSlipSafeErr_(err)
    };
  }
}

function smartSlipMobileNormalizarRespostasUsuario_(r) {
  r = r || {};

  if (r.loja) {
    r.loja = smartSlipNormalizarLoja4(r.loja);
  }

  if (r.houve_retirada === "sim") r.houve_retirada = true;
  if (r.houve_retirada === "nao") r.houve_retirada = false;
  if (r.houve_retirada === "") r.houve_retirada = null;

  if (r.valor_retirada !== undefined && r.valor_retirada !== null && r.valor_retirada !== "") {
    r.valor_retirada = Number(r.valor_retirada || 0);
  }

  return r;
}

function smartSlipMobileAplicarRespostasUsuario_(resultado, respostasUsuario) {
  resultado.dados_comprovante = resultado.dados_comprovante || {};
  var dados = resultado.dados_comprovante;

  if (respostasUsuario.loja) {
    dados.loja = respostasUsuario.loja;
  }

  if (!dados.banco && respostasUsuario.banco) {
    dados.banco = respostasUsuario.banco;
  }

  if (respostasUsuario.data_movimento) {
    dados.data_movimento = respostasUsuario.data_movimento;
  }

  if (typeof respostasUsuario.houve_retirada === "boolean") {
    dados.houve_retirada = respostasUsuario.houve_retirada;
  }

  if (dados.houve_retirada === false) {
    dados.valor_retirada = 0;
    dados.motivo_retirada = "Não houve retirada";
  }

  if (dados.houve_retirada === true) {
    if (respostasUsuario.valor_retirada !== undefined) {
      dados.valor_retirada = Number(respostasUsuario.valor_retirada || 0);
    }
    if (respostasUsuario.motivo_retirada) {
      dados.motivo_retirada = respostasUsuario.motivo_retirada;
    }
  }

  if (dados.codigo_autenticacao === null || dados.codigo_autenticacao === undefined) {
    dados.codigo_autenticacao = "";
  }

  if (dados.link_comprovante === null || dados.link_comprovante === undefined) {
    dados.link_comprovante = "";
  }
}

function smartSlipMobileConsultarEmpresaPorLoja_(resultado) {
  resultado = resultado || {};
  resultado.dados_comprovante = resultado.dados_comprovante || {};
  resultado.pendencias = Array.isArray(resultado.pendencias) ? resultado.pendencias : [];
  resultado.perguntas_complementares = Array.isArray(resultado.perguntas_complementares) ? resultado.perguntas_complementares : [];

  var dados = resultado.dados_comprovante;

  if (!dados.loja) return;

  var loja4 = smartSlipNormalizarLoja4(dados.loja);
  dados.loja = loja4;

  var info = smartSlipConsultarInfoLoja_(loja4);

  if (info && info.ok) {
    dados.loja = info.loja_normalizada || loja4;
    dados.empresa = info.empresa || "Nao identificado";
  } else {
    dados.empresa = "Nao identificado";
    smartSlipMobileAddUnico_(resultado.pendencias, "Loja não encontrada na Info_limites.");
    smartSlipMobileAddUnico_(resultado.perguntas_complementares, "Confirme o número correto da loja.");
  }
}

function smartSlipMobileNormalizarResultado_(resultado) {
  resultado = resultado || {};
  resultado.dados_comprovante = resultado.dados_comprovante || {};
  resultado.pendencias = [];
  resultado.perguntas_complementares = [];
  resultado.divergencias = Array.isArray(resultado.divergencias) ? resultado.divergencias : [];

  var dados = resultado.dados_comprovante;

  if (!dados.codigo_autenticacao) dados.codigo_autenticacao = "";
  if (!dados.link_comprovante) dados.link_comprovante = "";

  smartSlipMobilePerguntaSe_(resultado, !dados.loja, "Loja não informada.", "Qual o número da loja?");
  smartSlipMobilePerguntaSe_(resultado, !dados.empresa || dados.empresa === "Nao identificado", "Empresa não identificada pela Info_limites.", "Confirme o número da loja para consultar a empresa correta.");
  smartSlipMobilePerguntaSe_(resultado, !resultado.tipo_documento || resultado.tipo_documento === "nao_identificado", "Tipo de documento não identificado.", "Qual é o tipo do comprovante?");
  smartSlipMobilePerguntaSe_(resultado, !dados.data_deposito, "Data do depósito não identificada.", "Qual a data do depósito?");
  smartSlipMobilePerguntaSe_(resultado, dados.valor_deposito === null || dados.valor_deposito === undefined || dados.valor_deposito === "", "Valor do depósito não identificado.", "Qual o valor do depósito?");
  smartSlipMobilePerguntaSe_(resultado, !dados.banco, "Banco não identificado.", "Qual o banco do comprovante?");
  smartSlipMobilePerguntaSe_(resultado, !dados.data_movimento, "Data de movimento não informada.", "Qual a data do movimento desse depósito?");
  smartSlipMobilePerguntaSe_(resultado, dados.houve_retirada === null || dados.houve_retirada === undefined, "Informação sobre retirada não informada.", "Houve alguma retirada desse movimento para estorno, reembolso no ato ou algo do tipo?");
  smartSlipMobilePerguntaSe_(resultado, !dados.data_geracao_documento, "Data de geração do documento não identificada.", "Qual a data de geração do documento?");

  if (dados.houve_retirada === true) {
    smartSlipMobilePerguntaSe_(resultado, !dados.valor_retirada, "Valor da retirada não informado.", "Qual o valor da retirada?");
    smartSlipMobilePerguntaSe_(resultado, !dados.motivo_retirada, "Motivo da retirada não informado.", "Qual o motivo da retirada?");
  }

  if (dados.houve_retirada === false) {
    dados.valor_retirada = 0;
    dados.motivo_retirada = "Não houve retirada";
  }

  if (resultado.pendencias.length || resultado.perguntas_complementares.length) {
    resultado.status = "PRECISA_COMPLEMENTO";
  } else if (resultado.divergencias.length) {
    resultado.status = "DIVERGENCIA";
  } else if (resultado.status !== "INELEGIVEL") {
    resultado.status = "PRONTO_PARA_SALVAR";
  }

  return resultado;
}

function smartSlipMobilePerguntaSe_(resultado, condicao, pendencia, pergunta) {
  if (!condicao) return;
  smartSlipMobileAddUnico_(resultado.pendencias, pendencia);
  smartSlipMobileAddUnico_(resultado.perguntas_complementares, pergunta);
}

function smartSlipMobileAddUnico_(arr, value) {
  value = String(value || "").trim();
  if (!value) return;
  if (arr.indexOf(value) === -1) arr.push(value);
}
