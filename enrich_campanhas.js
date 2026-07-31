var XLSX = require('C:\\Users\\mar\\AppData\\Local\\Temp\\opencode\\node_modules\\xlsx');
var fs = require('fs');

var BASE_DIR = 'C:\\Users\\mar\\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\\Área de Trabalho\\';
var CAM_DIR = BASE_DIR + 'dashboards\\campanhas\\';

// ============================================================
// 1. Ler base_8026_2026.xlsx → indexar por CODCLI
// ============================================================
console.log('Lendo base_8026_2026.xlsx...');
var baseWb = XLSX.readFile(BASE_DIR + '_bases\\base_8026_2026.xlsx');
var baseWs = baseWb.Sheets['Plan1'];
var baseRaw = XLSX.utils.sheet_to_json(baseWs, { header: 1, defval: '' });

// Excel serial epoch: Jan 1 1900 = 1 (with the leap year bug, Feb 29 1900 = 60)
function serialToDate(serial) {
  if (typeof serial !== 'number' || serial < 1) return null;
  // Excel date serial: days since Jan 0 1900 (Dec 31 1899)
  var epoch = new Date(1899, 11, 30); // Dec 30, 1899
  var d = new Date(epoch.getTime() + serial * 86400000);
  return d;
}

function fmtDate(d) {
  if (!d) return '';
  var dd = String(d.getUTCDate()).padStart(2, '0');
  var mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  var yyyy = d.getUTCFullYear();
  return dd + '/' + mm + '/' + yyyy;
}

// col 2 = DATA (serial), col 11 = CODCLI, col 34 = FATURAMENTO, col 9 = TOTAL, col 10 = VLFRETE
var codcliIndex = {};
var baseCount = 0;
for (var i = 1; i < baseRaw.length; i++) {
  var r = baseRaw[i];
  var codcli = String(r[11]).trim();
  if (!codcli || codcli === '0' || codcli === 'undefined') continue;
  var fat = parseFloat(r[34]);
  if (isNaN(fat) || fat === 0) {
    fat = (parseFloat(r[9]) || 0) + (parseFloat(r[10]) || 0);
  }
  var orderDate = serialToDate(r[2]);
  if (!codcliIndex[codcli]) codcliIndex[codcli] = [];
  codcliIndex[codcli].push({ date: orderDate, fat: fat, prod: parseInt(String(r[4]).trim(), 10) });
  baseCount++;
}
console.log('Base 8026: ' + baseCount + ' pedidos, ' + Object.keys(codcliIndex).length + ' CODCLIs indexados');

// ============================================================
// 2. Ler RELATÓRIO - CAMPANHAS.xlsx → Dados Geral
// ============================================================
console.log('Lendo RELATÓRIO - CAMPANHAS.xlsx...');
var files = fs.readdirSync(CAM_DIR);
var camFile = files.find(function(f) { return f.includes('CAMPANHAS'); });
var camWb = XLSX.readFile(CAM_DIR + camFile);

// Parse campaign info from per-campaign sheets
var campInfo = {};
camWb.SheetNames.forEach(function(s) {
  if (s === 'Dados Geral' || s === 'Resumo por RCA' || s === 'Planilha1') return;
  var cws = camWb.Sheets[s];
  var crows = XLSX.utils.sheet_to_json(cws, { header: 1, defval: '' });
  campInfo[s] = {
    texto: (crows[4] && crows[4][0]) ? String(crows[4][0]).trim() : '',
    publico: (crows[8] && crows[8][0]) ? String(crows[8][0]).trim() : ''
  };
});

// Parse Dados Geral
var dgWs = camWb.Sheets['Dados Geral'];
var dgRaw = XLSX.utils.sheet_to_json(dgWs, { header: 1, defval: '' });
console.log('Dados Geral: ' + (dgRaw.length - 1) + ' linhas');

// Helper to generate slug (same as dashboard JS)
function limparNome(str) {
  return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-zA-Z0-9]/g, '');
}

function gerarSlug(dataStr, nomeCamp) {
  var ds = dataStr.replace(/\//g, '');
  var partes = nomeCamp.split('|');
  var nome = partes[0].replace('CAMPANHA ', '').trim();
  var setor = partes.length >= 3 ? partes[partes.length - 1].trim() : '';
  var sAbr = { TELEVENDAS: 'TV', REPRESENTANTES: 'RP', ATIVOS: 'AT', EXTERNO: 'EX', PERDIDO: 'PD', 'PRÉ VENDAS': 'PV' };
  var sa = sAbr[setor.toUpperCase()] || limparNome(setor).substring(0, 2).toUpperCase();
  var curto = limparNome(nome).substring(0, 10);
  var diasMatch = nomeCamp.match(/(\d+)\s*[-–]\s*(\d+)/);
  var slug = ds + '_' + curto + '_' + sa;
  if (diasMatch) slug += '_' + diasMatch[1] + '_' + diasMatch[2];
  if (slug.length > 31) slug = slug.substring(0, 31).replace(/_+$/, '');
  return slug;
}

// ============================================================
// 2b. Mapeamento campanha -> produto (SKUs da base 8026)
// ============================================================
// OBS: compra do produto específico da campanha (apenas informativo, não entra no ranking)
var PROD_CAMPANHAS = [
  { test: function(n) { return n.indexOf('OLEO DE ALGODAO') >= 0; }, label: 'Óleo de Algodão 5,1LT', prods: [12417] },
  { test: function(n) { return n.indexOf('FILÉ MEIO PEITO') >= 0; }, label: 'Filé Meio Peito', prods: [13746] },
  { test: function(n) { return n.indexOf('CARNE MOÍDA') >= 0; }, label: 'Carne Moída II', prods: [1376, 12900, 10597] },
  { test: function(n) { return n.indexOf('OURO DA TERRA') >= 0; }, label: 'Batata Ouro da Terra', prods: [13516] },
  { test: function(n) { return n.indexOf('AÇAÍ') >= 0; }, label: 'Açaí', prods: [10937, 11979, 10938, 13283, 13246, 11022, 11980, 11472, 12336, 13196, 10933, 13228, 10936, 13482, 12605, 10926, 12609, 13135, 13263, 12604, 12006, 12004, 11117, 13194, 12603, 10934, 13247, 10935, 12076, 12607, 12077, 12957, 12005, 12606, 12958, 11831, 11387, 13766] },
  { test: function(n) { return n.indexOf('BATATA') >= 0; }, label: 'Batata (fritas congeladas)', prods: [12217, 13229, 12074, 13516, 13231, 11400, 13488, 11844, 13013, 13232, 11934, 12291, 13230, 12429, 11936, 11399, 12969, 13579, 11743, 13170, 12230, 12000, 13221, 13240] }
];

function findProduto(campanhaM) {
  for (var pi = 0; pi < PROD_CAMPANHAS.length; pi++) {
    if (PROD_CAMPANHAS[pi].test(campanhaM)) return PROD_CAMPANHAS[pi];
  }
  return null;
}

// ============================================================
// 3. Para cada CODCLI+DATA, buscar faturamento na base 8026
// ============================================================
console.log('Enriquecendo dados com faturamento da base 8026...');

// Read campaign data using same logic as dashboard parseCampanhas
var rows = [];
var enriched = 0, notFound = 0, withFat = 0;

for (var ri = 1; ri < dgRaw.length; ri++) {
  var r = dgRaw[ri];
  if (!r || !r[0]) continue;

  var codcli = String(r[0]).trim();
  var campaignDateSerial = r[8]; // DATA column (serial)
  var campaignDate = serialToDate(campaignDateSerial);

  // Layout atual do "Dados Geral": col9 vazio, col10=CAMPANHA (disparo), col11=Data,
  // col12=Total de clientes, col13=Receberam, col14=SLUG
  var row = {
    id: r[0],
    resposta: String(r[1]).trim(),
    cod: String(r[2]).trim(),
    vendedor: String(r[3]).trim(),
    setor: String(r[4]).trim() || 'Todos',
    regiao: String(r[5]).trim(),
    supervisor: String(r[6]).trim(),
    campanha: String(r[7]).trim(),
    data: typeof r[8] === 'number' ? XLSX.SSF.format('dd/mm/yyyy', r[8]) : String(r[8]).trim(),
    campanhaM: String(r[10]).trim(),
    dataN: typeof r[11] === 'number' ? XLSX.SSF.format('dd/mm/yyyy', r[11]) : String(r[11]).trim(),
    totalClientes: r[12],
    receberam: r[13],
    slug: String(r[14]).trim()
  };

  // Look up faturamento for this CODCLI after campaign date
  if (codcli && campaignDate && codcliIndex[codcli]) {
    var orders = codcliIndex[codcli];
    var totalFat = 0;
    var hasOrderAfter = false;
    for (var oi = 0; oi < orders.length; oi++) {
      var o = orders[oi];
      if (o.date && o.date >= campaignDate) {
        totalFat += o.fat;
        hasOrderAfter = true;
      }
    }
    row.valor = Math.round(totalFat * 100) / 100;
    row.compra = hasOrderAfter ? 'Sim' : 'Não';
    if (hasOrderAfter) withFat++;
    enriched++;
  } else {
    row.valor = 0;
    row.compra = 'Não';
    notFound++;
  }

  // OBS: compra do produto específico da campanha (usa a campanha que gerou o lead)
  var prod = findProduto(row.campanha || row.campanhaM);
  if (prod) {
    row.produto = prod.label;
    var prodFat = 0, prodHas = false;
    if (codcli && campaignDate && codcliIndex[codcli]) {
      var pOrders = codcliIndex[codcli];
      for (var oi2 = 0; oi2 < pOrders.length; oi2++) {
        var o2 = pOrders[oi2];
        if (o2.date && campaignDate && o2.date >= campaignDate && prod.prods.indexOf(o2.prod) >= 0) {
          prodFat += o2.fat;
          prodHas = true;
        }
      }
    }
    row.comprouProduto = prodHas ? 'Sim' : 'Não';
    row.valorProduto = Math.round(prodFat * 100) / 100;
  } else {
    row.produto = '';
    row.comprouProduto = 'Não';
    row.valorProduto = 0;
  }

  rows.push(row);
}

console.log('Resultados:');
console.log('  Total linhas: ' + rows.length);
console.log('  CODCLI encontrado na base: ' + enriched);
console.log('  Com faturamento > 0: ' + withFat);
console.log('  CODCLI não encontrado: ' + notFound);

// ============================================================
// 4. Build slugMap: campanhaM -> slug (usando a coluna SLUG da planilha)
// ============================================================
var slugMap = {};
for (var si = 0; si < rows.length; si++) {
  var sr = rows[si];
  if (sr.campanhaM && sr.slug) {
    // Chave composta (dataN|campanhaM) resolve campanhas com o mesmo nome em datas diferentes
    slugMap[sr.dataN + '|' + sr.campanhaM] = sr.slug;
    // Fallback por nome (primeira ocorrência vence)
    if (!slugMap[sr.campanhaM]) slugMap[sr.campanhaM] = sr.slug;
  }
}

// Manual overrides for campaigns that don't match automatically
// FILÉ MEIO PEITO: no campInfo sheet, create entry manually
if (!campInfo['27072026_FILEMEIOPEITO_TODOS']) {
  campInfo['27072026_FILEMEIOPEITO_TODOS'] = {
    texto: 'Olá! Aqui é da MAR Food Service! 👋 Estamos com condições especiais em FILÉ MEIO PEITO, válidas somente hoje! Quer saber quais são as ofertas? Responda a esta mensagem que eu te explico os próximos passos. 🚀',
    publico: 'Todos'
  };
}
slugMap['FILÉ MEIO PEITO | SEM COMPRA | TODOS'] = '27072026_FILEMEIOPEITO_TODOS';
slugMap['27/07/2026|FILÉ MEIO PEITO | SEM COMPRA | TODOS'] = '27072026_FILEMEIOPEITO_TODOS';

// CARNE MOÍDA II — manual mapping
slugMap['CAMPANHA CARNE MOÍDA II| SEM COMPRA | TODOS'] = '23072026_CARNEMOIDA2_TODOS';

// OURO DA TERRA — manual mapping (overrides auto if wrong)
slugMap['OURO DA TERRA | SEM COMPRA | TODOS'] = '24072026_OURODATERRA_TODOS';

// Fix auto-mapped entries that were wrong
slugMap['COPA | 20-44 DIAS | REPRESENTANTES'] = '24062026_COPA_RP_20_44';
slugMap['CAMPANHA BATATA| 45-120 DIAS | TELEVENDAS'] = '23062026_BATATA_TV_45_120';
slugMap['CAMPANHA AÇAÍ | 20-44 DIAS | REPRESENTANTES'] = '06072026_ACAI_RP_20_44';
slugMap['CAMPANHA CNPJ| 20 -44 DIAS | ATIVOS'] = '08072026_CNPJ_AT_20_44';
slugMap['CAMPANHA CNPJ| 45-400 DIAS | TELEVENDAS'] = '08072026_CNPJ_TV_45_400';
slugMap['CAMPANHA FECHAMENTO DO MÊS | 45-89 DIAS | TELEVENDAS'] = '29062026_FECHAMENTO_TV_45_89';

// ============================================================
// 5. Salvar data.json atualizado
// ============================================================
var output = {
  rows: rows,
  campInfo: campInfo,
  slugMap: slugMap,
  updatedAt: new Date().toISOString(),
  fileName: camFile
};

var jsonPath = CAM_DIR + 'data.json';
fs.writeFileSync(jsonPath, JSON.stringify(output, null, 2));
console.log('\ndata.json salvo: ' + jsonPath);
console.log('Total rows: ' + rows.length);

// Summary by campaign
var campSummary = {};
rows.forEach(function(r) {
  if (!r.campanha) return;
  if (!campSummary[r.campanha]) campSummary[r.campanha] = { total: 0, recuperados: 0, fat: 0 };
  campSummary[r.campanha].total++;
  if (r.valor > 0) {
    campSummary[r.campanha].recuperados++;
    campSummary[r.campanha].fat += r.valor;
  }
});
console.log('\n=== Resumo por Campanha ===');
Object.keys(campSummary).forEach(function(k) {
  var s = campSummary[k];
  console.log(k + ': ' + s.total + ' clientes, ' + s.recuperados + ' recuperados, R$ ' + s.fat.toFixed(2));
});
