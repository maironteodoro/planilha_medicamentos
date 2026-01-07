let cidadesDados = {};
let cidadesGeladosLista = []; // lista de medicamentos gelados

document.getElementById('processBtn').addEventListener('click', async () => {
  const fileInput = document.getElementById('fileInput');
  const fileGeladoInput = document.getElementById('fileGeladoInput');

  if (!fileInput.files.length) {
    alert("Selecione um arquivo Excel!");
    return;
  }
  if (!fileGeladoInput.files.length) {
    alert("Selecione o arquivo de medicamentos gelados!");
    return;
  }

  // ================================
  // PROCESSA PLANILHA NORMAL
  // ================================
  const file = fileInput.files[0];
  const arrayBuffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  const sheet = workbook.worksheets[0];
  const headers = [];
  const rows = [];

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      row.eachCell(cell => headers.push(String(cell.value).trim()));
    } else {
      const obj = {};
      row.eachCell((cell, col) => {
        obj[headers[col - 1]] = cell.value;
      });
      rows.push(obj);
    }
  });

  const cidades = {};
  rows.forEach(r => {
    const cidade = String(r["Município de Residência do Paciente"] || "").toUpperCase().trim();
    if (!cidade) return;
    if (!cidades[cidade]) cidades[cidade] = [];
    cidades[cidade].push({
      nomePaciente: String(r["Nome do Paciente"] || "-").trim() || "-",
      med: String(r["Medicamento/Produto (Descrição Genérica)"] || "-").trim() || "-",
      lote: String(r["Lote do Medicamento ou Produto"] || "-").trim() || "-",
      qtd: Number(r["Quantidade do Medicamento ou Produto"] || 0)
    });
  });
  cidadesDados = cidades;

  // ================================
  // PROCESSA PLANILHA GELADOS
  // ================================
  const fileGelado = fileGeladoInput.files[0];
  const arrayBufferGelado = await fileGelado.arrayBuffer();
  const workbookGel = new ExcelJS.Workbook();
  await workbookGel.xlsx.load(arrayBufferGelado);
  const sheetGel = workbookGel.worksheets[0];

  // Lista de medicamentos gelados na coluna A
  cidadesGeladosLista = [];
  sheetGel.eachRow((row, rowNumber) => {
    if (rowNumber >= 1) {
      const med = String(row.getCell(1).value || "").trim();
      if (med) cidadesGeladosLista.push(med.toUpperCase());
    }
  });

  renderCidades(Object.keys(cidades).reduce((acc, c) => { acc[c] = true; return acc; }, {}));
});

function renderCidades(cidades) {
  const div = document.getElementById('cidadesList');
  div.innerHTML = "";

  Object.keys(cidades).sort().forEach(cidade => {
    const row = document.createElement('div');
    row.className = "cidade-item";

    const span = document.createElement('span');
    span.textContent = cidade;

    const btn = document.createElement('button');
    btn.textContent = "Download Excel";
    btn.onclick = () => exportCidade(cidade);

    row.appendChild(span);
    row.appendChild(btn);
    div.appendChild(row);
  });
}

async function exportCidade(cidade) {
  const dados = cidadesDados[cidade] || [];
  if (!dados.length) return;

  const wb = new ExcelJS.Workbook();

  // ================================
  // PLANILHA DETALHADA (_det)
  // ================================
  const wsDet = wb.addWorksheet(`${cidade}_det`);
  wsDet.addRow([cidade]);
  wsDet.addRow(["NOME", "MEDICAMENTO", "LOTE", "TOTAL"]);

  dados.sort((a, b) => a.nomePaciente.localeCompare(b.nomePaciente, undefined, { sensitivity: 'base' }));
  dados.forEach(d => wsDet.addRow([d.nomePaciente, d.med, d.lote, d.qtd]));
  formatarTabela(wsDet, 4, dados.length + 2, [3, 4]);

  // ================================
  // PLANILHA RESUMO (_res)
  // ================================
  const wsRes = wb.addWorksheet(`${cidade}_res`);
  wsRes.addRow([`${cidade} - TOTAL MEDICAMENTOS`]);
  wsRes.addRow(["MEDICAMENTO", "LOTE", "TOTAL"]);

  let totalAgulhas = 0;
  let totalInfliximabe = 0;
  const mapa = {};

  dados.forEach(d => {
    const primeiraPalavra = d.med.split(" ")[0].toUpperCase();
    if (primeiraPalavra === "INSULINA") {
      totalAgulhas += 30; // 30 agulhas por pedido
    } else if (primeiraPalavra === "INFLIXIMABE") {
      totalInfliximabe += 1; // 1 kit por pedido
    } else {
      const key = `${d.med}||${d.lote}`;
      mapa[key] = (mapa[key] || 0) + d.qtd;
    }
  });

  // Adiciona agulhas e kits no topo, se houver
  if (totalAgulhas > 0) wsRes.addRow(["AGULHA", "-", totalAgulhas]);
  if (totalInfliximabe > 0) wsRes.addRow(["KIT INFLIXIMABE", "-", totalInfliximabe]);

  // Adiciona os demais medicamentos ordenados
  Object.keys(mapa).sort((a, b) => {
    const medA = a.split("||")[0];
    const medB = b.split("||")[0];
    return medA.localeCompare(medB, undefined, { sensitivity: 'base' });
  }).forEach(k => {
    const [med, lote] = k.split("||");
    wsRes.addRow([med, lote, mapa[k]]);
  });

  formatarTabela(wsRes, 3, wsRes.rowCount, [2, 3]);

  // ================================
  // PLANILHA GELADOS (_gel)
  // ================================
  const wsGel = wb.addWorksheet(`${cidade}_gel`);
  wsGel.addRow([cidade + " - MEDICAMENTOS GELADOS"]);
  wsGel.addRow(["MEDICAMENTO", "LOTE", "TOTAL"]);

  const mapaGel = {};

  dados.forEach(d => {
    // checa se o medicamento está na lista gelada
    if (!cidadesGeladosLista.includes(d.med.toUpperCase())) return;

    const key = `${d.med}||${d.lote}`;
    mapaGel[key] = (mapaGel[key] || 0) + d.qtd;
  });

  Object.keys(mapaGel).sort((a, b) => {
    const medA = a.split("||")[0];
    const medB = b.split("||")[0];
    return medA.localeCompare(medB, undefined, { sensitivity: 'base' });
  }).forEach(k => {
    const [med, lote] = k.split("||");
    wsGel.addRow([med, lote, mapaGel[k]]);
  });

  // Linha final fixa
  const ultimaLinha = wsGel.addRow(["JUDICIAL GELADEIRA/MEDICAMENTOS MESA"]);
  wsGel.mergeCells(ultimaLinha.number, 1, ultimaLinha.number, 3); // mescla da coluna 1 até 3
  ultimaLinha.alignment = { horizontal: "center" };
  ultimaLinha.font = { bold: true };

  formatarTabela(wsGel, 3, wsGel.rowCount, [2, 3]);

  // ================================
  // DOWNLOAD
  // ================================
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${cidade}_medicamentos.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ================================
// FORMATAÇÃO PADRÃO DAS TABELAS
// ================================
function formatarTabela(ws, totalCols, totalRows, centralizarCols) {
  ws.mergeCells(1, 1, 1, totalCols);
  ws.getRow(1).font = { bold: true, size: 14 };
  ws.getRow(1).alignment = { horizontal: "center" };

  for (let c = 1; c <= totalCols; c++) {
    const cell = ws.getRow(2).getCell(c);
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4CAF50" } };
    cell.alignment = { horizontal: "center" };
  }

  for (let r = 2; r <= totalRows; r++) {
    for (let c = 1; c <= totalCols; c++) {
      const cell = ws.getRow(r).getCell(c);
      cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      if (centralizarCols.includes(c)) cell.alignment = { horizontal: "center" };
    }
  }

  ws.columns.forEach(col => col.width = 30);
}
