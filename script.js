let cidadesDados = {};

document.getElementById('processBtn').addEventListener('click', async () => {
  const fileInput = document.getElementById('fileInput');
  if (!fileInput.files.length) {
    alert("Selecione um arquivo Excel!");
    return;
  }

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
  renderCidades(cidades);
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
  const dados = cidadesDados[cidade];
  if (!dados || !dados.length) return;

  const wb = new ExcelJS.Workbook();

  // ==================================================
  // PLANILHA DETALHADA (_det)
  // ORDENADA POR NOME DO PACIENTE (A-Z)
  // ==================================================
  const wsDet = wb.addWorksheet(`${cidade}_det`);

  wsDet.addRow([cidade]);
  wsDet.addRow(["NOME", "MEDICAMENTO", "LOTE", "TOTAL"]);

  dados.sort((a, b) =>
    a.nomePaciente.localeCompare(b.nomePaciente, undefined, { sensitivity: 'base' })
  );

  dados.forEach(d =>
    wsDet.addRow([d.nomePaciente, d.med, d.lote, d.qtd])
  );

  formatarTabela(wsDet, 4, dados.length + 2, [3, 4]);

  // ==================================================
  // PLANILHA RESUMO (_res)
  // ORDENADA POR MEDICAMENTO (A-Z)
  // ==================================================
  const wsRes = wb.addWorksheet(`${cidade}_res`);

  wsRes.addRow([`${cidade} - TOTAL MEDICAMENTOS`]);
  wsRes.addRow(["MEDICAMENTO", "LOTE", "TOTAL"]);

  const mapa = {};
  dados.forEach(d => {
    const key = `${d.med}||${d.lote}`;
    mapa[key] = (mapa[key] || 0) + d.qtd;
  });

  Object.keys(mapa)
    .sort((a, b) => {
      const medA = a.split("||")[0];
      const medB = b.split("||")[0];
      return medA.localeCompare(medB, undefined, { sensitivity: 'base' });
    })
    .forEach(k => {
      const [med, lote] = k.split("||");
      wsRes.addRow([med, lote, mapa[k]]);
    });

  formatarTabela(wsRes, 3, Object.keys(mapa).length + 2, [2, 3]);

  // Download
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${cidade}_medicamentos.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ==================================================
// FORMATAÇÃO PADRÃO DAS TABELAS
// ==================================================
function formatarTabela(ws, totalCols, totalRows, centralizarCols) {
  // Título
  ws.mergeCells(1, 1, 1, totalCols);
  ws.getRow(1).font = { bold: true, size: 14 };
  ws.getRow(1).alignment = { horizontal: "center" };

  // Cabeçalho
  for (let c = 1; c <= totalCols; c++) {
    const cell = ws.getRow(2).getCell(c);
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4CAF50" } };
    cell.alignment = { horizontal: "center" };
  }

  // Bordas + alinhamento
  for (let r = 2; r <= totalRows; r++) {
    for (let c = 1; c <= totalCols; c++) {
      const cell = ws.getRow(r).getCell(c);
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
      if (centralizarCols.includes(c)) {
        cell.alignment = { horizontal: "center" };
      }
    }
  }

  ws.columns.forEach(col => col.width = 30);
}
