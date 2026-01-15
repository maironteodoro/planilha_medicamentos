let cidadesDados = {};
let medicamentosGelados = [];

/********************************
 * LOCALSTORAGE – GELADOS
 ********************************/
function carregarGelados() {
  const salvo = localStorage.getItem("medicamentosGelados");
  medicamentosGelados = salvo ? JSON.parse(salvo) : [];
  renderGelados();
}

function salvarGelados() {
  localStorage.setItem("medicamentosGelados", JSON.stringify(medicamentosGelados));
}

function adicionarGelado() {
  const input = document.getElementById("novoGelado");
  const nome = input.value.trim().toUpperCase();
  if (!nome) return;

  if (!medicamentosGelados.includes(nome)) {
    medicamentosGelados.push(nome);
    salvarGelados();
    renderGelados();
  }
  input.value = "";
}

function removerGelado(nome) {
  medicamentosGelados = medicamentosGelados.filter(m => m !== nome);
  salvarGelados();
  renderGelados();
}

function renderGelados() {
  const div = document.getElementById("listaGelados");
  div.innerHTML = "";

  medicamentosGelados.sort().forEach(med => {
    const linha = document.createElement("div");

    const texto = document.createElement("span");
    texto.textContent = med;

    const btn = document.createElement("button");
    btn.textContent = "X";
    btn.style.marginLeft = "10px";
    btn.onclick = () => removerGelado(med);

    linha.appendChild(texto);
    linha.appendChild(btn);
    div.appendChild(linha);
  });
}

carregarGelados();

/********************************
 * PROCESSAR PLANILHA
 ********************************/
document.getElementById("processBtn").addEventListener("click", async () => {
  const fileInput = document.getElementById("fileInput");
  if (!fileInput.files.length) {
    alert("Selecione um arquivo Excel!");
    return;
  }

  const file = fileInput.files[0];
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheet = workbook.worksheets[0];
  const headers = [];
  const rows = [];

  sheet.eachRow((row, n) => {
    if (n === 1) {
      row.eachCell(c => headers.push(String(c.value).trim()));
    } else {
      const obj = {};
      row.eachCell((c, i) => {
        obj[headers[i - 1]] = c.value;
      });
      rows.push(obj);
    }
  });

  const cidades = {};
  rows.forEach(r => {
    const cidade = String(r["Município de Residência do Paciente"] || "")
      .toUpperCase()
      .trim();

    if (!cidade) return;

    if (!cidades[cidade]) cidades[cidade] = [];

    cidades[cidade].push({
      nomePaciente: String(r["Nome do Paciente"] || "-").trim(),
      med: String(r["Medicamento/Produto (Descrição Genérica)"] || "-").trim(),
      lote: String(r["Lote do Medicamento ou Produto"] || "-").trim(),
      qtd: Number(r["Quantidade do Medicamento ou Produto"] || 0)
    });
  });

  cidadesDados = cidades;
  renderCidades(Object.keys(cidades));
});

/********************************
 * LISTAR CIDADES
 ********************************/
function renderCidades(cidades) {
  const div = document.getElementById("cidadesList");
  div.innerHTML = "";

  cidades.sort().forEach(cidade => {
    const linha = document.createElement("div");

    const texto = document.createElement("span");
    texto.textContent = cidade;

    const btn = document.createElement("button");
    btn.textContent = "Download Excel";
    btn.style.marginLeft = "10px";
    btn.onclick = () => exportCidade(cidade);

    linha.appendChild(texto);
    linha.appendChild(btn);
    div.appendChild(linha);
  });
}

/********************************
 * EXPORTAR EXCEL
 ********************************/
async function exportCidade(cidade) {
  const dados = cidadesDados[cidade];
  if (!dados || !dados.length) return;

  const wb = new ExcelJS.Workbook();

  /* ===== DET ===== */
  const wsDet = wb.addWorksheet(`${cidade}_det`);
  wsDet.addRow([cidade]);
  wsDet.addRow(["NOME", "MEDICAMENTO", "LOTE", "TOTAL"]);

  dados.sort((a, b) => a.nomePaciente.localeCompare(b.nomePaciente, "pt-BR"));
  dados.forEach(d => wsDet.addRow([d.nomePaciente, d.med, d.lote, d.qtd]));
  formatarTabela(wsDet, 4, wsDet.rowCount, [3, 4]);

  /* ===== RES ===== */
  const wsRes = wb.addWorksheet(`${cidade}_res`);
  wsRes.addRow([`${cidade} - TOTAL MEDICAMENTOS`]);
  wsRes.addRow(["MEDICAMENTO", "LOTE", "TOTAL"]);

  let totalAgulhas = 0;
  let totalInflix = 0;
  const mapa = {};

  dados.forEach(d => {
    const prim = d.med.split(" ")[0].toUpperCase();
    if (prim === "INSULINA") totalAgulhas += 30;
    else if (prim === "INFLIXIMABE") totalInflix += 1;
    else {
      const key = `${d.med}||${d.lote}`;
      mapa[key] = (mapa[key] || 0) + d.qtd;
    }
  });

  if (totalAgulhas) wsRes.addRow(["AGULHA", "-", totalAgulhas]);
  if (totalInflix) wsRes.addRow(["KIT INFLIXIMABE", "-", totalInflix]);

  Object.keys(mapa).sort().forEach(k => {
    const [med, lote] = k.split("||");
    wsRes.addRow([med, lote, mapa[k]]);
  });

  formatarTabela(wsRes, 3, wsRes.rowCount, [2, 3]);

  /* ===== GEL ===== */
  const wsGel = wb.addWorksheet(`${cidade}_gel`);
  wsGel.addRow([`${cidade} - MEDICAMENTOS GELADOS`]);
  wsGel.addRow(["MEDICAMENTO", "LOTE", "TOTAL"]);

  const mapaGel = {};
  dados.forEach(d => {
    if (!medicamentosGelados.includes(d.med.toUpperCase())) return;
    const key = `${d.med}||${d.lote}`;
    mapaGel[key] = (mapaGel[key] || 0) + d.qtd;
  });

  Object.keys(mapaGel).sort().forEach(k => {
    const [med, lote] = k.split("||");
    wsGel.addRow([med, lote, mapaGel[k]]);
  });

  const fim = wsGel.addRow(["JUDICIAL GELADEIRA/MEDICAMENTOS MESA"]);
  wsGel.mergeCells(fim.number, 1, fim.number, 3);
  fim.alignment = { horizontal: "center" };
  fim.font = { bold: true };

  formatarTabela(wsGel, 3, wsGel.rowCount, [2, 3]);

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${cidade}_medicamentos.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/********************************
 * FORMATAÇÃO
 ********************************/
function formatarTabela(ws, totalCols, totalRows, centralizar) {
  ws.mergeCells(1, 1, 1, totalCols);
  ws.getRow(1).font = { bold: true, size: 14 };
  ws.getRow(1).alignment = { horizontal: "center" };

  for (let r = 2; r <= totalRows; r++) {
    for (let c = 1; c <= totalCols; c++) {
      const cell = ws.getRow(r).getCell(c);
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
      if (centralizar.includes(c)) {
        cell.alignment = { horizontal: "center" };
      }
    }
  }

  ws.columns.forEach(col => col.width = 30);
}
