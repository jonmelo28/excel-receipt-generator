const path = require("path");
const fs = require("fs");
const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const puppeteer = require("puppeteer");
const { PDFDocument } = require("pdf-lib");
const { formatCpf, valorEmReaisPorExtenso } = require("./extenso");

const app = express();
app.use(express.json());

app.use(express.urlencoded({ extended: true })); // ler campos do form
app.use("/public", express.static(path.join(__dirname, "public"))); // servir imagens

const upload = multer({ dest: path.join(__dirname, "uploads") });

const TEMPLATE_PATH = path.join(__dirname, "templates", "recibo.html");
const TEMPLATE_FRAGMENT_PATH = path.join(__dirname, "templates", "recibo_fragment.html");
const TEMPLATE_2P_PATH = path.join(__dirname, "templates", "impressao_2porpagina.html");
const OUTPUT_DIR = path.join(__dirname, "output");
fs.mkdirSync(OUTPUT_DIR, { recursive: true });

function brl(v) {
  const n = Number(v ?? 0);
  return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }).replace("R$", "").trim();
}

function hojeBR() {
  const d = new Date();
  return d.toLocaleDateString("pt-BR");
}

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function imageToDataUri(absPath) {
  const ext = path.extname(absPath).toLowerCase();
  const mime =
    ext === ".png" ? "image/png" :
    (ext === ".jpg" || ext === ".jpeg") ? "image/jpeg" :
    ext === ".webp" ? "image/webp" :
    "application/octet-stream";

  const buf = fs.readFileSync(absPath);
  return `data:${mime};base64,${buf.toString("base64")}`;
}

function parseMoneyBR(value) {
  if (value == null) return NaN;

  // Se já for número (excel geralmente traz assim)
  if (typeof value === "number") return value;

  let s = String(value).trim();

  // remove "R$", espaços e outros
  s = s.replace(/\s/g, "").replace(/^R\$/i, "");

  const hasComma = s.includes(",");
  const hasDot = s.includes(".");

  // Caso "1.234,56" => '.' milhares e ',' decimal
  if (hasComma && hasDot) {
    // Assume que o último separador indica o decimal
    const lastComma = s.lastIndexOf(",");
    const lastDot = s.lastIndexOf(".");
    if (lastComma > lastDot) {
      // decimal = ',' -> remove dots e troca comma por dot
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      // decimal = '.' -> remove commas
      s = s.replace(/,/g, "");
    }
  } else if (hasComma && !hasDot) {
    // "67,32" -> decimal = ','
    s = s.replace(",", ".");
  } else {
    // "67.32" ou "6732" -> deixa como está
    // mas remove separadores de milhar caso exista algo tipo "1.234"
    // (se tiver 1 ponto e 3 dígitos depois, provavelmente é milhar)
    const parts = s.split(".");
    if (parts.length === 2 && parts[1].length === 3) {
      s = parts.join("");
    }
  }

  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

function renderTemplate(html, data) {
  // substituição simples {{CAMPO}}
  return html.replace(/\{\{(\w+)\}\}/g, (_, key) => escapeHtml(data[key] ?? ""));
}

function normalizeHeader(s) {
  return String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9]+/g, "");
}

function readRowsFromExcel(filePath) {
  const wb = xlsx.readFile(filePath, { cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  // lê como matriz (rows/cols)
  const aoa = xlsx.utils.sheet_to_json(ws, { header: 1, defval: "" }); // array of arrays

  // acha a linha do cabeçalho
  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(30, aoa.length); i++) {
    const row = aoa[i].map(c => normalizeHeader(c));
    const hasCPF = row.includes("CPF") || row.includes("CPFCNPJ") || row.includes("DOCUMENTO");
    const hasNOME = row.includes("NOME") || row.includes("NOMECOMPLETO");
    const hasVALOR = row.includes("VALOR") || row.includes("TOTAL") || row.includes("PAGAMENTO");
    if (hasCPF && hasNOME && hasVALOR) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) {
    // debug útil
    const preview = aoa.slice(0, 8).map(r => r.join(" | "));
    throw new Error(
      "Não encontrei o cabeçalho (CPF, NOME, VALOR). " +
      "Verifique se há linhas/títulos acima do cabeçalho.\nPrévia:\n" +
      preview.join("\n")
    );
  }

  const header = aoa[headerRowIndex].map(h => String(h ?? "").trim());
  const dataRows = aoa.slice(headerRowIndex + 1);

  // transforma em objetos usando o header encontrado
  const rows = dataRows
    .filter(r => r.some(cell => String(cell ?? "").trim() !== "")) // remove linhas totalmente vazias
    .map((r, idx) => {
      const obj = {};
      header.forEach((h, col) => (obj[h] = r[col] ?? ""));
      return { _index: idx + 1, ...obj };
    });

  // agora mapeia os campos tolerantes
  function getField(rowObj, candidates) {
    const entries = Object.entries(rowObj);
    for (const [k, v] of entries) {
      const nk = normalizeHeader(k);
      if (candidates.includes(nk)) return v;
    }
    return "";
  }

  return rows.map((r, i) => ({
    _index: i + 1,
    CPF: getField(r, ["CPF", "CPFCNPJ", "DOCUMENTO", "DOC"]),
    NOME: getField(r, ["NOME", "NOMECOMPLETO", "CLIENTE", "BENEFICIARIO"]),
    VALOR: getField(r, ["VALOR", "TOTAL", "PAGAMENTO", "VALORR", "VALORRS"]),
    POR_EXTENSO: getField(r, ["POREXTENSO", "EXTENSO", "VALOREXTENSO", "VALORPOREXTENSO"]),
  }));
}

async function htmlToPdfBuffer(browser, html) {
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: "networkidle0" });
  const pdf = await page.pdf({
    format: "A4",
    printBackground: true,
    margin: { top: "12mm", right: "12mm", bottom: "12mm", left: "12mm" },
  });
  await page.close();
  return pdf;
}

async function mergePdfs(buffers) {
  const out = await PDFDocument.create();
  for (const b of buffers) {
    const doc = await PDFDocument.load(b);
    const pages = await out.copyPages(doc, doc.getPageIndices());
    pages.forEach((p) => out.addPage(p));
  }
  return await out.save();
}

async function gerarRecibos(req) {
  const planilhaPath = req.file?.path;
  if (!planilhaPath) throw new Error("Envie o arquivo no campo planilha");

  const { diaInicial, diaFinal, mes, ano, cnpjTipo } = req.body;
  const periodo = `${diaInicial} à ${diaFinal} de ${String(mes).toUpperCase()} de ${ano}`;

  const CNPJ_TO_LOGO = {
    "11.111.111/0001-11": path.join(__dirname, "public", "logos", "1.png"),
    "22.222.222/0001-22": path.join(__dirname, "public", "logos", "2.png"),
  };

  const CNPJ_DESC = {
    "11.111.111/0001-11": "EMPRESA 1",
    "22.222.222/0001-22": "EMPRESA 2",
  };

  function cnpjByDesc(cnpj) {
    const key = String(cnpj || "").trim();
    return CNPJ_DESC[key] || "DGM DISTRIBUIDORA";
  }

  function logoByCnpj(cnpj) {
    const key = String(cnpj || "").trim();
    return CNPJ_TO_LOGO[key] || path.join(__dirname, "public", "logos", "DGM.jpg");
  }

  const logoFile = logoByCnpj(cnpjTipo);
  const logoUrl = imageToDataUri(logoFile);

  const CIDADE_UF = "Itabaiana/SE";
  const DATA_EXTENSO = dataHojePorExtenso();
  const EMPRESA = cnpjByDesc(cnpjTipo);

  const template = fs.readFileSync(TEMPLATE_PATH, "utf8");
  const frag = fs.readFileSync(TEMPLATE_FRAGMENT_PATH, "utf8");
  const master = fs.readFileSync(TEMPLATE_2P_PATH, "utf8");

  try {
    const rows = readRowsFromExcel(planilhaPath);

    const valid = rows.filter(r => String(r.NOME).trim() && String(r.CPF).trim() && String(r.VALOR).trim());
    if (valid.length === 0) throw new Error("Nenhuma linha válida. Precisa ter CPF, NOME e VALOR.");

    const runId = Date.now().toString();
    const batchDir = path.join(OUTPUT_DIR, runId);
    fs.mkdirSync(batchDir, { recursive: true });

    const browser = await puppeteer.launch({ headless: "new" });

    const pdfBuffers = [];
    const files = [];

    // 2 por página
    const blocks = [];

    for (const r of valid) {
      const valorNum = parseMoneyBR(r.VALOR);
      if (!Number.isFinite(valorNum)) throw new Error(`Valor inválido na linha ${r._index}: ${r.VALOR}`);

      const cpfFmt = formatCpf(r.CPF);
      const valorFmt = brl(valorNum);
      const extenso = valorEmReaisPorExtenso(valorNum);

      const data = {
        NOME: r.NOME,
        CPF: cpfFmt,
        VALOR_FORMATADO: valorFmt,
        VALOR_EXTENSO: extenso,
        PERIODO: periodo,
        CIDADE_UF,
        CNPJ: cnpjTipo,
        DATA_EXTENSO,
        LOGO_URL: logoUrl,
        EMPRESA
      };

      // fragmento (2 por página)
      blocks.push(renderTemplate(frag, data));

      // PDF individual (opcional, você mantém)
      const html = renderTemplate(template, data);
      const pdf = await htmlToPdfBuffer(browser, html);
      pdfBuffers.push(pdf);

      const safeName = String(r.NOME).replace(/[^\w\s-]/g, "").trim().slice(0, 60).replace(/\s+/g, "_");
      const fileName = `${String(r._index).padStart(3, "0")}_${safeName || "recibo"}.pdf`;
      const filePath = path.join(batchDir, fileName);

      fs.writeFileSync(filePath, pdf);
      files.push({ fileName, filePath });
    }

    // PDF 2 por página
    const pages = [];
    for (let i = 0; i < blocks.length; i += 2) {
      const first = blocks[i] || "";
      const second = blocks[i + 1] || `<div class="frame receipt" style="border:0"></div>`;
      pages.push(`<div class="page">${first}${second}</div>`);
    }

    const finalHtml2 = master.replace("{{PAGES}}", pages.join("\n"));
    const pdf2porpagina = await htmlToPdfBuffer(browser, finalHtml2);

    const merged2Name = `recibos_2porpagina_${runId}.pdf`;
    fs.writeFileSync(path.join(batchDir, merged2Name), pdf2porpagina);

    await browser.close();

    // PDF único 1 por página (merge)
    const merged = await mergePdfs(pdfBuffers);
    const mergedName = `recibos_${runId}.pdf`;
    fs.writeFileSync(path.join(batchDir, mergedName), merged);

    return {
      runId,
      mergedName,
      merged2Name,
      total: files.length,
      arquivos: files.map(f => f.fileName),
    };
  } finally {
    // sempre remove o upload temporário
    try { fs.unlinkSync(planilhaPath); } catch {}
  }
}

function dataHojePorExtenso() {
  const d = new Date();
  // pt-BR: "6 de junho de 2024"
  const dia = String(d.getDate()).padStart(2, "0");
  const mes = d.toLocaleDateString("pt-BR", { month: "long" });
  const ano = d.getFullYear();
  return `${dia} de ${mes} de ${ano}`;
}

// servir PDFs gerados
app.use("/output", express.static(OUTPUT_DIR));

app.get("/", (req, res) => {
  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.end(`
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Gerador de Recibos</title>
  <style>
    :root{
      --bg:#f6f7fb;
      --card:#ffffff;
      --text:#0f172a;
      --muted:#64748b;
      --border:#e2e8f0;
      --primary:#6b1fa6;
      --primary2:#7c3aed;
      --shadow: 0 14px 40px rgba(2,6,23,.10);
      --radius: 16px;
    }
    *{ box-sizing:border-box; }
    body{
      margin:0;
      font-family: Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
    }
    .wrap{
      min-height:100vh;
      display:flex;
      align-items:center;
      justify-content:center;
      padding: 28px 16px;
    }
    .card{
      width: 100%;
      max-width: 860px;
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      overflow: hidden;
    }
    .header{
      padding: 20px 22px;
      border-bottom: 1px solid var(--border);
      background: linear-gradient(135deg, rgba(107,31,166,.08), rgba(124,58,237,.05));
    }
    .title{
      margin:0;
      font-size: 20px;
      font-weight: 800;
      letter-spacing: .2px;
    }
    .subtitle{
      margin:6px 0 0;
      color: var(--muted);
      font-size: 13px;
    }

    form{
      padding: 22px;
    }
    .grid{
      display:grid;
      grid-template-columns: 1.2fr 1fr 1fr;
      gap: 14px;
    }
    @media (max-width: 820px){
      .grid{ grid-template-columns: 1fr; }
    }

    .field{
      display:flex;
      flex-direction:column;
      gap:6px;
    }
    label{
      font-size: 12px;
      color: var(--muted);
      font-weight: 700;
    }
    input, select{
      width: 100%;
      padding: 12px 12px;
      border: 1px solid var(--border);
      border-radius: 12px;
      font-size: 14px;
      outline: none;
      background: #fff;
    }
    input:focus, select:focus{
      border-color: rgba(107,31,166,.55);
      box-shadow: 0 0 0 4px rgba(107,31,166,.12);
    }
    .hint{
      margin-top: 6px;
      font-size: 12px;
      color: var(--muted);
    }
    .section{
      margin-top: 16px;
      padding-top: 16px;
      border-top: 1px dashed var(--border);
    }
    .actions{
      display:flex;
      gap: 12px;
      align-items:center;
      margin-top: 18px;
    }
    .btn{
      cursor:pointer;
      border: 0;
      padding: 12px 16px;
      border-radius: 12px;
      font-weight: 800;
      font-size: 14px;
    }
    .btn-primary{
      color:#fff;
      background: linear-gradient(135deg, var(--primary), var(--primary2));
      box-shadow: 0 10px 24px rgba(107,31,166,.25);
    }
    .btn-primary:active{ transform: translateY(1px); }
    .btn-secondary{
      background:#eef2ff;
      color:#1f2937;
    }
    .footer{
      padding: 14px 22px 18px;
      border-top: 1px solid var(--border);
      color: var(--muted);
      font-size: 12px;
    }

    /* LOADING OVERLAY */
    #loadingOverlay{
      display:none;
      position:fixed;
      inset:0;
      background:rgba(255,255,255,0.86);
      z-index:9999;
      align-items:center;
      justify-content:center;
      flex-direction:column;
    }
    .spinner{
      width:60px;
      height:60px;
      border:6px solid #ddd;
      border-top-color: var(--primary);
      border-radius:50%;
      animation:spin 1s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
  </style>
</head>

<body>
  <div class="wrap">
    <div class="card">
      <div class="header">
        <h1 class="title">Gerador de Recibos</h1>
        <p class="subtitle">Importe sua planilha (.xlsx) e gere um PDF pronto para impressão (2 recibos por página).</p>
      </div>

      <form id="formRecibo" action="/recibos/gerar-ui" method="post" enctype="multipart/form-data" target="hiddenFrame">
        <div class="grid">
          <div class="field">
            <label>Planilha (.xlsx)</label>
            <input type="file" name="planilha" accept=".xlsx" required />
            <div class="hint">Colunas esperadas: CPF, NOME, VALOR</div>
          </div>

          <div class="field">
            <label>Dia inicial</label>
            <input type="number" name="diaInicial" min="1" max="31" required />
          </div>

          <div class="field">
            <label>Dia final</label>
            <input type="number" name="diaFinal" min="1" max="31" required />
          </div>
        </div>

        <div class="grid section">
          <div class="field">
            <label>Mês</label>
            <select name="mes" required>
              <option value="">Selecione</option>
              <option value="JANEIRO">Janeiro</option>
              <option value="FEVEREIRO">Fevereiro</option>
              <option value="MARÇO">Março</option>
              <option value="ABRIL">Abril</option>
              <option value="MAIO">Maio</option>
              <option value="JUNHO">Junho</option>
              <option value="JULHO">Julho</option>
              <option value="AGOSTO">Agosto</option>
              <option value="SETEMBRO">Setembro</option>
              <option value="OUTUBRO">Outubro</option>
              <option value="NOVEMBRO">Novembro</option>
              <option value="DEZEMBRO">Dezembro</option>
            </select>
          </div>

          <div class="field">
            <label>Ano</label>
            <input type="number" name="ano" min="2000" max="2100" required />
          </div>

          <div class="field">
            <label>CNPJ (empresa pagadora)</label>
            <select name="cnpjTipo" required>
              <option value="">Selecione</option>
              <option value="09.350.550/0001-05">09.350.550/0001-05</option>
              <option value="54.731.011/0001-70">54.731.011/0001-70</option>
            </select>
            <div class="hint">A logo será escolhida automaticamente conforme o CNPJ.</div>
          </div>
        </div>

        <div class="actions">
          <button class="btn btn-primary" type="submit">Gerar recibos</button>
        </div>
      </form>

      <div class="footer">
        Dica: use o PDF “2 por página” para economizar papel na impressão.
      </div>
    </div>
  </div>

  <!-- LOADING -->
  <div id="loadingOverlay">
    <div class="spinner"></div>
    <div style="margin-top:14px;font-size:16px;font-weight:800;">Gerando recibos...</div>
    <div style="margin-top:6px;font-size:12px;color:#64748b;">Aguarde, isso pode levar alguns segundos.</div>
  </div>

  <iframe name="hiddenFrame" style="display:none;"></iframe>

  <script>
    const form = document.getElementById("formRecibo");
    const overlay = document.getElementById("loadingOverlay");
    const iframe = document.querySelector('iframe[name="hiddenFrame"]');

    form.addEventListener("submit", () => {
      overlay.style.display = "flex";
    });

    iframe.onload = function () {
      overlay.style.display = "none";
    };
  </script>
</body>
</html>
  `);
});

app.get("/gerando", (req, res) => {
  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.end(`
  <html>
  <head>
    <meta charset="utf-8"/>
    <title>Gerando recibos...</title>
    <style>
      body { font-family: Arial, sans-serif; background:#f6f7fb; display:flex; align-items:center; justify-content:center; height:100vh; }
      .card { background:#fff; padding:24px 28px; border-radius:14px; box-shadow:0 10px 30px rgba(0,0,0,.08); width:420px; text-align:center; }
      .spin { width:44px; height:44px; border:5px solid #ddd; border-top-color:#6b1fa6; border-radius:50%; margin:0 auto 14px; animation: r 1s linear infinite; }
      @keyframes r { to { transform: rotate(360deg); } }
      .muted { color:#666; font-size:13px; margin-top:6px; }
    </style>
  </head>
  <body>
    <div class="card">
      <div class="spin"></div>
      <h2>Gerando recibos…</h2>
      <div class="muted">Não feche esta página. Assim que terminar, o sistema abrirá o resultado.</div>
    </div>
  </body>
  </html>
  `);
});

app.post("/recibos/gerar", upload.single("planilha"), async (req, res) => {
  try {
    const result = await gerarRecibos(req);
    res.json({
      ok: true,
      total: result.total,
      pasta: result.runId,
      pdf_unico: `/output/${result.runId}/${result.mergedName}`,
      pdf_2porpagina: `/output/${result.runId}/${result.merged2Name}`,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.post("/recibos/gerar-ui", upload.single("planilha"), async (req, res) => {
  try {
    const result = await gerarRecibos(req);
    res.send(`
      <script>
        window.top.location.href="/recibos/resultado/${result.runId}";
      </script>
    `);
  } catch (e) {
    console.error(e);
    res.send(`<pre>${e.message}</pre>`);
  }
});

app.get("/recibos/resultado/:runId", (req, res) => {
  const runId = req.params.runId;

  const pdf2 = `recibos_2porpagina_${runId}.pdf`;
  const pdf1 = `recibos_${runId}.pdf`;

  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.end(`
  <html>
  <head>
    <meta charset="utf-8"/>
    <title>Recibos gerados</title>
    <style>
      body { font-family: Arial, sans-serif; background:#f6f7fb; padding:24px; }
      .wrap { max-width:900px; margin:0 auto; }
      .card { background:#fff; padding:20px; border-radius:14px; box-shadow:0 10px 30px rgba(0,0,0,.08); }
      .btn { display:inline-block; padding:12px 16px; border-radius:10px; text-decoration:none; margin-right:10px; font-weight:700; }
      .primary { background:#6b1fa6; color:#fff; }
      .secondary { background:#eaeaf2; color:#111; }
      iframe { width:100%; height:80vh; border:1px solid #ddd; border-radius:12px; margin-top:16px; background:#fff; }
      .muted { color:#666; font-size:13px; margin-top:6px; }
    </style>
  </head>
  <body>
    <div class="wrap">
      <div class="card">
        <h2>Recibos gerados ✅</h2>
        <div class="muted">Preferência: imprimir 2 por página para economizar papel.</div>

        <div style="margin-top:14px;">
          <a class="btn primary" target="_blank" href="/output/${runId}/${pdf2}">Imprimir (2 por página)</a>
          <a class="btn secondary" target="_blank" href="/output/${runId}/${pdf1}">PDF normal (1 por página)</a>
          <a class="btn secondary" href="/">Gerar novos recibos</a>
        </div>

        <iframe src="/output/${runId}/${pdf2}"></iframe>
      </div>
    </div>
  </body>
  </html>
  `);
});

app.listen(3007, () => console.log("Rodando em http://localhost:3007"));