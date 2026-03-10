import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.worker.min.mjs";

const state = {
  ordersMap: new Map(),
  altoMap: new Map(),
  shipMap: new Map(),
  mergedRows: []
};

const ORDER_ID_REGEX = /\b\d{2}-\d{5}-\d{5}\b/g;
const MONEY_REGEX = /\$\d{1,3}(?:,\d{3})*\.\d{2}/g;

document.getElementById("processBtn").addEventListener("click", processFiles);
document.getElementById("exportBtn").addEventListener("click", exportCsv);

function money(value) {
  return `$${(Number(value) || 0).toFixed(2)}`;
}

function num(value) {
  if (value === null || value === undefined) return 0;
  const cleaned = String(value).replace(/[$,]/g, "").trim();
  const n = parseFloat(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function normalizeOrderId(value) {
  return String(value || "")
    .trim()
    .replace(/^"+|"+$/g, "")
    .replace(/\s+/g, "");
}

function findHeaderRow(rows, requiredHeaders) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].map(x => String(x || "").trim());
    const hasAll = requiredHeaders.every(h => row.includes(h));
    if (hasAll) return i;
  }
  return -1;
}

function getFirstExisting(obj, names) {
  for (const name of names) {
    if (obj[name] !== undefined && String(obj[name]).trim() !== "") {
      return obj[name];
    }
  }
  return "";
}

function parseCsvFile(file, requiredHeaders) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      complete(results) {
        try {
          const rows = results.data || [];
          const headerIndex = findHeaderRow(rows, requiredHeaders);

          if (headerIndex === -1) {
            reject(new Error(`Could not find required header row: ${requiredHeaders.join(", ")}`));
            return;
          }

          const headers = rows[headerIndex].map(h => String(h || "").trim());
          const records = rows
            .slice(headerIndex + 1)
            .filter(row => Array.isArray(row) && row.some(cell => String(cell || "").trim() !== ""))
            .map(row => {
              const obj = {};
              headers.forEach((h, idx) => {
                obj[h] = row[idx] !== undefined ? row[idx] : "";
              });
              return obj;
            });

          resolve(records);
        } catch (err) {
          reject(err);
        }
      },
      error(err) {
        reject(err);
      },
      skipEmptyLines: false
    });
  });
}

async function processFiles() {
  const ordersFile = document.getElementById("ordersFile").files[0];
  const shipFile = document.getElementById("shipFile").files[0];
  const altoFile = document.getElementById("altoFile").files[0];

  if (!ordersFile) {
    alert("Upload the eBay Orders CSV first.");
    return;
  }

  try {
    const orderRecords = await parseCsvFile(ordersFile, ["Order Number", "Custom Label"]);
    state.ordersMap = buildOrdersMap(orderRecords);

    state.shipMap = new Map();
    if (shipFile) {
      const shipRecords = await parseCsvFile(shipFile, ["Order #", "Shipping Cost"]);
      state.shipMap = buildShipMap(shipRecords);
    }

    state.altoMap = new Map();
    if (altoFile) {
      state.altoMap = await parseAltoStatementPdf(altoFile);
    }

    state.mergedRows = buildMergedRows();
    renderAll();
    document.getElementById("exportBtn").disabled = false;
  } catch (err) {
    console.error(err);
    alert(`Processing failed: ${err.message}`);
  }
}

function buildOrdersMap(records) {
  const map = new Map();

  for (const r of records) {
    const orderId = normalizeOrderId(getFirstExisting(r, ["Order Number"]));
    if (!orderId) continue;

    const sold = num(getFirstExisting(r, ["Sold For"]));
    const buyerShip = num(getFirstExisting(r, ["Shipping And Handling"]));
    const tax = num(getFirstExisting(r, ["eBay Collected Tax", "eBay Collect And Remit Tax", "Sales Tax"]));
    const sku = String(getFirstExisting(r, ["Custom Label"])).trim();
    const soldDate = String(getFirstExisting(r, ["Sale Date", "Sales Record Number"])).trim();

    if (!map.has(orderId)) {
      map.set(orderId, {
        orderId,
        soldDate,
        sold: 0,
        buyerShip: 0,
        tax: 0,
        skus: new Set()
      });
    }

    const row = map.get(orderId);
    row.sold += sold;
    row.buyerShip += buyerShip;
    row.tax += tax;
    if (sku) row.skus.add(sku);
  }

  return map;
}

function buildShipMap(records) {
  const map = new Map();

  for (const r of records) {
    const orderId = normalizeOrderId(getFirstExisting(r, ["Order #"]));
    if (!orderId) continue;

    const cost = num(getFirstExisting(r, ["Shipping Cost"]));
    map.set(orderId, (map.get(orderId) || 0) + cost);
  }

  return map;
}

async function parseAltoStatementPdf(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const altoMap = new Map();

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map(item => item.str).join("\n");

    const orderIds = [...pageText.matchAll(ORDER_ID_REGEX)].map(m => m[0]);
    const amounts = [...pageText.matchAll(MONEY_REGEX)].map(m => num(m[0]));

    if (!orderIds.length) continue;

    const tailAmounts = amounts.slice(-orderIds.length);

    if (tailAmounts.length !== orderIds.length) {
      console.warn(`PDF page ${pageNum}: order/amount count mismatch`, orderIds.length, tailAmounts.length);
    }

    const count = Math.min(orderIds.length, tailAmounts.length);

    for (let i = 0; i < count; i++) {
      const orderId = normalizeOrderId(orderIds[i]);
      const amount = tailAmounts[i];

      if (!altoMap.has(orderId)) {
        altoMap.set(orderId, 0);
      }
      altoMap.set(orderId, altoMap.get(orderId) + amount);
    }
  }

  return altoMap;
}

function estimateEbayFee(sold, buyerShip, tax) {
  const feeBase = sold + buyerShip + tax;
  const rate = 0.136;
  const orderFee = (sold + buyerShip) <= 10 ? 0.30 : 0.40;
  return (feeBase * rate) + orderFee;
}

function buildMergedRows() {
  const allOrderIds = new Set([
    ...state.ordersMap.keys(),
    ...state.shipMap.keys(),
    ...state.altoMap.keys()
  ]);

  const rows = [];

  for (const orderId of allOrderIds) {
    const saleRow = state.ordersMap.get(orderId);
    const sold = saleRow ? saleRow.sold : 0;
    const buyerShip = saleRow ? saleRow.buyerShip : 0;
    const tax = saleRow ? saleRow.tax : 0;
    const soldDate = saleRow ? saleRow.soldDate : "";
    const skus = saleRow ? [...saleRow.skus].join(" | ") : "";

    const shipCost = state.shipMap.get(orderId) || 0;
    const partsCost = state.altoMap.get(orderId) || 0;
    const estFee = saleRow ? estimateEbayFee(sold, buyerShip, tax) : 0;

    rows.push({
      orderId,
      soldDate,
      skus,
      sold,
      buyerShip,
      tax,
      shipCost,
      partsCost,
      estFee
    });
  }

  rows.sort((a, b) => b.orderId.localeCompare(a.orderId));
  return rows;
}

function getStatus(row) {
  const hasSale = state.ordersMap.has(row.orderId);
  const hasShip = state.shipMap.has(row.orderId);
  const hasAlto = state.altoMap.has(row.orderId);

  if (hasSale && hasShip && hasAlto) {
    return { text: "Matched all 3", className: "status-all" };
  }
  if (hasSale && (hasShip || hasAlto)) {
    return { text: hasShip ? "Missing Alto cost" : "Missing shipping", className: "status-partial" };
  }
  if (hasSale) {
    return { text: "Sales only", className: "status-missing" };
  }
  return { text: "Not in eBay orders", className: "status-missing" };
}

function calcProfit(row) {
  return row.sold + row.buyerShip - row.estFee - row.shipCost - row.partsCost;
}

function calcMargin(row) {
  return row.sold > 0 ? (calcProfit(row) / row.sold) * 100 : 0;
}

function renderAll() {
  renderSummary();
  renderMatchSummary();
  renderTable();
}

function renderSummary() {
  const summary = document.getElementById("summary");
  const rows = state.mergedRows.filter(r => state.ordersMap.has(r.orderId));

  const totalOrders = rows.length;
  const totalSold = rows.reduce((s, r) => s + r.sold, 0);
  const totalBuyerShip = rows.reduce((s, r) => s + r.buyerShip, 0);
  const totalTax = rows.reduce((s, r) => s + r.tax, 0);
  const totalShipCost = rows.reduce((s, r) => s + r.shipCost, 0);
  const totalPartsCost = rows.reduce((s, r) => s + r.partsCost, 0);
  const totalFees = rows.reduce((s, r) => s + r.estFee, 0);
  const totalProfit = rows.reduce((s, r) => s + calcProfit(r), 0);

  summary.innerHTML = `
    ${summaryCard("Orders", totalOrders)}
    ${summaryCard("Sold", money(totalSold))}
    ${summaryCard("Buyer Shipping", money(totalBuyerShip))}
    ${summaryCard("Tax", money(totalTax))}
    ${summaryCard("Ship Cost", money(totalShipCost))}
    ${summaryCard("Parts Cost", money(totalPartsCost))}
    ${summaryCard("Est. eBay Fees", money(totalFees))}
    ${summaryCard("Profit", money(totalProfit))}
  `;
}

function renderMatchSummary() {
  const matchSummary = document.getElementById("matchSummary");
  const rows = state.mergedRows.filter(r => state.ordersMap.has(r.orderId));

  let all3 = 0;
  let missingShip = 0;
  let missingAlto = 0;
  let salesOnly = 0;

  for (const row of rows) {
    const hasShip = state.shipMap.has(row.orderId);
    const hasAlto = state.altoMap.has(row.orderId);

    if (hasShip && hasAlto) all3++;
    else if (hasShip && !hasAlto) missingAlto++;
    else if (!hasShip && hasAlto) missingShip++;
    else salesOnly++;
  }

  matchSummary.innerHTML = `
    ${summaryCard("Matched all 3", all3)}
    ${summaryCard("Missing Alto cost", missingAlto)}
    ${summaryCard("Missing shipping", missingShip)}
    ${summaryCard("Sales only", salesOnly)}
  `;
}

function summaryCard(label, value) {
  return `
    <div class="summary-card">
      <div class="label">${label}</div>
      <div class="value">${value}</div>
    </div>
  `;
}

function renderTable() {
  const tbody = document.querySelector("#resultTable tbody");
  tbody.innerHTML = "";

  state.mergedRows.forEach((row, index) => {
    if (!state.ordersMap.has(row.orderId)) return;

    const profit = calcProfit(row);
    const margin = calcMargin(row);
    const status = getStatus(row);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.orderId}</td>
      <td>${row.soldDate || ""}</td>
      <td>${row.skus || '<span class="small">—</span>'}</td>
      <td>${money(row.sold)}</td>
      <td>${money(row.buyerShip)}</td>
      <td>${money(row.tax)}</td>
      <td>
        <input
          class="edit-cell"
          type="number"
          step="0.01"
          value="${row.shipCost.toFixed(2)}"
          data-index="${index}"
          data-field="shipCost"
        />
      </td>
      <td>
        <input
          class="edit-cell"
          type="number"
          step="0.01"
          value="${row.partsCost.toFixed(2)}"
          data-index="${index}"
          data-field="partsCost"
        />
      </td>
      <td>${money(row.estFee)}</td>
      <td class="${profit >= 0 ? "money-good" : "money-bad"}">${money(profit)}</td>
      <td>${margin.toFixed(1)}%</td>
      <td><span class="status ${status.className}">${status.text}</span></td>
    `;
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll(".edit-cell").forEach(input => {
    input.addEventListener("input", onEditCost);
    input.addEventListener("change", onEditCost);
  });
}

function onEditCost(event) {
  const index = Number(event.target.dataset.index);
  const field = event.target.dataset.field;
  const value = num(event.target.value);

  if (!state.mergedRows[index]) return;
  state.mergedRows[index][field] = value;

  renderSummary();
  renderMatchSummary();

  const tbody = document.querySelector("#resultTable tbody");
  const row = state.mergedRows[index];
  const profit = calcProfit(row);
  const margin = calcMargin(row);

  const tr = tbody.children.find ? tbody.children.find(() => false) : null;
  // simplest reliable refresh:
  renderTable();
}

function exportCsv() {
  const rows = state.mergedRows
    .filter(r => state.ordersMap.has(r.orderId))
    .map(r => {
      const status = getStatus(r).text;
      return {
        "Order #": r.orderId,
        "Date": r.soldDate,
        "SKU(s)": r.skus,
        "Sold": r.sold.toFixed(2),
        "Buyer Shipping": r.buyerShip.toFixed(2),
        "Tax": r.tax.toFixed(2),
        "Shipping Cost": r.shipCost.toFixed(2),
        "Parts Cost": r.partsCost.toFixed(2),
        "Estimated eBay Fee": r.estFee.toFixed(2),
        "Profit": calcProfit(r).toFixed(2),
        "Margin %": calcMargin(r).toFixed(2),
        "Status": status
      };
    });

  const csv = Papa.unparse(rows);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "reconciled_pnl.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
