const state = {
  ordersMap: new Map(),
  shipMap: new Map(),
  statementMap: new Map(),
  mergedRows: []
};

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("processBtn").addEventListener("click", processFiles);
  document.getElementById("exportBtn").addEventListener("click", exportCsv);
});

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
    if (obj[name] !== undefined && obj[name] !== null && String(obj[name]).trim() !== "") {
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

function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = function(err) {
      reject(err);
    };

    reader.readAsArrayBuffer(file);
  });
}

function findStatementHeaderRow(rows) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].map(x =>
      String(x || "").toLowerCase().replace(/\s+/g, "")
    );

    const hasDate = row.some(x => x === "date" || x.includes("date"));
    const hasInvoice = row.some(x => x === "invoice" || x.includes("invoice"));
    const hasPO = row.some(x => x.includes("cust.p.o.") || x.includes("custpo") || x.includes("customerpo"));
    const hasOrigInvAmount = row.some(x =>
      x.includes("originvamount") ||
      x.includes("orig.inv.amount") ||
      x.includes("origamount") ||
      x.includes("amount")
    );

    if (hasDate && hasInvoice && hasPO && hasOrigInvAmount) {
      return i;
    }
  }

  return -1;
}

function buildStatementMapFromRows(rows) {
  const map = new Map();
  const headerIndex = findStatementHeaderRow(rows);

  if (headerIndex === -1) {
    throw new Error("Could not find statement headers.");
  }

  const headers = rows[headerIndex].map(h => String(h || "").trim());
  const dataRows = rows.slice(headerIndex + 1);

  const records = dataRows
    .filter(row => Array.isArray(row) && row.some(cell => String(cell || "").trim() !== ""))
    .map(row => {
      const obj = {};
      headers.forEach((h, idx) => {
        obj[h] = row[idx] !== undefined ? row[idx] : "";
      });
      return obj;
    });

  for (const r of records) {
    const orderId = normalizeOrderId(
      getFirstExisting(r, ["Cust. P.O.", "Cust. PO", "Customer PO"])
    );

    // only keep eBay-style order numbers
    if (!/^\d{2}-\d{5}-\d{5}$/.test(orderId)) {
      continue;
    }

    const invoice = String(getFirstExisting(r, ["Invoice"])).trim();
    const date = String(getFirstExisting(r, ["Date"])).trim();

    const amount = num(
      getFirstExisting(r, [
        "Orig. Inv. Amount",
        "Orig. Amount",
        "Original Amount",
        "Amount"
      ])
    );

    if (!map.has(orderId)) {
      map.set(orderId, {
        orderId,
        invoice,
        date,
        partsCost: 0
      });
    }

    const row = map.get(orderId);
    row.partsCost += amount;

    if (!row.invoice && invoice) row.invoice = invoice;
    if (!row.date && date) row.date = date;
  }

  return map;
}

async function processFiles() {
  const statementFile = document.getElementById("statementFile").files[0];
  const ordersFile = document.getElementById("ordersFile").files[0];
  const shipFile = document.getElementById("shipFile").files[0];

  if (!statementFile) {
    alert("Upload the statement Excel file first. The statement drives the report.");
    return;
  }

  try {
    state.ordersMap = new Map();
    state.shipMap = new Map();
    state.statementMap = new Map();
    state.mergedRows = [];

    // Statement Excel
    if (statementFile.name.toLowerCase().endsWith(".csv")) {
      const statementRecords = await parseCsvFile(statementFile, ["Date", "Invoice", "Orig. Amount"]);
      const csvRows = [
        Object.keys(statementRecords[0] || {}),
        ...statementRecords.map(obj => Object.values(obj))
      ];
      state.statementMap = buildStatementMapFromRows(csvRows);
    } else {
      const statementRows = await parseExcelFile(statementFile);
      state.statementMap = buildStatementMapFromRows(statementRows);
    }

    // eBay CSV
    if (ordersFile) {
      const orderRecords = await parseCsvFile(ordersFile, ["Order Number", "Custom Label"]);
      state.ordersMap = buildOrdersMap(orderRecords);
    }

    // ShipStation CSV
    if (shipFile) {
      const shipRecords = await parseCsvFile(shipFile, ["Order #", "Shipping Cost"]);
      state.shipMap = buildShipMap(shipRecords);
    }

    state.mergedRows = buildMergedRowsFromStatementOnly();
    renderAll();

    document.getElementById("exportBtn").disabled = false;
    alert(`Processed ${state.mergedRows.length} statement-driven orders.`);
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
    const soldDate = String(getFirstExisting(r, ["Sale Date"])).trim();

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

function estimateEbayFee(sold, buyerShip, tax) {
  if (sold <= 0 && buyerShip <= 0) return 0;

  const feeBase = sold + buyerShip + tax;
  const rate = 0.136;
  const orderFee = (sold + buyerShip) <= 10 ? 0.30 : 0.40;
  return (feeBase * rate) + orderFee;
}

function buildMergedRowsFromStatementOnly() {
  const rows = [];

  for (const [orderId, statementRow] of state.statementMap.entries()) {
    const ebay = state.ordersMap.get(orderId);
    const ship = state.shipMap.get(orderId);

    const sold = ebay ? ebay.sold : 0;
    const buyerShip = ebay ? ebay.buyerShip : 0;
    const tax = ebay ? ebay.tax : 0;
    const soldDate = ebay ? ebay.soldDate : "";
    const skus = ebay ? [...ebay.skus].join(" | ") : "";

    rows.push({
      orderId,
      statementDate: statementRow.date || "",
      invoice: statementRow.invoice || "",
      soldDate,
      skus,
      sold,
      buyerShip,
      tax,
      shipCost: ship || 0,
      partsCost: statementRow.partsCost || 0,
      estFee: estimateEbayFee(sold, buyerShip, tax)
    });
  }

  rows.sort((a, b) => b.orderId.localeCompare(a.orderId));
  return rows;
}

function calcProfit(row) {
  return row.sold + row.buyerShip - row.estFee - row.shipCost - row.partsCost;
}

function calcMargin(row) {
  return row.sold > 0 ? (calcProfit(row) / row.sold) * 100 : 0;
}

function getStatus(row) {
  const hasEbay = state.ordersMap.has(row.orderId);
  const hasShip = state.shipMap.has(row.orderId);

  if (hasEbay && hasShip) {
    return { text: "Matched eBay + Ship", className: "status-all" };
  }
  if (hasEbay && !hasShip) {
    return { text: "Missing shipping", className: "status-partial" };
  }
  if (!hasEbay && hasShip) {
    return { text: "Missing eBay sale", className: "status-partial" };
  }
  return { text: "Statement only", className: "status-missing" };
}

function renderAll() {
  renderSummary();
  renderMatchSummary();
  renderTable();
}

function renderSummary() {
  const summary = document.getElementById("summary");
  const rows = state.mergedRows;

  const totalOrders = rows.length;
  const totalSold = rows.reduce((s, r) => s + r.sold, 0);
  const totalBuyerShip = rows.reduce((s, r) => s + r.buyerShip, 0);
  const totalTax = rows.reduce((s, r) => s + r.tax, 0);
  const totalShipCost = rows.reduce((s, r) => s + r.shipCost, 0);
  const totalPartsCost = rows.reduce((s, r) => s + r.partsCost, 0);
  const totalFees = rows.reduce((s, r) => s + r.estFee, 0);
  const totalProfit = rows.reduce((s, r) => s + calcProfit(r), 0);

  summary.innerHTML = `
    ${summaryCard("Statement Orders", totalOrders)}
    ${summaryCard("Sold", money(totalSold))}
    ${summaryCard("Buyer Shipping", money(totalBuyerShip))}
    ${summaryCard("Tax", money(totalTax))}
    ${summaryCard("Shipping Cost", money(totalShipCost))}
    ${summaryCard("Parts Cost", money(totalPartsCost))}
    ${summaryCard("Est. eBay Fees", money(totalFees))}
    ${summaryCard("Profit", money(totalProfit))}
  `;
}

function renderMatchSummary() {
  const matchSummary = document.getElementById("matchSummary");
  const rows = state.mergedRows;

  let matchedAll = 0;
  let missingShip = 0;
  let missingEbay = 0;
  let statementOnly = 0;

  for (const row of rows) {
    const hasEbay = state.ordersMap.has(row.orderId);
    const hasShip = state.shipMap.has(row.orderId);

    if (hasEbay && hasShip) matchedAll++;
    else if (hasEbay && !hasShip) missingShip++;
    else if (!hasEbay && hasShip) missingEbay++;
    else statementOnly++;
  }

  matchSummary.innerHTML = `
    ${summaryCard("Matched eBay + Ship", matchedAll)}
    ${summaryCard("Missing shipping", missingShip)}
    ${summaryCard("Missing eBay sale", missingEbay)}
    ${summaryCard("Statement only", statementOnly)}
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
    const profit = calcProfit(row);
    const margin = calcMargin(row);
    const status = getStatus(row);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.orderId}</td>
      <td>${row.statementDate || row.soldDate || ""}</td>
      <td>${row.invoice || '<span class="small">—</span>'}</td>
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

  renderAll();
}

function exportCsv() {
  const rows = state.mergedRows.map(r => ({
    "Order #": r.orderId,
    "Statement Date": r.statementDate,
    "Invoice": r.invoice,
    "SKU(s)": r.skus,
    "Sold": r.sold.toFixed(2),
    "Buyer Shipping": r.buyerShip.toFixed(2),
    "Tax": r.tax.toFixed(2),
    "Shipping Cost": r.shipCost.toFixed(2),
    "Parts Cost": r.partsCost.toFixed(2),
    "Estimated eBay Fee": r.estFee.toFixed(2),
    "Profit": calcProfit(r).toFixed(2),
    "Margin %": calcMargin(r).toFixed(2),
    "Status": getStatus(r).text
  }));

  const csv = Papa.unparse(rows);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "statement_driven_pnl.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
