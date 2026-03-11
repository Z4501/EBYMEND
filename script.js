const state = {
  ordersMap: new Map(),
  amazonMap: new Map(),
  shipMap: new Map(),
  statementMap: new Map(),
  statementAltMap: new Map(),
  mergedRows: []
};

const SKU_COST_RULES = {
  "FK032946ESK": 6.13,
  "WWPOFK032946ESK+": 6.34,
  "R2-QS1S-MLLF": 8.0
};

const AMAZON_TITLE_COST_RULES = {
  "GM TH350 TRANSMISSION SUPER MAX FILTER K": 6.13,
  "GM TH350 TRANSMISSION SUPER SEAL FILTER": 6.34,
  "GM TH400 TRANSMISSION SUPER MAX FILTER K": 8.0
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

function extractFirstNumber(value) {
  if (value === null || value === undefined) return "";
  const match = String(value).match(/\d+/);
  return match ? match[0] : "";
}

function extractEmbeddedMarketplaceOrderId(value) {
  const text = String(value || "").trim();

  const ebayMatch = text.match(/\b\d{2}-\d{5}-\d{5}\b/);
  if (ebayMatch) return ebayMatch[0];

  const amazonMatch = text.match(/\b\d{3}-\d{7}-\d{7}\b/);
  if (amazonMatch) return amazonMatch[0];

  return "";
}

function buildStatementMatchKey(rawValue) {
  const raw = String(rawValue || "").trim();
  const normalized = normalizeOrderId(raw);
  const embeddedMarketplaceOrderId = extractEmbeddedMarketplaceOrderId(raw);
  const numericOnly = extractFirstNumber(raw);

  if (embeddedMarketplaceOrderId) {
    return {
      raw,
      normalized: embeddedMarketplaceOrderId,
      numericOnly,
      isMarketplaceOrder: true,
      embeddedMarketplaceOrderId
    };
  }

  if (isEbayOrderId(normalized) || isAmazonOrderId(normalized)) {
    return {
      raw,
      normalized,
      numericOnly,
      isMarketplaceOrder: true,
      embeddedMarketplaceOrderId: normalized
    };
  }

  return {
    raw,
    normalized,
    numericOnly,
    isMarketplaceOrder: false,
    embeddedMarketplaceOrderId: ""
  };
}

function isEbayOrderId(orderId) {
  return /^\d{2}-\d{5}-\d{5}$/.test(orderId);
}

function isAmazonOrderId(orderId) {
  return /^\d{3}-\d{7}-\d{7}$/.test(orderId);
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

function rowsToObjects(rows, requiredHeaders) {
  const headerIndex = findHeaderRow(rows, requiredHeaders);

  if (headerIndex === -1) {
    throw new Error(`Could not find required headers: ${requiredHeaders.join(", ")}`);
  }

  const headers = rows[headerIndex].map(h => String(h || "").trim());
  return rows
    .slice(headerIndex + 1)
    .filter(row => Array.isArray(row) && row.some(cell => String(cell || "").trim() !== ""))
    .map(row => {
      const obj = {};
      headers.forEach((h, idx) => {
        obj[h] = row[idx] !== undefined ? row[idx] : "";
      });
      return obj;
    });
}

function findStatementHeaderRow(rows) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].map(x =>
      String(x || "").toLowerCase().replace(/\s+/g, "")
    );

    const hasDate = row.some(x => x === "date" || x.includes("date"));
    const hasInvoice = row.some(x => x === "invoice" || x.includes("invoice"));
    const hasPO = row.some(x =>
      x.includes("cust.p.o.") ||
      x.includes("custpo") ||
      x.includes("customerpo")
    );
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

function buildStatementMapsFromRows(rows) {
  const statementMap = new Map();
  const statementAltMap = new Map();
  const linkedAltKeys = new Set();

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
    const custPoRaw = getFirstExisting(r, ["Cust. P.O.", "Cust. PO", "Customer PO"]);
    const matchInfo = buildStatementMatchKey(custPoRaw);

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

    const primaryId = matchInfo.embeddedMarketplaceOrderId || matchInfo.normalized;
    const altNumericId = matchInfo.numericOnly;

    if (matchInfo.isMarketplaceOrder && primaryId) {
      if (!statementMap.has(primaryId)) {
        statementMap.set(primaryId, {
          orderId: primaryId,
          rawPO: matchInfo.raw,
          normalizedPO: primaryId,
          numericPO: altNumericId,
          invoice,
          date,
          partsCost: 0,
          matchSource: "order"
        });
      }

      const row = statementMap.get(primaryId);
      row.partsCost += amount;
      if (!row.invoice && invoice) row.invoice = invoice;
      if (!row.date && date) row.date = date;
      if (!row.rawPO && matchInfo.raw) row.rawPO = matchInfo.raw;
      if (!row.numericPO && altNumericId) row.numericPO = altNumericId;

      if (altNumericId && altNumericId !== primaryId) {
        linkedAltKeys.add(altNumericId);
      }
    } else if (altNumericId) {
      if (!statementAltMap.has(altNumericId)) {
        statementAltMap.set(altNumericId, {
          orderId: "",
          rawPO: matchInfo.raw,
          normalizedPO: primaryId,
          numericPO: altNumericId,
          invoice,
          date,
          partsCost: 0,
          matchSource: "po"
        });
      }

      const row = statementAltMap.get(altNumericId);
      row.partsCost += amount;
      if (!row.invoice && invoice) row.invoice = invoice;
      if (!row.date && date) row.date = date;
      if (!row.rawPO && matchInfo.raw) row.rawPO = matchInfo.raw;
      if (!row.numericPO && altNumericId) row.numericPO = altNumericId;
    }
  }

  for (const linkedKey of linkedAltKeys) {
    statementAltMap.delete(linkedKey);
  }

  return { statementMap, statementAltMap };
}

function buildOrdersMap(records) {
  const map = new Map();

  for (const r of records) {
    const orderId = normalizeOrderId(getFirstExisting(r, ["Order Number"]));
    if (!orderId) continue;

    const qty = Math.max(1, num(getFirstExisting(r, ["Quantity", "Qty"])));
    const soldFor = num(getFirstExisting(r, ["Sold For"]));
    const buyerShip = num(getFirstExisting(r, ["Shipping And Handling"]));
    const tax = num(getFirstExisting(r, ["eBay Collected Tax", "eBay Collect And Remit Tax", "Sales Tax"]));
    const sku = String(getFirstExisting(r, ["Custom Label"])).trim();
    const soldDate = String(getFirstExisting(r, ["Sale Date"])).trim();

    const lineSold = qty > 1 ? soldFor * qty : soldFor;

    if (!map.has(orderId)) {
      map.set(orderId, {
        marketplace: "eBay",
        orderId,
        soldDate,
        sold: 0,
        buyerShip: 0,
        tax: 0,
        qty: 0,
        fee: 0,
        skus: [],
        skuQtyMap: {}
      });
    }

    const row = map.get(orderId);
    row.sold += lineSold;
    row.buyerShip += buyerShip;
    row.tax += tax;
    row.qty += qty;

    if (sku) {
      row.skus.push(sku);
      row.skuQtyMap[sku] = (row.skuQtyMap[sku] || 0) + qty;
    }
  }

  for (const row of map.values()) {
    row.fee = estimateEbayFee(row.sold, row.buyerShip, row.tax);
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

function extractKnownSkusFromText(text) {
  const found = [];
  const upper = String(text || "").toUpperCase();

  for (const sku of Object.keys(SKU_COST_RULES)) {
    if (upper.includes(sku.toUpperCase())) {
      found.push(sku);
    }
  }

  return found;
}

function getAmazonTitleCost(detailText, qty) {
  const text = String(detailText || "").toUpperCase();

  for (const [fragment, cost] of Object.entries(AMAZON_TITLE_COST_RULES)) {
    if (text.includes(fragment.toUpperCase())) {
      return {
        matched: true,
        cost: cost * Math.max(1, qty),
        fragment
      };
    }
  }

  return {
    matched: false,
    cost: 0,
    fragment: ""
  };
}

function buildAmazonMap(records) {
  const map = new Map();

  for (const r of records) {
    const orderId = normalizeOrderId(getFirstExisting(r, ["Order ID"]));
    if (!orderId || !isAmazonOrderId(orderId)) continue;

    const txType = String(getFirstExisting(r, ["Transaction type"])).trim();
    const txTypeLower = txType.toLowerCase();

    if (
      txTypeLower.includes("refund") ||
      txTypeLower.includes("return") ||
      txTypeLower.includes("reversal") ||
      txTypeLower.includes("chargeback")
    ) {
      continue;
    }

    if (!map.has(orderId)) {
      map.set(orderId, {
        marketplace: "Amazon",
        orderId,
        soldDate: "",
        sold: 0,
        buyerShip: 0,
        tax: 0,
        qty: 0,
        fee: 0,
        shipCost: 0,
        skus: [],
        skuQtyMap: {},
        titleRuleCost: 0,
        titleRuleMatched: false
      });
    }

    const row = map.get(orderId);
    const dateVal = String(getFirstExisting(r, ["Date"])).trim();
    if (!row.soldDate && dateVal) row.soldDate = dateVal;

    if (txType === "Order Payment") {
      const productCharges = num(getFirstExisting(r, ["Total product charges"]));
      const qty = Math.max(1, num(getFirstExisting(r, ["Quantity"])));
      const detail = String(getFirstExisting(r, ["Product Details"])).trim();

      row.sold += productCharges;
      row.fee += Math.abs(num(getFirstExisting(r, ["Amazon fees"])));
      row.qty += qty;

      if (detail && detail.toLowerCase() !== "billing") {
        row.skus.push(detail);

        const matchedSkus = extractKnownSkusFromText(detail);
        if (matchedSkus.length) {
          const splitQty = qty / matchedSkus.length;
          matchedSkus.forEach(sku => {
            row.skuQtyMap[sku] = (row.skuQtyMap[sku] || 0) + splitQty;
          });
        }

        const titleCost = getAmazonTitleCost(detail, qty);
        if (titleCost.matched) {
          row.titleRuleCost += titleCost.cost;
          row.titleRuleMatched = true;
        }
      }
    } else if (txType === "Shipping services purchased through Amazon") {
      row.shipCost += Math.abs(num(getFirstExisting(r, ["Total (USD)"])));
    }
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

function getSkuCostFromSkuQtyMap(skuQtyMap) {
  let total = 0;
  let matchedAny = false;

  for (const [sku, qty] of Object.entries(skuQtyMap || {})) {
    if (SKU_COST_RULES[sku] !== undefined) {
      total += SKU_COST_RULES[sku] * qty;
      matchedAny = true;
    }
  }

  return {
    total,
    matchedAny
  };
}

async function processFiles() {
  const statementFile = document.getElementById("statementFile").files[0];
  const ordersFile = document.getElementById("ordersFile").files[0];
  const amazonFile = document.getElementById("amazonFile").files[0];
  const shipFile = document.getElementById("shipFile").files[0];

  if (!statementFile) {
    alert("Upload the statement Excel file first. The statement drives the report.");
    return;
  }

  try {
    state.ordersMap = new Map();
    state.amazonMap = new Map();
    state.shipMap = new Map();
    state.statementMap = new Map();
    state.statementAltMap = new Map();
    state.mergedRows = [];

    if (statementFile.name.toLowerCase().endsWith(".csv")) {
      const statementRecords = await parseCsvFile(statementFile, ["Date", "Invoice"]);
      const csvRows = [
        Object.keys(statementRecords[0] || {}),
        ...statementRecords.map(obj => Object.values(obj))
      ];
      const maps = buildStatementMapsFromRows(csvRows);
      state.statementMap = maps.statementMap;
      state.statementAltMap = maps.statementAltMap;
    } else {
      const statementRows = await parseExcelFile(statementFile);
      const maps = buildStatementMapsFromRows(statementRows);
      state.statementMap = maps.statementMap;
      state.statementAltMap = maps.statementAltMap;
    }

    if (ordersFile) {
      const orderRecords = await parseCsvFile(ordersFile, ["Order Number", "Custom Label"]);
      state.ordersMap = buildOrdersMap(orderRecords);
    }

    if (amazonFile) {
      let amazonRows;
      if (amazonFile.name.toLowerCase().endsWith(".csv")) {
        const amazonRecords = await parseCsvFile(amazonFile, ["Transaction type", "Order ID"]);
        amazonRows = [
          Object.keys(amazonRecords[0] || {}),
          ...amazonRecords.map(obj => Object.values(obj))
        ];
      } else {
        amazonRows = await parseExcelFile(amazonFile);
      }

      const amazonRecords = rowsToObjects(amazonRows, ["Transaction type", "Order ID"]);
      state.amazonMap = buildAmazonMap(amazonRecords);
    }

    if (shipFile) {
      const shipRecords = await parseCsvFile(shipFile, ["Order #", "Shipping Cost"]);
      state.shipMap = buildShipMap(shipRecords);
    }

    state.mergedRows = buildMergedRows();
    renderAll();

    document.getElementById("exportBtn").disabled = false;
    alert(`Processed ${state.mergedRows.length} orders.`);
  } catch (err) {
    console.error(err);
    alert(`Processing failed: ${err.message}`);
  }
}

function buildMergedRows() {
  const rows = [];
  const includedOrderIds = new Set();
  const usedStatementPrimary = new Set();
  const usedStatementAlt = new Set();

  for (const [orderId, ebay] of state.ordersMap.entries()) {
    const statementRow = state.statementMap.get(orderId);
    const shipstationCost = state.shipMap.get(orderId) || 0;

    let matchedStatement = statementRow || null;
    let statementMatchType = statementRow ? "Order #" : "";
    let partsCost = statementRow ? (statementRow.partsCost || 0) : 0;
    let invoice = statementRow ? (statementRow.invoice || "") : "";
    let statementDate = statementRow ? (statementRow.date || "") : "";
    let statementPO = statementRow ? (statementRow.rawPO || "") : "";

    if (!matchedStatement) {
      const alt = state.statementAltMap.get(orderId);
      if (alt) {
        matchedStatement = alt;
        statementMatchType = "Cust. P.O. #";
        partsCost = alt.partsCost || 0;
        invoice = alt.invoice || "";
        statementDate = alt.date || "";
        statementPO = alt.rawPO || "";
        usedStatementAlt.add(orderId);
      }
    } else {
      usedStatementPrimary.add(orderId);
    }

    let costSource = partsCost > 0 ? "Statement" : "None";

    if (!(partsCost > 0)) {
      const skuCost = getSkuCostFromSkuQtyMap(ebay.skuQtyMap);
      if (skuCost.matchedAny) {
        partsCost = skuCost.total;
        costSource = "SKU Rule";
      }
    }

    rows.push({
      marketplace: "eBay",
      orderId,
      statementDate,
      invoice,
      statementPO,
      statementMatchType,
      soldDate: ebay.soldDate || "",
      skus: ebay.skus.join(" | "),
      qty: ebay.qty || 0,
      sold: ebay.sold || 0,
      buyerShip: ebay.buyerShip || 0,
      tax: ebay.tax || 0,
      shipCost: shipstationCost,
      partsCost,
      costSource,
      fee: ebay.fee || 0
    });

    includedOrderIds.add(orderId);
  }

  for (const [orderId, amazon] of state.amazonMap.entries()) {
    if (includedOrderIds.has(orderId)) continue;

    const statementRow = state.statementMap.get(orderId);
    let matchedStatement = statementRow || null;
    let statementMatchType = statementRow ? "Order #" : "";
    let partsCost = statementRow ? (statementRow.partsCost || 0) : 0;
    let invoice = statementRow ? (statementRow.invoice || "") : "";
    let statementDate = statementRow ? (statementRow.date || "") : "";
    let statementPO = statementRow ? (statementRow.rawPO || "") : "";

    if (!matchedStatement) {
      const alt = state.statementAltMap.get(orderId);
      if (alt) {
        matchedStatement = alt;
        statementMatchType = "Cust. P.O. #";
        partsCost = alt.partsCost || 0;
        invoice = alt.invoice || "";
        statementDate = alt.date || "";
        statementPO = alt.rawPO || "";
        usedStatementAlt.add(orderId);
      }
    } else {
      usedStatementPrimary.add(orderId);
    }

    let costSource = partsCost > 0 ? "Statement" : "None";

    if (!(partsCost > 0)) {
      const skuCost = getSkuCostFromSkuQtyMap(amazon.skuQtyMap);
      if (skuCost.matchedAny) {
        partsCost = skuCost.total;
        costSource = "SKU Rule";
      } else if (amazon.titleRuleMatched) {
        partsCost = amazon.titleRuleCost;
        costSource = "Amazon Title Rule";
      }
    }

    rows.push({
      marketplace: "Amazon",
      orderId,
      statementDate,
      invoice,
      statementPO,
      statementMatchType,
      soldDate: amazon.soldDate || "",
      skus: amazon.skus.join(" | "),
      qty: amazon.qty || 0,
      sold: amazon.sold || 0,
      buyerShip: amazon.buyerShip || 0,
      tax: amazon.tax || 0,
      shipCost: amazon.shipCost || 0,
      partsCost,
      costSource,
      fee: amazon.fee || 0
    });

    includedOrderIds.add(orderId);
  }

  for (const [orderId, statementRow] of state.statementMap.entries()) {
    if (includedOrderIds.has(orderId)) continue;
    if (usedStatementPrimary.has(orderId)) continue;

    rows.push({
      marketplace: isAmazonOrderId(orderId) ? "Amazon" : isEbayOrderId(orderId) ? "eBay" : "Unknown",
      orderId,
      statementDate: statementRow.date || "",
      invoice: statementRow.invoice || "",
      statementPO: statementRow.rawPO || "",
      statementMatchType: "Order #",
      soldDate: "",
      skus: "",
      qty: 0,
      sold: 0,
      buyerShip: 0,
      tax: 0,
      shipCost: state.shipMap.get(orderId) || 0,
      partsCost: statementRow.partsCost || 0,
      costSource: (statementRow.partsCost || 0) > 0 ? "Statement" : "None",
      fee: 0
    });
  }

  for (const [numericKey, statementRow] of state.statementAltMap.entries()) {
    if (usedStatementAlt.has(numericKey)) continue;

    rows.push({
      marketplace: "Unknown",
      orderId: numericKey,
      statementDate: statementRow.date || "",
      invoice: statementRow.invoice || "",
      statementPO: statementRow.rawPO || "",
      statementMatchType: "Cust. P.O. #",
      soldDate: "",
      skus: "",
      qty: 0,
      sold: 0,
      buyerShip: 0,
      tax: 0,
      shipCost: 0,
      partsCost: statementRow.partsCost || 0,
      costSource: (statementRow.partsCost || 0) > 0 ? "Statement" : "None",
      fee: 0
    });
  }

  rows.sort((a, b) => String(b.orderId).localeCompare(String(a.orderId)));
  return rows;
}

function calcProfit(row) {
  return row.sold + row.buyerShip - row.fee - row.shipCost - row.partsCost;
}

function calcMargin(row) {
  return row.sold > 0 ? (calcProfit(row) / row.sold) * 100 : 0;
}

function getStatus(row) {
  if (row.marketplace === "eBay") {
    const hasEbay = state.ordersMap.has(row.orderId);
    const hasShip = state.shipMap.has(row.orderId);
    const hasStatementByOrder = state.statementMap.has(row.orderId);
    const hasStatementByPO = state.statementAltMap.has(row.orderId);

    if (hasEbay && hasStatementByOrder && hasShip) {
      return { text: "Matched eBay + Ship", className: "status-all" };
    }
    if (hasEbay && hasStatementByPO && hasShip) {
      return { text: "Matched by PO + Ship", className: "status-all" };
    }
    if (hasEbay && hasStatementByOrder && !hasShip) {
      return { text: "Matched by order / Missing shipping", className: "status-partial" };
    }
    if (hasEbay && hasStatementByPO && !hasShip) {
      return { text: "Matched by PO / Missing shipping", className: "status-partial" };
    }
    if (hasEbay) {
      return { text: "eBay only", className: "status-partial" };
    }
    return { text: "Statement only", className: "status-missing" };
  }

  if (row.marketplace === "Amazon") {
    const hasAmazon = state.amazonMap.has(row.orderId);
    const hasStatementByOrder = state.statementMap.has(row.orderId);
    const hasStatementByPO = state.statementAltMap.has(row.orderId);

    if (hasAmazon && hasStatementByOrder) {
      return { text: "Matched Amazon", className: "status-all" };
    }
    if (hasAmazon && hasStatementByPO) {
      return { text: "Matched Amazon by PO", className: "status-all" };
    }
    if (!hasStatementByOrder && !hasStatementByPO && hasAmazon && (row.costSource === "SKU Rule" || row.costSource === "Amazon Title Rule")) {
      return { text: "Amazon SKU-only", className: "status-partial" };
    }
    if (hasAmazon) {
      return { text: "Amazon only", className: "status-partial" };
    }
    return { text: "Statement only", className: "status-missing" };
  }

  if (row.statementMatchType === "Cust. P.O. #") {
    return { text: "Statement only (PO)", className: "status-missing" };
  }

  return { text: "Unknown order type", className: "status-missing" };
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
  const totalFees = rows.reduce((s, r) => s + r.fee, 0);
  const totalProfit = rows.reduce((s, r) => s + calcProfit(r), 0);

  summary.innerHTML = `
    ${summaryCard("Orders", totalOrders)}
    ${summaryCard("Sold", money(totalSold))}
    ${summaryCard("Buyer Shipping", money(totalBuyerShip))}
    ${summaryCard("Tax", money(totalTax))}
    ${summaryCard("Shipping Cost", money(totalShipCost))}
    ${summaryCard("Parts Cost", money(totalPartsCost))}
    ${summaryCard("Marketplace Fees", money(totalFees))}
    ${summaryCard("Profit", money(totalProfit))}
  `;
}

function renderMatchSummary() {
  const matchSummary = document.getElementById("matchSummary");
  const rows = state.mergedRows;

  let ebayMatched = 0;
  let ebayMissingShip = 0;
  let ebayStatementOnly = 0;
  let amazonMatched = 0;
  let amazonSkuOnly = 0;
  let amazonStatementOnly = 0;

  for (const row of rows) {
    const statusText = getStatus(row).text;

    if (row.marketplace === "eBay") {
      if (statusText === "Matched eBay + Ship" || statusText === "Matched by PO + Ship") ebayMatched++;
      else if (statusText === "Matched by order / Missing shipping" || statusText === "Matched by PO / Missing shipping") ebayMissingShip++;
      else if (statusText === "Statement only") ebayStatementOnly++;
    } else if (row.marketplace === "Amazon") {
      if (statusText === "Matched Amazon" || statusText === "Matched Amazon by PO") amazonMatched++;
      else if (statusText === "Amazon SKU-only") amazonSkuOnly++;
      else if (statusText === "Statement only") amazonStatementOnly++;
    }
  }

  matchSummary.innerHTML = `
    ${summaryCard("eBay matched", ebayMatched)}
    ${summaryCard("eBay missing shipping", ebayMissingShip)}
    ${summaryCard("eBay statement only", ebayStatementOnly)}
    ${summaryCard("Amazon matched", amazonMatched)}
    ${summaryCard("Amazon SKU-only", amazonSkuOnly)}
    ${summaryCard("Amazon statement only", amazonStatementOnly)}
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
      <td>${row.marketplace}</td>
      <td>${row.orderId}</td>
      <td>${row.statementDate || row.soldDate || ""}</td>
      <td>${row.invoice || '<span class="small">—</span>'}</td>
      <td>${row.statementPO || '<span class="small">—</span>'}</td>
      <td>${row.statementMatchType || '<span class="small">—</span>'}</td>
      <td>${row.skus || '<span class="small">—</span>'}</td>
      <td>${row.qty || ""}</td>
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
      <td>${row.costSource || ""}</td>
      <td>${money(row.fee)}</td>
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

  if (field === "partsCost") {
    state.mergedRows[index].costSource = "Manual";
  }

  renderAll();
}

function exportCsv() {
  const rows = state.mergedRows.map(r => ({
    "Marketplace": r.marketplace,
    "Order #": r.orderId,
    "Statement Date": r.statementDate,
    "Invoice": r.invoice,
    "Statement PO": r.statementPO,
    "Statement Match Type": r.statementMatchType,
    "SKU(s)": r.skus,
    "Qty": r.qty,
    "Sold": r.sold.toFixed(2),
    "Buyer Shipping": r.buyerShip.toFixed(2),
    "Tax": r.tax.toFixed(2),
    "Shipping Cost": r.shipCost.toFixed(2),
    "Parts Cost": r.partsCost.toFixed(2),
    "Cost Source": r.costSource,
    "Fee": r.fee.toFixed(2),
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
