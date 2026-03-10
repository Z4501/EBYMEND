let orders = [];
let shippingMap = {};

function money(n) {
  return `$${(Number(n) || 0).toFixed(2)}`;
}

function num(v) {
  if (v === null || v === undefined) return 0;
  const s = String(v).replace(/[$,]/g, "").trim();
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
}

function findHeaderRow(rows, requiredHeaders) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].map(x => String(x || "").trim());
    const hasAll = requiredHeaders.every(h => row.includes(h));
    if (hasAll) return i;
  }
  return -1;
}

function parseCsvFile(file, requiredHeaders) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      complete: function(results) {
        try {
          const rows = results.data || [];
          const headerRowIndex = findHeaderRow(rows, requiredHeaders);

          if (headerRowIndex === -1) {
            reject(new Error(`Could not find headers: ${requiredHeaders.join(", ")}`));
            return;
          }

          const headers = rows[headerRowIndex].map(h => String(h || "").trim());
          const dataRows = rows.slice(headerRowIndex + 1);

          const records = dataRows
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
      error: function(err) {
        reject(err);
      },
      skipEmptyLines: false
    });
  });
}

function normalizeOrderNumber(v) {
  return String(v || "").trim();
}

function getFirstExisting(obj, keys) {
  for (const key of keys) {
    if (obj[key] !== undefined && obj[key] !== null && String(obj[key]).trim() !== "") {
      return obj[key];
    }
  }
  return "";
}

function estimateFee(sale, buyerShip, tax) {
  const feeBase = sale + buyerShip + tax;
  return (feeBase * 0.136) + ((sale + buyerShip) <= 10 ? 0.30 : 0.40);
}

async function processFiles() {
  const ordersFile = document.getElementById("ordersFile").files[0];
  const shipFile = document.getElementById("shipFile").files[0];

  if (!ordersFile) {
    alert("Upload eBay Orders CSV first.");
    return;
  }

  try {
    const orderRecords = await parseCsvFile(ordersFile, [
      "Order Number",
      "Custom Label"
    ]);

    orders = orderRecords.map(r => {
      const orderNumber = normalizeOrderNumber(getFirstExisting(r, ["Order Number"]));
      const sale = num(getFirstExisting(r, ["Sold For", "Sale Price", "Item Price"]));
      const buyerShip = num(getFirstExisting(r, ["Shipping And Handling", "Shipping", "Shipping price"]));
      const tax = num(getFirstExisting(r, ["eBay Collected Tax", "Tax", "Sales Tax"]));
      const sku = String(getFirstExisting(r, ["Custom Label"])).trim();
      const title = String(getFirstExisting(r, ["Item Title"])).trim();

      return {
        raw: r,
        orderNumber,
        sku,
        title,
        sale,
        buyerShip,
        tax
      };
    }).filter(r => r.orderNumber !== "");

    shippingMap = {};

    if (shipFile) {
      const shipRecords = await parseCsvFile(shipFile, [
        "Order #",
        "Shipping Cost"
      ]);

      shipRecords.forEach(r => {
        const orderNo = normalizeOrderNumber(getFirstExisting(r, ["Order #"]));
        const cost = num(getFirstExisting(r, ["Shipping Cost"]));
        if (orderNo) {
          shippingMap[orderNo] = cost;
        }
      });
    }

    buildTable();
  } catch (err) {
    console.error(err);
    alert("File parsing failed: " + err.message);
  }
}

function buildTable() {
  const tbody = document.querySelector("#resultTable tbody");
  tbody.innerHTML = "";

  let totalSales = 0;
  let totalProfit = 0;
  let totalFees = 0;
  let totalShipping = 0;

  orders.forEach(o => {
    const shipCost = shippingMap[o.orderNumber] || 0;
    const vendorCost = 0;
    const fee = estimateFee(o.sale, o.buyerShip, o.tax);
    const profit = o.sale + o.buyerShip - fee - vendorCost - shipCost;
    const margin = o.sale > 0 ? (profit / o.sale) * 100 : 0;

    totalSales += o.sale;
    totalProfit += profit;
    totalFees += fee;
    totalShipping += shipCost;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${o.orderNumber}</td>
      <td>${o.sku}</td>
      <td>${money(o.sale)}</td>
      <td>${money(o.buyerShip)}</td>
      <td>${money(o.tax)}</td>
      <td>${money(vendorCost)}</td>
      <td>${money(shipCost)}</td>
      <td>${money(fee)}</td>
      <td class="${profit >= 0 ? "good" : "bad"}">${money(profit)}</td>
      <td>${margin.toFixed(1)}%</td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById("summary").innerHTML = `
    Total Orders: ${orders.length}<br>
    Total Sales: ${money(totalSales)}<br>
    Total Estimated Fees: ${money(totalFees)}<br>
    Total Shipping Cost: ${money(totalShipping)}<br>
    Total Profit: ${money(totalProfit)}
  `;
}
