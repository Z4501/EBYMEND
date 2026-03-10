# Statement-Driven Marketplace P&L Viewer

Upload:
- Statement Excel (.xlsx / .xls / .csv)
- eBay Orders CSV
- Amazon Transactions (.xlsx / .xls / .csv)
- ShipStation CSV

The statement drives which orders appear.

Matching keys:
- Statement `Cust. P.O.`
- eBay `Order Number`
- Amazon `Order ID`
- ShipStation `Order #`

Rules:
- eBay orders use eBay sales + ShipStation shipping
- Amazon orders use Amazon Order Payment + Amazon Shipping Services Purchased
- returns/refunds/reversals are ignored for Amazon in this version
- shipping cost and parts cost remain editable
