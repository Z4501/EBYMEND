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

Features:
- statement-driven month view
- eBay support
- Amazon support
- ShipStation shipping capture
- statement parts cost capture
- in-house SKU cost table fallback
- manual cost overrides
- CSV export

## SKU cost table

Edit `SKU_COST_RULES` near the top of `script.js`.

Example:
- FK032946ESK: 6.13
- WWPOFK032946ESK+: 6.34
- R2-QS1S-MLLF: 8.00

Logic:
1. Statement cost wins
2. If no statement cost, SKU cost rule fills it
3. Manual edits override both
