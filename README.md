# Marketplace P&L Viewer

Upload:
- Statement Excel (.xlsx / .xls / .csv)
- eBay Orders CSV
- Amazon Transactions (.xlsx / .xls / .csv)
- ShipStation CSV

Matching keys:
- Statement `Cust. P.O.`
- eBay `Order Number`
- Amazon `Order ID`
- ShipStation `Order #`

Logic:
1. Statement cost wins
2. SKU cost rules fill missing costs
3. Manual edits override both

Amazon-only SKU-rule orders can appear even if they are not on the Alto statement.
