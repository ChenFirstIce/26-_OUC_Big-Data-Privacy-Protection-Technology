# UCI Online Retail 数据只读审计摘要

- 原始文件：`Online Retail.xlsx`（23,715,344 bytes）
- SHA-256：`43465a06f2ccf7c8b5bd2892bc7defb52f97487934fe93b16ae4c3936424676d`
- 工作表：`Online Retail`；交易明细行：541,909
- 过滤后交易明细：396,337；用户：4,334；商品：3,659
- 用户内去重移除的重复商品明细：130,116
- 篮长：min=1，median=35.0，p95=204.0，max=1785

## 唯一主排除原因（固定优先级）

- `invalid_customer_id`：135,080
- `cancelled_invoice`：8,905
- `nonpositive_quantity`：0
- `nonpositive_unit_price`：40
- `invalid_stock_code`：1,547

所有原始行均已唯一归入“保留”或一个主排除原因。原始 Excel 是一行一个交易明细，只有按 `CustomerID` 聚合并在用户内对 `StockCode` 去重后，才形成“一行一个用户、变长购物篮”的逻辑数据。

本阶段仅预览字典序稳定映射 `StockCode -> 1..d`，未生成映射文件、固定单商品数据、GT、隐私机制输出或 MSE。
