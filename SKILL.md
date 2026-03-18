---
name: sales-recorder
description: 记录每日销售数据到Excel表格。用于用户发送商品销售信息（日期、商品、数量、单价、快递单数、快递费）时，自动录入到桌面「拿货记数/弟弟.xlsx」表格，包含公式自动计算小计。
---

# 销售记数

## 表格位置

| 表格 | 路径 |
|------|------|
| 弟弟 | `/Users/mac/Desktop/拿货记数/弟弟.xlsx` |
| 央央 | `/Users/mac/Desktop/拿货记数/央央.xlsx` |
| 宝宝 | `/Users/mac/Desktop/拿货记数/宝宝.xlsx` |
| 超宝 | `/Users/mac/Desktop/拿货记数/超宝.xlsx` |

## 规则

**用户说记到哪个表格就记到哪个表格，不再自动判断商品对应哪个表格。**

- 用户说"记到央央" → 录入央央.xlsx
- 用户说"记到弟弟" → 录入弟弟.xlsx
- 以此类推

## 表格结构

| A列 | B列 | C列 | D列 | E列 | F列 | G列 | H列 |
|-----|-----|-----|-----|-----|-----|-----|-----|
| 日期 | 商品 | 数量 | 单价 | 合计 | 快递单数 | 价格 | 合计 |

- E列公式：`=C列*D列`
- H列公式：`=F列*G列`
- 第101行有汇总公式：`=SUM(E2:E100)`、`=SUM(F2:F100)`、`=SUM(H2:H100)`

## 录入规则

1. **找空行**：从第2行开始扫描，找到B列（商品）第一个为空的行
2. **日期**：当天第一条记录写日期（如"18号"），后续同天记录留空
3. **商品**：必填
4. **数量**：必填
5. **单价**：必填
6. **快递单数**：用户未提及则留空
7. **快递价格**：用户未提及则留空
8. **公式**：C列和D列有值后，E列自动写入公式 `=C{row}*D{row}`
9. **快递公式**：F列和G列都有值后，H列自动写入公式 `=F{row}*G{row}`

## 解析用户消息格式

用户可能发送的格式：
- `润滑油 数量10价格5快递10价格2.8 填写到央央`
- `风流果10价格5.5快递3价格2.8 填写到央央`
- `润滑油10价格5快递2价格2.8 填写到弟弟`

解析逻辑：
1. 提取目标表格：查找"弟弟"、"央央"、"宝宝"、"超宝"等关键词
2. 提取商品：匹配商品名称
3. 提取数量：`数量(\d+)` 或直接数字
4. 提取单价：`价格(\d+\.?\d*)`
5. 提取快递单数：`快递(\d+)`
6. 提取快递价格：`快递价格(\d+\.?\d*)` 或第二个`价格(\d+\.?\d*)`

## 录入流程

```python
from openpyxl import load_workbook

file_path = f'/Users/mac/Desktop/拿货记数/{target}.xlsx'

wb = load_workbook(file_path)
ws = wb.active

# 找空行：必须整行都是空的才能用
next_row = 2
for row in range(2, 110):
    # 检查整行是否完全空白（A到H列都为空）
    is_empty = all(ws.cell(row, col).value is None for col in range(1, 9))
    if is_empty:
        next_row = row
        break
    
    # 如果这一行有任何数据（包括之前被清空的），继续往下找
    # 确保不会覆盖已有数据的行

# 如果所有行都有数据，则追加到最后一行之后
if not any(ws.cell(next_row, col).value is None for col in range(1, 9)):
    next_row = 110  # 或者找到最后一个有数据的行+1

# 写入数据
ws.cell(next_row, 1).value = date或None  # A列
ws.cell(next_row, 2).value = product     # B列
ws.cell(next_row, 3).value = qty         # C列
ws.cell(next_row, 4).value = price      # D列
ws.cell(next_row, 5).value = f"=C{next_row}*D{next_row}"  # E列

if expr_count:
    ws.cell(next_row, 6).value = expr_count  # F列
    ws.cell(next_row, 7).value = expr_price  # G列
    ws.cell(next_row, 8).value = f"=F{next_row}*G{next_row}"  # H列

wb.save('/Users/mac/Desktop/拿货记数/弟弟.xlsx')
```

## 确认回复

录入完成后，回复格式：
```
✅ 已添加到 {表格名}.xlsx 第{row}行！

| 商品 | 数量 | 单价 | 小计 | 快递单数 |
|------|------|------|------|----------|
| {商品} | {数量} | {单价} | ¥{小计} | {快递单数或"（留空）"} |
```

然后提示用户继续：`继续 📝`
