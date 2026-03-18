---
name: sales-recorder
description: 记录每日销售数据到Excel表格。用户发送商品销售信息（日期、商品、数量、单价、快递单数、快递费）时，自动录入到桌面「拿货记数/{表格名}.xlsx」，包含公式自动计算小计。
---

# 销售记数

## 表格位置

| 表格 | 路径 |
|------|------|
| 弟弟 | `/Users/mac/Desktop/拿货记数/弟弟.xlsx` |
| 央央 | `/Users/mac/Desktop/拿货记数/央央.xlsx` |
| 宝宝 | `/Users/mac/Desktop/拿货记数/宝宝.xlsx` |
| 超宝 | `/Users/mac/Desktop/拿货记数/超宝.xlsx` |
| 薛泽凯 | `/Users/mac/Desktop/拿货记数/薛泽凯.xlsx` |

## 规则

**用户说记到哪个表格就记到哪个表格。**

## 表格结构

| A列 | B列 | C列 | D列 | E列 | F列 | G列 | H列 |
|-----|-----|-----|-----|-----|-----|-----|-----|
| 日期 | 商品 | 数量 | 单价 | 合计 | 快递单数 | 价格 | 合计 |

- E列公式：`=C列*D列`
- H列公式：`=F列*G列`

## 解析用户消息格式

用户可能发送的格式：
- `润滑油 数量10价格5快递10价格2.8 填写到央央`
- `风流果10价格5.5快递3价格2.8 填写到央央`
- `润滑油10价格5快递2价格2.8 填写到弟弟`
- `快递10价格10 填写到薛泽凯`
- `高潮液数量10价格5 填写到薛泽凯`

解析逻辑：
1. 提取目标表格：查找"弟弟"、"央央"、"宝宝"、"超宝"、"薛泽凯"等关键词
2. 提取商品：匹配商品名称
3. 提取数量：`数量(\d+)` 或直接数字
4. 提取单价：`价格(\d+\.?\d*)`
5. 提取快递单数：`快递(\d+)`
6. 提取快递价格：第二个`价格(\d+\.?\d*)`或`快递价格(\d+\.?\d*)`

## 录入流程

```python
from openpyxl import load_workbook

file_path = f'/Users/mac/Desktop/拿货记数/{target}.xlsx'

wb = load_workbook(file_path)
ws = wb.active

# 找空行：必须整行都是空的才能用（A到H列都为空）
next_row = 2
for row in range(2, 110):
    is_empty = all(ws.cell(row, col).value is None for col in range(1, 9))
    if is_empty:
        next_row = row
        break

# 写入数据
ws.cell(next_row, 2).value = product  # 商品
ws.cell(next_row, 3).value = qty      # 数量
ws.cell(next_row, 4).value = price    # 单价
ws.cell(next_row, 5).value = f"=C{next_row}*D{next_row}"  # 合计

if expr_count:
    ws.cell(next_row, 6).value = expr_count  # 快递单数
    ws.cell(next_row, 7).value = expr_price  # 快递价格
    ws.cell(next_row, 8).value = f"=F{next_row}*G{next_row}"  # 快递合计

wb.save(file_path)

# 录入后检查：删除刚录入行后面的空行，保持紧凑
if next_row + 1 <= 109:
    next_is_empty = all(ws.cell(next_row + 1, col).value is None for col in range(1, 9))
    if next_is_empty:
        ws.delete_rows(next_row + 1)
        wb.save(file_path)
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
