import pandas as pd
from openpyxl import Workbook

# 手动解析（如果数据已经是文本层级，你可以读取文本或图片OCR，这里假设我们用结构化list模拟）
data = [
    {"level": 1, "value": "720000035"},
    {"level": 2, "value": "PT000020"},
    {"level": 3, "value": "A000500301"},
    {"level": 4, "value": "SC500063"},
    {"level": 5, "value": "SRE0001"},
    {"level": 5, "value": "SRE0002"},
    {"level": 3, "value": "A000500302"},
    {"level": 4, "value": "SC500064"},
    {"level": 5, "value": "SRE0003"},
    # 模拟多层...
]

# 状态跟踪
current = {'sold_to': None, 'ship_to': None, 'customer_pn': None, 'company_pn': None}
rows = []
sre_list = []

for entry in data:
    if entry['level'] == 1:
        current['sold_to'] = entry['value']
    elif entry['level'] == 2:
        current['ship_to'] = entry['value']
    elif entry['level'] == 3:
        current['customer_pn'] = entry['value']
    elif entry['level'] == 4:
        current['company_pn'] = entry['value']
    elif entry['level'] == 5:
        sre_list.append(entry['value'])
        
    # 如果下一个是 level 3 或结束，写入一次
    next_entry = data[data.index(entry)+1] if data.index(entry)+1 < len(data) else {"level": -1}
    if (entry['level'] == 5 and next_entry['level'] <= 3) or next_entry['level'] == -1:
        # 拆 SRE，每2个一组，超出则拆行
        for i in range(0, len(sre_list), 2):
            row = [
                current['sold_to'],
                current['ship_to'],
                current['customer_pn'],
                current['company_pn']
            ]
            sres = sre_list[i:i+2]
            row.extend(sres + [''] * (2 - len(sres)))  # 保证2列SRE
            rows.append(row)
        sre_list = []  # 重置

# 生成 DataFrame
df = pd.DataFrame(rows, columns=["Sold-to BP", "Ship-to BP", "客户产品型号", "本公司产品型号", "SRE1", "SRE2"])

# 导出到Excel
output_file = "flattened_output.xlsx"
df.to_excel(output_file, index=False)
print(f"✅ Exported to {output_file}")
