from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


out = Path(__file__).with_name('셀리_시연용_거래내역.xlsx')
wb = Workbook()
ws = wb.active
ws.title = '거래내역'
ws.append(['거래일', '거래처', '계정명', '금액', '적요'])
rows = [
    ('2026-10-02', '한빛상사', '매출', 245000000, '10월 공급 계약'),
    ('2026-10-03', '새봄유통', '매출', 87000000, '정기 납품'),
    ('2026-10-04', '라온테크', '매출', 318000000, '장비 공급'),
    ('2026-10-05', '가온물류', '운송비', 42000000, '배송 대금'),
    ('2026-10-07', '다온전자', '매출', 176000000, '부품 납품'),
    ('2026-10-09', '푸른산업', '매출', 92000000, '정기 공급'),
    ('2026-10-10', '미래솔루션', '매출', 412000000, '시스템 구축'),
    ('2026-10-13', '온유상사', '매출', 125000000, '사무기기 공급'),
    ('2026-10-15', '해솔무역', '매출', 73000000, '수출 물품'),
    ('2026-10-18', '아람에너지', '매출', 284000000, '설비 공급'),
    ('2026-10-20', '별하유통', '매출', 51000000, '소모품 납품'),
    ('2026-10-22', '누리정보', '매출', 139000000, '서비스 계약'),
]
for row in rows:
    ws.append(row)
ws.freeze_panes = 'A2'
ws.sheet_view.showGridLines = False
for col, width in {'A': 18, 'B': 22, 'C': 16, 'D': 22, 'E': 28}.items():
    ws.column_dimensions[col].width = width
for cell in ws[1]:
    cell.fill = PatternFill('solid', fgColor='284B3D')
    cell.font = Font(name='맑은 고딕', bold=True, color='FFFFFF', size=12)
    cell.alignment = Alignment(vertical='center')
ws.row_dimensions[1].height = 30
for row in ws.iter_rows(min_row=2):
    for cell in row:
        cell.font = Font(name='맑은 고딕', size=11, color='24312C')
        cell.alignment = Alignment(vertical='center')
    row[3].number_format = '#,##0" 원"'
    ws.row_dimensions[row[0].row].height = 26
table = Table(displayName='DemoTransactions', ref=f'A1:E{ws.max_row}')
table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium4', showRowStripes=True)
ws.add_table(table)
wb.save(out)
print(out)
