from openpyxl import load_workbook

wb = load_workbook('sagatave_eksamenam.xlsx',data_only=True)
ws = wb['Lapa_0']
max_row = ws.max_row

sum = 0

for row in range(2, max_row + 1):
    client = ws['F' + str(row)].value
    qty = ws['L' + str(row)].value
    total = ws['N' + str(row)].value

    if client == 'Korporatīvais' and isinstance(qty, (int, float)) and 40 <= qty <= 50:
        sum += int(total)

sum = int(sum)
print(sum)