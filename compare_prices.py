from pathlib import Path
from datetime import date
import pandas as pd
import re
from string import whitespace



ROOT = Path('./')
NULL = ["<NA>", "nan", "", None]
xlsx_pattern = r'.*(.xlsx)$'
UPC = ["UPC", "UPC CODE"]
SKU = ["SKU", "SKU CODE", "CODE"]
BRAND = ["BRAND"]
PROD = ["PRODUCT", "DESCRIPTION"]
SIZE = ["CASE PACK", "SIZE"]
PRICE = ["REG PRICE", "BASE PRICE"]
SALE = ["PROMO PRICE", "YOUR  COST", "SALE PRICE"]
HEADER_VALUES = list(set().union(UPC,SKU))
ALLNUM_VALUES = list(set().union(UPC,SKU,SIZE,PRICE,SALE))

keep_null = input("Keep rows without a UPC and/or [sale and base] price? (y/n): ")
if keep_null == "y":
    keep_null = True
elif keep_null == "n":
    keep_null = False
else:
    input("Inavalid input. Try again.")
    exit()

def remove_whitespace(code):
    for char in whitespace:
        code = code.replace(char, '_')
    return code

UPCU = []
SKUU = []
BRANDU = []
PRODU = []
SIZEU = []
PRICEU = []
SALEU = []
for code in UPC:
    UPCU.append(remove_whitespace(code))
for code in SKU:
    SKUU.append(remove_whitespace(code))
for code in BRAND:
    BRANDU.append(remove_whitespace(code))
for code in PROD:
    PRODU.append(remove_whitespace(code))
for code in SIZE:
    SIZEU.append(remove_whitespace(code))
for code in PRICE:
    PRICEU.append(remove_whitespace(code))
for code in SALE:
    SALEU.append(remove_whitespace(code))



class Row:
    def __init__(self, upc, sku, brand, prod, size, price, sale, filename, best, value):
        self.upc = upc
        self.sku = sku
        self.brand = brand
        self.prod = prod
        self.size = size
        self.price = price
        self.sale = sale
        self.filename = filename
        self.best = best
        self.value = value


def f7(seq):
    seen = set()
    seen_add = seen.add
    return [x for x in seq if not (x in seen or seen_add(x))]


def openw_header(file):
    df_clean_header = None
    df = pd.read_excel(file)
    header = df[df.isin(HEADER_VALUES).any(axis=1)]
    header_index = 0
    if len(header.index.values) != 0:
        header_index = (int(header.index.values) + 1)
    df = pd.read_excel(file, header=header_index, dtype="string")
    for char in whitespace:
        df.columns = df.columns.str.replace(char, '_')
    df_clean_header = df.dropna(axis=1, how='all')
    return df_clean_header



def get_columns(file):
    df = openw_header(file)
    upc_index = None
    sku_index = None
    brand_index = None
    prod_index = None
    size_index = None
    size_fail = []
    price_index = None
    sale_index = None
    for code in UPCU:
        try:
            upc_index = (int(df.columns.get_loc(code))+1)
        except:
            continue
        if upc_index != None:
            break
    for code in SKUU:
        try:
            sku_index = (int(df.columns.get_loc(code))+1)
        except:
            continue
        if sku_index != None:
            break
    for code in BRANDU:
        try:
            brand_index = (int(df.columns.get_loc(code))+1)
        except:
            continue
        if brand_index != None:
            break
    for code in PRODU:
        try:
            prod_index = (int(df.columns.get_loc(code))+1)
        except:
            continue
        if prod_index != None:
            break
    for code in SIZEU:
        try:
            size_index = (int(df.columns.get_loc(code))+1)
        except:
            size_fail.append(code)
        if size_index != None:
            break
    for code in PRICEU:
        try:
            price_index = (int(df.columns.get_loc(code))+1)
        except:
            continue
        if price_index != None:
            break
    for code in SALEU:
        try:
            sale_index = (int(df.columns.get_loc(code))+1)
        except:
            continue
        if sale_index != None:
            break

    rows = []
    for line in df.itertuples():
        upc = line[upc_index]
        if str(upc) not in NULL:
            for char in whitespace:
                upc = str(upc).replace(char, '')
            while len(upc) < 12:
                temp_upc = upc
                upc = ("0" + temp_upc)
        sku = line[sku_index]
        brand = line[brand_index]
        prod = line[prod_index]
        size = ""
        if size_fail == SIZEU:
            find_size = re.findall(r'[0-9]{1,3}\/[0-9]{1,4}[.x]?[0-9]{0,4}[a-zA-Z]+', prod)
            if len(find_size) > 0:
                size = find_size[0]
        else:
            size = line[size_index]
        price = line[price_index]
        if "$" in str(price):
            price = str(price).replace("$", '')
        sale = line[sale_index]
        if "$" in str(sale):
            sale = str(sale).replace("$", '')
        filename = file.name
        count = "1"
        if str(size) not in NULL:
            count = (re.findall(r'[0-9]{1,3}', size)[0])
        best = sale
        if str(sale) in NULL:
            best = price
        value = 0
        if str(best) not in NULL and str(count) not in NULL:
            value = float(best)/float(count)
        row = Row(upc, sku, brand, prod, size, price, sale, filename, best, value)
        rows.append(row)
    return rows



items = [item for item in ROOT.iterdir()]
files = []
for item in items:
    if item.is_file() and re.match(xlsx_pattern, str(item)):
        files.append(item)
    else:
        continue


duplicates = []
values = []
no_upc = []
no_price = []
for file in files:
    print(file)
    if len(values) == 0:
        values = get_columns(file) # for item in values
    else:
        values_upcs = []
        values_add = []
        for line in values:
            values_upcs.append(line.upc)
        rows = get_columns(file)
        for line1 in values:
            if str(line1.upc) in NULL:
                no_upc.append(line1)
            elif str(line1.best) in NULL:
                no_price.append(line1)
            else:
                for line2 in rows:
                    if str(line2.upc) in NULL:
                        no_upc.append(line2)
                    elif str(line2.best) in NULL:
                        no_price.append(line2)
                    else:
                        if line1.upc == line2.upc:
                            duplicates.append(str(line1.upc) + " - " + str(line1.best) + " - " + line1.filename + " | " + str(line2.upc) + " - " + str(line2.best) + " - " + line2.filename)
                            if line2.value < line1.value:
                                values[int(values.index(line1))] = rows[int(rows.index(line2))]
                            else:
                                continue
                        elif line2.upc in values_upcs: # <values_upcs> list of upcs in <values> before comparing to current file.
                            continue
                        else:
                            values_add.append(line2)
        values.extend(values_add)
        if keep_null == True:
            values.extend(no_upc)
            values.extend(no_price)
            no_upc.clear()
            no_price.clear()
values_dedupe = f7(values)


upc = []
sku = []
brand = []
prod = []
size = []
price = []
sale = []
filename = []
for item in values_dedupe:
    upc.append(item.upc)
    sku.append(item.sku)
    brand.append(item.brand)
    prod.append(item.prod)
    size.append(item.size)
    price.append('$' + item.price)
    sale.append('$' + item.sale)
    filename.append(item.filename)

data = []
data.append(upc)
data.append(sku)
data.append(brand)
data.append(prod)
data.append(size)
data.append(price)
data.append(sale)
data.append(filename)

today = str(date.today())

df = pd.DataFrame(data).transpose()
df.columns=['UPC', 'SKU', 'BRAND', 'PRODUCT', 'SIZE', 'BASE_PRICE', 'SALE_PRICE', 'FILENAME']
df.to_excel("best-sales_" + today + ".xlsx", sheet_name='sheet1', index=False)
with open('matching_UPCs.txt', 'w') as file_write:
    file_write.write('\n'.join(duplicates))