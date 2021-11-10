from src.div import user_build_only
from src.matcher import *
import json

with open('json\\hw_keyword.json') as f:
    keyword = json.load(f)

with open('json\\keywords.json') as k:
    older_kw = json.load(k)

bench_kw = keyword['bench']
bench_kw_split = keyword['bench_split']
valuecan_kw = keyword['valuecan']
pcan_kw = keyword['pcan']
user_kw = older_kw['user_build']

def bench(cell_data):
    for cell in cell_data:
        if matcher_slice(bench_kw, cell) or matcher_split(bench_kw_split, cell):
            return True
    return False

def valuecan(cell_data):
    for cell in cell_data:
        if matcher_slice(valuecan_kw, cell):
            return True
    return False

def pcan(cell_data, tcid):
    # User-build contain did
    if (matcher_slice(user_kw, cell_data[3]) or matcher_slice(user_kw, cell_data[0])) and 'did' in tcid.lower():
        return True

    # 'test screen' in expected result
    elif matcher_slice(['test screen'], cell_data[3]):
        return True

    # '14DA80F2' in test step
    elif matcher_slice(['14DA80F2'], cell_data[1]):
        return True
    else:
        return False
