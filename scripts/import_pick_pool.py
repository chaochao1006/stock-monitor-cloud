"""Import the union of the two index sheets into a cloud-readable fixed pool."""
from __future__ import annotations

import argparse
import hashlib
import json
from datetime import datetime
from pathlib import Path
import re
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

DEFAULT_OUTPUT = Path(__file__).with_name('pick_stock_pool.json')


def import_pool(source, output=DEFAULT_OUTPUT):
    source = Path(source)
    workbook = load_workbook(source, read_only=True, data_only=True)
    stocks, sheets = {}, {}
    try:
        for sheet in workbook:
            header = None
            sheet_symbols = set()
            for row in sheet.iter_rows(values_only=True):
                if header is None:
                    if '股票代码' in row:
                        header = {value: i for i, value in enumerate(row) if isinstance(value, str)}
                    continue
                value = row[header['股票代码']]
                if value is None or not str(value).strip():
                    continue
                display = str(value).strip().upper()
                if not re.fullmatch(r'[A-Z][A-Z0-9]*(?:[.-][A-Z0-9]+)?', display):
                    raise ValueError(f'{sheet.title}: 非法股票代码 {display!r}')
                ticker = display.replace('.', '-')
                def field(name):
                    return str(row[header[name]] or '').strip() if name in header else ''
                item = stocks.setdefault(ticker, {
                    'ticker': ticker, 'display_ticker': display,
                    'company_name': field('公司名称（英文）'),
                    'sector': field('一级板块'), 'industry': field('细分行业'), 'indexes': [],
                })
                if sheet.title not in item['indexes']:
                    item['indexes'].append(sheet.title)
                sheet_symbols.add(ticker)
            if header is None:
                raise ValueError(f'{sheet.title}: 没找到股票代码列')
            sheets[sheet.title] = len(sheet_symbols)
    finally:
        workbook.close()
    if not stocks:
        raise ValueError('表格没有股票')
    payload = {
        'schema_version': 1, 'source_file': source.name,
        'source_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
        'imported_at_cn': datetime.now(ZoneInfo('Asia/Shanghai')).isoformat(timespec='seconds'),
        'sheet_counts': sheets, 'stock_pool_count': len(stocks), 'stocks': list(stocks.values()),
    }
    output = Path(output)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    return payload


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('--output', type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    pool = import_pool(args.source, args.output)
    print(json.dumps({'sheets': pool['sheet_counts'], 'unique': pool['stock_pool_count']}, ensure_ascii=True))
