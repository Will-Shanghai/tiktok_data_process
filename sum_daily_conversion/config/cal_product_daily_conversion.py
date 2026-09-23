#!/usr/bin/env python3
import argparse
import csv
import json
import os
import re
import sys
import time
from http.client import IncompleteRead, RemoteDisconnected
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from urllib import request, error
from xml.etree import ElementTree as ET
from zipfile import ZipFile

PROJECT_ROOT = Path(os.getenv("TIKTOK_CONVERSION_ROOT", Path(__file__).resolve().parents[1])).resolve()


STANDARD_COLUMNS = [
    ("", "日期"),
    ("", "商品ID"),
    ("", "发品状态"),
    ("全部", "GMV"),
    ("全部", "订单数"),
    ("全部", "商品成交件数"),
    ("全部", "平均订单金额（SKU 订单）"),
    ("商家视频", "归因订单数"),
    ("商家视频", "新视频数"),
    ("联盟", "归因订单数"),
    ("联盟", "新视频数"),
    ("联盟", "ATC 用户数（视频）"),
    ("联盟", "ATC 用户数（直播）"),
    ("商家商品卡", "归因订单数"),
    ("全部", "商城页商品成交件数"),
    ("全部", "商品曝光次数"),
    ("全部", "商品点击量"),
    ("全部", "商品点击率"),
    ("全部", "CTOR（SKU 订单）"),
    ("全部", "加购次数"),
    ("全部", "加购率"),
    ("商家视频", "商品曝光次数"),
    ("商家视频", "商品点击量"),
    ("商家视频", "商品点击率"),
    ("商家视频", "CTOR（SKU 订单）"),
    ("商家视频", "加购次数"),
    ("商家视频", "加购率"),
    ("联盟", "商品曝光次数"),
    ("联盟", "商品点击量"),
    ("联盟", "商品点击率"),
    ("联盟", "CTOR（SKU 订单）"),
    ("联盟", "加购次数（视频）"),
    ("联盟", "去重加购率（视频）"),
    ("商家商品卡", "去重商品曝光次数"),
    ("商家商品卡", "去重点击次数"),
    ("商家商品卡", "去重点击率"),
    ("商家商品卡", "去重点击成交转化率（SKU 订单）"),
    ("商家商品卡", "已加购的用户数"),
    ("商家商品卡", "去重加购率"),
    ("全部", "退款金额"),
    ("全部", "已退款的商品件数"),
]

STANDARD_GROUPS = [group for group, _ in STANDARD_COLUMNS]
STANDARD_FIELDS = [field for _, field in STANDARD_COLUMNS]

DISPLAY_GROUP_NAMES = {
    "商家商品卡": "商品卡",
}

DISPLAY_FIELD_NAMES = {
    ("全部", "平均订单金额（SKU 订单）"): "平均订单金额",
    ("联盟", "ATC 用户数（视频）"): "下单用户数（视频）",
    ("联盟", "ATC 用户数（直播）"): "下单用户数（直播）",
    ("全部", "商城页商品成交件数"): "商城页成交件数",
    ("全部", "CTOR（SKU 订单）"): "商品转化率",
    ("商家视频", "CTOR（SKU 订单）"): "商品转化率",
    ("联盟", "CTOR（SKU 订单）"): "商品转化率",
    ("商家商品卡", "去重点击成交转化率（SKU 订单）"): "去重转化率",
}

DISPLAY_GROUPS = [DISPLAY_GROUP_NAMES.get(group, group) for group, _ in STANDARD_COLUMNS]
DISPLAY_FIELDS = [DISPLAY_FIELD_NAMES.get((group, field), field) for group, field in STANDARD_COLUMNS]

SOURCE_FIELD_NAMES = {}

HEADER_SECTION_COLORS = {
    "": ("FFD9E0E8", "FFF7F9FB"),
    "全部": ("FFBDD7FF", "FFEAF3FF"),
    "商家视频": ("FFCDEFD8", "FFF0FAF3"),
    "联盟": ("FFE1D9FF", "FFF6F2FF"),
    "商品卡": ("FFFFE1C7", "FFFFF4E8"),
}

VIDEO_SUMMARY_HEADER_GROUPS = ["", "", "", "商家视频", "联盟"]
VIDEO_SUMMARY_HEADER_FIELDS = ["日期", "商品ID", "产品名称", "商家短视频", "联盟短视频"]

DEFAULT_STORE_ORDER = ["日本本土店", "日本直邮一店", "日本直邮二店", "日本跨境店"]
DEFAULT_FLAT_STORE_NAME = "商品每日转化统计"
NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


@dataclass
class ProductRow:
    store: str
    product_id: str
    product_name: str
    sheet_name: str
    values: List[str]
    source_file: Path


@dataclass
class SkippedProduct:
    store: str
    source_file: Path
    product_id: str
    product_name: str
    reason: str


def load_config(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"找不到配置文件：{path}。请在项目目录下创建 config.json。")
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_project_path(path_like: str | Path) -> Path:
    path = Path(path_like)
    return path if path.is_absolute() else PROJECT_ROOT / path


def load_env(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def col_to_num(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch.upper()) - 64
    return n


def col_to_letter(index: int) -> str:
    n = index
    letters = ""
    while n:
        n, remainder = divmod(n - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def read_shared_strings(zf: ZipFile) -> List[str]:
    try:
        raw = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(raw)
    strings = []
    for si in root.findall("x:si", NS):
        parts = [node.text or "" for node in si.findall(".//x:t", NS)]
        strings.append("".join(parts))
    return strings


def cell_text(cell: ET.Element, shared_strings: List[str]) -> str:
    cell_type = cell.attrib.get("t")
    value = cell.find("x:v", NS)
    inline = cell.find("x:is", NS)

    if cell_type == "s" and value is not None:
        idx = int(value.text or "0")
        return shared_strings[idx] if idx < len(shared_strings) else ""
    if inline is not None:
        return "".join(node.text or "" for node in inline.findall(".//x:t", NS))
    if value is not None:
        return value.text or ""
    return ""


def read_xlsx_rows(path: Path) -> List[List[str]]:
    with ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

    rows = []
    for row in root.findall(".//x:sheetData/x:row", NS):
        values: Dict[int, str] = {}
        for cell in row.findall("x:c", NS):
            ref = cell.attrib.get("r", "")
            col = col_to_num(ref)
            values[col] = cell_text(cell, shared_strings)
        if values:
            rows.append([values.get(i, "") for i in range(1, max(values) + 1)])
    return rows


def format_date_value(value: date, output_format: str) -> str:
    if output_format == "mdd":
        return f"{value.month}{value.day:02d}"
    if output_format == "mmdd":
        return f"{value.month:02d}{value.day:02d}"
    return value.isoformat()


def parse_date_from_filename(path: Path, default_year: int, output_format: str) -> str:
    match = re.search(r"(?<!\d)(\d{2})(\d{2})(?!\d)", path.stem)
    if not match:
        raise ValueError(f"文件名无法识别日期：{path.name}。请使用 0712.xlsx 或 日本本土店-0712.xlsx 这种格式。")
    month, day = int(match.group(1)), int(match.group(2))
    return format_date_value(date(default_year, month, day), output_format)


def parse_date_from_workbook(rows: List[List[str]], default_year: int, output_format: str) -> Optional[str]:
    for row in rows[:5]:
        for cell in row[:5]:
            match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", cell or "")
            if match:
                first, second, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
                # TikTok Japan exports usually display day/month/year, e.g. 12/07/2026.
                day, month = first, second
                if first > 12:
                    day, month = first, second
                elif second > 12:
                    month, day = first, second
                return format_date_value(date(year or default_year, month, day), output_format)
    return None


def find_header_and_data(rows: List[List[str]]) -> Tuple[List[str], List[str], List[List[str]]]:
    header: Optional[List[str]] = None
    group_header: List[str] = []
    data_rows: List[List[str]] = []
    for row_index, row in enumerate(rows):
        if "商品 ID" in row:
            header = row
            group_header = rows[row_index - 1] if row_index > 0 else []
            continue
        if header:
            indexes = first_header_index(header)
            product_id_index = indexes.get("商品 ID")
            if product_id_index is not None and product_id_index < len(row) and row[product_id_index].strip():
                data_rows.append(row)
    if not header:
        raise ValueError("没有找到表头行：缺少“商品 ID”列。")
    return group_header, header, data_rows


def first_header_index(header: List[str]) -> Dict[str, int]:
    indexes: Dict[str, int] = {}
    for i, name in enumerate(header):
        if name and name not in indexes:
            indexes[name] = i
    return indexes


def get_row_value(row: List[str], indexes: Dict[str, int], field: str, file_date: str) -> str:
    if field == "日期":
        return file_date
    if field == "商品ID":
        return normalize_product_id(row[indexes["商品 ID"]])
    source_field = SOURCE_FIELD_NAMES.get(field, field)
    index = indexes.get(source_field)
    if index is None or index >= len(row):
        return ""
    return row[index]


def grouped_header_value(group_header: List[str], index: int) -> str:
    if index < len(group_header) and group_header[index]:
        return group_header[index]
    for i in range(min(index, len(group_header) - 1), -1, -1):
        if group_header[i]:
            return group_header[i]
    return ""


def grouped_header_index(group_header: List[str], header: List[str], group_name: str, source_field: str) -> Optional[int]:
    for index, name in enumerate(header):
        if name != source_field:
            continue
        group_value = grouped_header_value(group_header, index)
        if group_value == group_name:
            return index
    return None


def get_standard_value(
    row: List[str],
    group_header: List[str],
    header: List[str],
    indexes: Dict[str, int],
    group: str,
    field: str,
    file_date: str,
) -> str:
    if group and group != "全部":
        index = grouped_header_index(group_header, header, group, field)
        if index is None or index >= len(row):
            return ""
        return row[index]
    return get_row_value(row, indexes, field, file_date)


def get_required_cell(row: List[str], indexes: Dict[str, int], field: str) -> str:
    index = indexes.get(field)
    if index is None:
        raise ValueError(f"原始表缺少必要字段：{field}")
    if index >= len(row):
        return ""
    return row[index].strip()


@dataclass
class MappingEntry:
    store: str
    product_id: str
    product_name: str
    sheet_name: str
    match_count: int = 0
    match_modes: set = field(default_factory=set)


class ProductMapping:
    """商品映射索引。

    映射表的主键是 (店铺, 商品ID)。但 sum_daily_conversion/data 允许两种摆放方式：

    1. 分店铺子目录：data/日本本土店/xxx.xlsx —— 此时 store 是真实店铺，走精确匹配。
    2. 扁平目录：data/xxx.xlsx —— 此时无法从路径得到店铺，store 会退化成占位名。
       如果仍然要求 (店铺, 商品ID) 精确匹配，扁平目录下的商品会全部判定为“未配置映射”，
       输出 0 条（这就是之前 743 条被跳过的原因）。

    因此这里额外维护“无店铺”索引：当源数据来自扁平目录（store 不是映射表里的任何店铺）时，
    退化为按 商品ID / 商品名 匹配。为避免写错 sheet，只有在候选 sheet 唯一时才回退成功；
    同一商品ID 对应多个不同 sheet 时按“映射歧义”跳过，并输出候选明细。
    """

    def __init__(self) -> None:
        self.entries: List[MappingEntry] = []
        self.stores: set = set()
        self._by_store_id: Dict[Tuple[str, str], MappingEntry] = {}
        self._by_id: Dict[str, List[MappingEntry]] = {}
        self._by_name: Dict[str, List[MappingEntry]] = {}

    def add(self, store: str, product_id: str, product_name: str, sheet_name: str) -> None:
        entry = MappingEntry(store, product_id, product_name, sheet_name)
        self.entries.append(entry)
        self.stores.add(store)
        if product_id:
            self._by_store_id[(store, product_id)] = entry
            self._by_id.setdefault(product_id, []).append(entry)
        if product_name:
            self._by_name.setdefault(product_name, []).append(entry)

    def is_known_store(self, store: str) -> bool:
        return store in self.stores

    def ambiguous_product_ids(self) -> Dict[str, List[MappingEntry]]:
        """商品ID 在映射表里对应多个不同 sheet 的情况。"""
        return {
            product_id: entries
            for product_id, entries in self._by_id.items()
            if len({entry.sheet_name for entry in entries}) > 1
        }

    @staticmethod
    def _count(entries: Iterable[MappingEntry], mode: str) -> None:
        for entry in entries:
            entry.match_count += 1
            entry.match_modes.add(mode)

    def resolve(self, store: str, product_id: str, product_name: str) -> Tuple[Optional[str], str]:
        """返回 (飞书sheet名, 跳过原因)。匹配成功时原因为空字符串。"""
        if product_id:
            entry = self._by_store_id.get((store, product_id))
            if entry is not None:
                self._count([entry], "精确(店铺+商品ID)")
                return entry.sheet_name, ""

        # 店铺是映射表里已知的真实店铺，但该商品没配置 → 就是真的没配，不做跨店铺回退。
        if self.is_known_store(store):
            return None, "未配置映射"

        # 扁平目录 / 未知店铺：退化为无店铺匹配。
        for key, table, label in (
            (product_id, self._by_id, "商品ID"),
            (product_name, self._by_name, "商品名"),
        ):
            if not key:
                continue
            candidates = table.get(key, [])
            if not candidates:
                continue
            distinct = sorted({item.sheet_name for item in candidates})
            if len(distinct) == 1:
                self._count(candidates, f"{label}回退")
                return distinct[0], ""
            detail = ", ".join(f"{item.store}={item.sheet_name}" for item in candidates)
            return None, f"映射歧义：{label} 对应多个不同 sheet（{detail}），请拆分数据目录或调整映射表"
        return None, "未配置映射"


def load_mapping(path: Path) -> ProductMapping:
    mapping = ProductMapping()
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        required = {"店铺", "商品ID", "飞书sheet名"}
        if not required.issubset(reader.fieldnames or []):
            raise ValueError("映射表必须包含列：店铺, 商品ID, 飞书sheet名")
        for row in reader:
            store = (row.get("店铺") or "").strip()
            product_id = normalize_product_id(row.get("商品ID") or "")
            product_name = (row.get("商品名") or "").strip()
            sheet_name = (row.get("飞书sheet名") or "").strip()
            if not store or not sheet_name:
                continue
            if not product_id and not product_name:
                continue
            if product_id.lower() == "xxx":
                product_id = ""
            if not product_id and not product_name:
                continue
            mapping.add(store, product_id, product_name, sheet_name)
    return mapping


def normalize_product_id(value: str) -> str:
    return str(value).strip().lstrip("'").replace("\u200b", "")


def parse_workbook(path: Path, store: str, file_date: str, mapping: ProductMapping) -> Tuple[List[ProductRow], List[str], List[SkippedProduct]]:
    rows = read_xlsx_rows(path)
    return parse_workbook_from_rows(path, store, file_date, mapping, rows)


def parse_workbook_from_rows(
    path: Path,
    store: str,
    file_date: str,
    mapping: ProductMapping,
    rows: List[List[str]],
) -> Tuple[List[ProductRow], List[str], List[SkippedProduct]]:
    group_header, header, data_rows = find_header_and_data(rows)
    indexes = first_header_index(header)
    warnings: List[str] = []
    skipped: List[SkippedProduct] = []
    parsed: List[ProductRow] = []

    missing_fields = []
    for group, field in STANDARD_COLUMNS:
        if field in {"日期", "商品ID"}:
            continue
        if group and group != "全部":
            if grouped_header_index(group_header, header, group, field) is None:
                missing_fields.append(f"{group}-{field}")
            continue
        if field not in indexes:
            missing_fields.append(field)
    if missing_fields:
        warnings.append(f"{path.name} 缺少字段：{', '.join(missing_fields)}")

    for row in data_rows:
        product_name = get_required_cell(row, indexes, "商品名")
        product_id = normalize_product_id(get_required_cell(row, indexes, "商品 ID"))
        sheet_name, skip_reason = mapping.resolve(store, product_id, product_name)
        if not sheet_name:
            skipped.append(SkippedProduct(store, path, product_id, product_name, skip_reason))
            continue
        values = [
            get_standard_value(row, group_header, header, indexes, group, field, file_date)
            for group, field in STANDARD_COLUMNS
        ]
        parsed.append(ProductRow(store, product_id, product_name, sheet_name, values, path))
    return parsed, warnings, skipped


class FeishuClient:
    def __init__(self, app_id: str, app_secret: str, spreadsheet_token: str):
        self.app_id = app_id
        self.app_secret = app_secret
        self.spreadsheet_token = spreadsheet_token
        self.tenant_access_token: Optional[str] = None

    def request_json(self, method: str, url: str, payload: Optional[dict] = None, action: str = "飞书接口") -> dict:
        headers = {"Content-Type": "application/json; charset=utf-8"}
        if self.tenant_access_token:
            headers["Authorization"] = f"Bearer {self.tenant_access_token}"
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
        for attempt in range(1, 6):
            req = request.Request(url, data=data, method=method, headers=headers)
            try:
                with request.urlopen(req, timeout=60) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except (IncompleteRead, RemoteDisconnected, ConnectionResetError, TimeoutError, error.URLError) as exc:
                if attempt == 5:
                    raise RuntimeError(f"{action}网络连接失败，已重试 5 次仍失败：{exc}；接口：{method} {url}") from exc
                time.sleep(attempt * 2)
            except error.HTTPError as exc:
                body = exc.read().decode("utf-8", errors="replace")
                hint = ""
                if '"code":99991672' in body or "No permission" in body:
                    hint = "；这通常是飞书应用没有该表格权限，或缺少云文档/电子表格 API 权限"
                if '"code":91403' in body or "Forbidden" in body:
                    hint = "；这通常是应用没有当前接口权限、应用版本未发布生效，或该表格没有授权给应用/机器人"
                raise RuntimeError(f"{action}失败 HTTP {exc.code}{hint}: {body}；接口：{method} {url}") from exc
        raise RuntimeError("飞书接口请求失败：未知错误")

    def authenticate(self) -> None:
        url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
        res = self.request_json("POST", url, {"app_id": self.app_id, "app_secret": self.app_secret}, "获取 tenant_access_token")
        if res.get("code") != 0:
            raise RuntimeError(f"获取 tenant_access_token 失败：{res}")
        self.tenant_access_token = res["tenant_access_token"]

    def get_sheet_id_by_name(self) -> Dict[str, str]:
        url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{self.spreadsheet_token}/sheets/query"
        res = self.request_json("GET", url, action="获取飞书 sheet 列表")
        if res.get("code") != 0:
            raise RuntimeError(f"获取飞书 sheet 列表失败：{res}")
        sheets = res.get("data", {}).get("sheets", [])
        return {sheet["title"]: sheet["sheet_id"] for sheet in sheets}

    def create_sheet(self, title: str) -> str:
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/sheets_batch_update"
        payload = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": title,
                            "index": 0,
                            "rowCount": 2000,
                            "columnCount": len(STANDARD_FIELDS),
                        }
                    }
                }
            ]
        }
        res = self.request_json("POST", url, payload, "创建飞书 sheet")
        if res.get("code") != 0:
            raise RuntimeError(f"创建飞书 sheet 失败：{res}")
        replies = res.get("data", {}).get("replies", [])
        for reply in replies:
            properties = reply.get("addSheet", {}).get("properties", {})
            if properties.get("title") == title and properties.get("sheetId"):
                return properties["sheetId"]
        sheet_map = self.get_sheet_id_by_name()
        if title in sheet_map:
            return sheet_map[title]
        raise RuntimeError(f"创建 sheet 后未能获取 sheet_id：{title}")

    def read_values(self, sheet_id: str, range_a1: str) -> List[List[str]]:
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/values/{sheet_id}!{range_a1}"
        res = self.request_json("GET", url, action=f"读取飞书数据 {range_a1}")
        if res.get("code") != 0:
            raise RuntimeError(f"读取飞书数据失败：{res}")
        return res.get("data", {}).get("valueRange", {}).get("values", [])

    def update_values(self, sheet_id: str, range_a1: str, values: List[List[str]]) -> None:
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/values"
        payload = {"valueRange": {"range": f"{sheet_id}!{range_a1}", "values": values}}
        res = self.request_json("PUT", url, payload, f"更新飞书数据 {range_a1}")
        if res.get("code") != 0:
            raise RuntimeError(f"更新飞书数据失败：{res}")

    def append_values(self, sheet_id: str, values: List[List[str]]) -> None:
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/values_append"
        last_col = col_to_letter(len(STANDARD_FIELDS))
        payload = {"valueRange": {"range": f"{sheet_id}!A1:{last_col}1", "values": values}}
        res = self.request_json("POST", url, payload, "追加飞书数据")
        if res.get("code") != 0:
            raise RuntimeError(f"追加飞书数据失败：{res}")

    def find_date_row(self, sheet_id: str, target_date: str, start_row: int = 3, end_row: int = 1000, chunk_size: int = 100) -> Optional[int]:
        for chunk_start in range(start_row, end_row + 1, chunk_size):
            chunk_end = min(chunk_start + chunk_size - 1, end_row)
            date_rows = self.read_values(sheet_id, f"A{chunk_start}:A{chunk_end}")
            for offset, row in enumerate(date_rows):
                if row and str(row[0]).strip().zfill(4) == target_date:
                    return chunk_start + offset
        return None

    def upsert_daily_row(self, sheet_id: str, values: List[str]) -> str:
        last_col = col_to_letter(len(STANDARD_FIELDS))
        header_rows = self.read_values(sheet_id, f"A1:{last_col}2")
        if not header_rows:
            self.update_values(sheet_id, f"A1:{last_col}2", [DISPLAY_GROUPS, DISPLAY_FIELDS])
            self.append_values(sheet_id, [values])
            return "created_header_and_appended"

        group_header = header_rows[0] if len(header_rows) >= 1 else []
        field_header = header_rows[1] if len(header_rows) >= 2 else []
        if group_header[: len(DISPLAY_GROUPS)] != DISPLAY_GROUPS or field_header[: len(DISPLAY_FIELDS)] != DISPLAY_FIELDS:
            self.update_values(sheet_id, f"A1:{last_col}2", [DISPLAY_GROUPS, DISPLAY_FIELDS])

        target_date = values[0]
        row_index = self.find_date_row(sheet_id, target_date)
        if row_index is not None:
            self.update_values(sheet_id, f"A{row_index}:{last_col}{row_index}", [values])
            return "updated"

        self.append_values(sheet_id, [values])
        return "appended"


def spreadsheet_token_for_store(store: str) -> str:
    env_names = {
        "日本本土店": "FEISHU_LOCAL_SPREADSHEET_TOKEN",
        "日本跨境店": "FEISHU_CROSS_BORDER_SPREADSHEET_TOKEN",
        "日本直邮一店": "FEISHU_DIRECT_SPREADSHEET_TOKEN",
        "日本直邮二店": "FEISHU_DIRECT2_SPREADSHEET_TOKEN",
    }
    env_name = env_names.get(store)
    if not env_name:
        raise ValueError(f"未知店铺：{store}")
    token = os.getenv(env_name, "").strip()
    if not token:
        raise ValueError(f"{store} 缺少飞书表格 token，请在 .env 填写 {env_name}。")
    if not is_valid_spreadsheet_token(token):
        raise ValueError(f"{store} 的 {env_name} 不是有效飞书表格 token，请把占位文本替换成真实 token。")
    return token


def is_valid_spreadsheet_token(token: str) -> bool:
    if not token or any(ord(ch) > 127 for ch in token):
        return False
    lowered = token.lower()
    if any(word in lowered for word in ["token", "spreadsheet", "placeholder"]):
        return False
    return bool(re.fullmatch(r"[A-Za-z0-9_-]{10,}", token))


def discover_store_dirs(data_dir: Path) -> List[str]:
    if not data_dir.exists():
        return []
    stores = [path.name for path in data_dir.iterdir() if path.is_dir() and not path.name.startswith(".")]
    ordered = [store for store in DEFAULT_STORE_ORDER if store in stores]
    extras = sorted(store for store in stores if store not in DEFAULT_STORE_ORDER)
    return ordered + extras


def iter_input_files(data_dir: Path, store_dirs: List[str]) -> Iterable[Tuple[str, Path]]:
    if not store_dirs:
        for path in sorted(data_dir.glob("*.xlsx")):
            if not path.name.startswith(("~$", ".~")):
                yield DEFAULT_FLAT_STORE_NAME, path
        return

    for store in store_dirs:
        store_dir = data_dir / store
        if not store_dir.exists():
            continue
        for path in sorted(store_dir.glob("*.xlsx")):
            if not path.name.startswith(("~$", ".~")):
                yield store, path


def resolve_data_dir(configured_dir: Path) -> Path:
    candidates = [
        configured_dir,
        configured_dir / "daily_convarsion_data",
        configured_dir / "daily_conversion_data",
    ]
    for candidate in candidates:
        if not candidate.exists():
            continue
        if discover_store_dirs(candidate):
            return candidate
    return configured_dir


def write_preview(rows: List[ProductRow], log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    path = log_dir / "dry_run_preview.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["", "", "", "", *DISPLAY_GROUPS])
        writer.writerow(["店铺", "商品ID", "商品名", "飞书sheet名", *DISPLAY_FIELDS])
        for row in rows:
            writer.writerow([row.store, row.product_id, row.product_name, row.sheet_name, *row.values])
    return path


def write_skipped(skipped: List[SkippedProduct], log_dir: Path) -> Optional[Path]:
    if not skipped:
        return None
    log_dir.mkdir(parents=True, exist_ok=True)
    path = log_dir / "skipped_unmapped_products.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["店铺", "文件", "商品ID", "商品名", "原因"])
        for item in skipped:
            writer.writerow([item.store, item.source_file.name, item.product_id, item.product_name, item.reason])
    return path


def write_mapping_diagnostics(mapping: ProductMapping, log_dir: Path) -> Optional[Path]:
    """输出映射表体检表，方便运营补齐/修正 product_sheet_mapping.csv。"""
    if not mapping.entries:
        return None
    log_dir.mkdir(parents=True, exist_ok=True)
    path = log_dir / "mapping_diagnostics.csv"
    ambiguous = set(mapping.ambiguous_product_ids())
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["店铺", "商品ID", "商品名", "飞书sheet名", "状态", "命中次数", "命中方式"])
        for entry in mapping.entries:
            if entry.match_count:
                status = "已匹配"
            elif entry.product_id and entry.product_id in ambiguous:
                status = "歧义（同ID多sheet）"
            else:
                status = "未出现在源数据"
            writer.writerow(
                [
                    entry.store,
                    entry.product_id,
                    entry.product_name,
                    entry.sheet_name,
                    status,
                    entry.match_count,
                    "/".join(sorted(entry.match_modes)),
                ]
            )
    return path


def write_store_workbook(store: str, rows: List[ProductRow], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    column_index = {(group, field): idx for idx, (group, field) in enumerate(STANDARD_COLUMNS)}
    merchant_new_video_idx = column_index[("商家视频", "新视频数")]
    alliance_new_video_idx = column_index[("联盟", "新视频数")]

    def escape_xml(value: object) -> str:
        text = "" if value is None else str(value)
        return (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    def safe_sheet_name(value: str) -> str:
        name = re.sub(r"[\[\]\:\*\?\/\\]", "_", value).strip() or "未命名"
        return name[:31]

    def build_styles_xml() -> Tuple[str, Dict[Tuple[str, str], int]]:
        fills = [
            '<fill><patternFill patternType="none"/></fill>',
            '<fill><patternFill patternType="gray125"/></fill>',
        ]
        xfs = [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>',
        ]
        style_ids: Dict[Tuple[str, str], int] = {}

        def add_fill(color: str) -> int:
            fill_id = len(fills)
            fills.append(
                '<fill><patternFill patternType="solid">'
                f'<fgColor rgb="{color}"/>'
                f'<bgColor rgb="{color}"/>'
                '</patternFill></fill>'
            )
            return fill_id

        def add_header_style(fill_id: int) -> int:
            style_id = len(xfs)
            xfs.append(
                '<xf numFmtId="0" fontId="1" fillId="{fill_id}" borderId="1" xfId="0" '
                'applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
                '<alignment horizontal="center" vertical="center" wrapText="1"/>'
                '</xf>'.format(fill_id=fill_id)
            )
            return style_id

        for group_name in ["", "全部", "商家视频", "联盟", "商品卡"]:
            group_fill, field_fill = HEADER_SECTION_COLORS[group_name]
            style_ids[(group_name, "group")] = add_header_style(add_fill(group_fill))
            style_ids[(group_name, "field")] = add_header_style(add_fill(field_fill))

        styles_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="Microsoft YaHei"/><family val="2"/></font><font><b/><sz val="11"/><name val="Microsoft YaHei"/><family val="2"/></font></fonts>
  <fills count="{len(fills)}">{"".join(fills)}</fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FF000000"/></left>
      <right style="thin"><color rgb="FF000000"/></right>
      <top style="thin"><color rgb="FF000000"/></top>
      <bottom style="thin"><color rgb="FF000000"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''
        return styles_xml, style_ids

    styles_xml, style_ids = build_styles_xml()

    def sheet_xml(
        matrix: List[List[str]],
        header_groups: Optional[List[str]] = None,
        header_fields: Optional[List[str]] = None,
        header_row_count: int = 2,
    ) -> str:
        display_matrix = [row[:] for row in matrix]
        merge_ranges = []
        if (
            header_row_count == 2
            and header_groups
            and header_fields
            and len(display_matrix) >= 2
            and display_matrix[0] == header_groups
        ):
            for col_index, group in enumerate(header_groups, start=1):
                if not group:
                    merge_ranges.append(f"{col_to_letter(col_index)}1:{col_to_letter(col_index)}2")
                    display_matrix[0][col_index - 1] = header_fields[col_index - 1]
                    display_matrix[1][col_index - 1] = ""
            start = 1
            while start <= len(header_groups):
                group = header_groups[start - 1]
                end = start
                while end + 1 <= len(header_groups) and header_groups[end] == group:
                    end += 1
                if group and end > start:
                    merge_ranges.append(f"{col_to_letter(start)}1:{col_to_letter(end)}1")
                    for col_index in range(start + 1, end + 1):
                        display_matrix[0][col_index - 1] = ""
                start = end + 1
        col_widths = []
        if display_matrix:
            for col_index in range(1, len(display_matrix[0]) + 1):
                max_len = 0
                for row_values in display_matrix:
                    if col_index <= len(row_values):
                        max_len = max(max_len, len(str(row_values[col_index - 1] or "")))
                width = min(max(max_len + 3, 10), 30)
                col_widths.append(f'<col min="{col_index}" max="{col_index}" width="{width}" customWidth="1"/>')
        cols_xml = f"<cols>{''.join(col_widths)}</cols>" if col_widths else ""
        rows_xml = []
        for row_index, row_values in enumerate(display_matrix, start=1):
            cells = []
            for col_index, value in enumerate(row_values, start=1):
                ref = f"{col_to_letter(col_index)}{row_index}"
                if row_index == 1:
                    group_name = header_groups[col_index - 1] if header_groups and col_index - 1 < len(header_groups) else ""
                    style_key = "field" if header_row_count == 1 else "group"
                    style_id = style_ids.get((group_name, style_key), style_ids[("", style_key)])
                elif row_index == 2 and header_row_count >= 2:
                    group_name = header_groups[col_index - 1] if header_groups and col_index - 1 < len(header_groups) else ""
                    style_id = style_ids.get((group_name, "field"), style_ids[("", "field")])
                else:
                    style_id = 0
                style = f' s="{style_id}"'
                cells.append(f'<c r="{ref}" t="inlineStr"{style}><is><t>{escape_xml(value)}</t></is></c>')
            rows_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')
        merge_xml = ""
        if merge_ranges:
            merge_xml = f'<mergeCells count="{len(merge_ranges)}">' + "".join(
                f'<mergeCell ref="{merge_range}"/>' for merge_range in merge_ranges
            ) + "</mergeCells>"
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="{header_row_count}" topLeftCell="A{header_row_count + 1}" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft"/>
    </sheetView>
  </sheetViews>
  {cols_xml}
  <sheetData>{"".join(rows_xml)}</sheetData>
  {merge_xml}
</worksheet>'''

    rows_by_product: Dict[str, List[ProductRow]] = {}
    for row in rows:
        rows_by_product.setdefault(row.sheet_name, []).append(row)
    sheets = []
    summary_matrix = [VIDEO_SUMMARY_HEADER_FIELDS]
    for row in sorted(rows, key=lambda item: (item.values[0], item.sheet_name, item.product_id)):
        summary_matrix.append(
            [
                row.values[0],
                row.product_id,
                row.sheet_name,
                row.values[merchant_new_video_idx],
                row.values[alliance_new_video_idx],
            ]
        )
    if len(summary_matrix) > 2:
        sheets.append(("新视频数汇总", summary_matrix, VIDEO_SUMMARY_HEADER_GROUPS, VIDEO_SUMMARY_HEADER_FIELDS, 1))
    for sheet_name, product_rows in rows_by_product.items():
        matrix = [DISPLAY_GROUPS, DISPLAY_FIELDS]
        for row in sorted(product_rows, key=lambda item: item.values[0]):
            matrix.append(row.values)
        sheets.append((safe_sheet_name(sheet_name), matrix, DISPLAY_GROUPS, DISPLAY_FIELDS, 2))
    if not sheets:
        sheets.append(("无数据", [DISPLAY_GROUPS, DISPLAY_FIELDS], DISPLAY_GROUPS, DISPLAY_FIELDS, 2))

    workbook_sheets = []
    workbook_rels = []
    with ZipFile(output_path, "w") as zf:
        zf.writestr("[Content_Types].xml", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
''' + "".join(
            f'  <Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
            for i in range(1, len(sheets) + 1)
        ) + "</Types>")
        zf.writestr("_rels/.rels", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''')
        zf.writestr("xl/styles.xml", styles_xml)

        for i, (sheet_name, matrix, header_groups, header_fields, header_row_count) in enumerate(sheets, start=1):
            workbook_sheets.append(f'<sheet name="{escape_xml(sheet_name)}" sheetId="{i}" r:id="rId{i}"/>')
            workbook_rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
            zf.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(matrix, header_groups, header_fields, header_row_count))

        workbook_rels.append(f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
        zf.writestr("xl/workbook.xml", f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>{"".join(workbook_sheets)}</sheets>
</workbook>''')
        zf.writestr("xl/_rels/workbook.xml.rels", f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{"".join(workbook_rels)}</Relationships>''')
    return output_path


def write_local_excels_by_store(rows: List[ProductRow], output_dir: Path, store_dirs: List[str]) -> List[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    target_stores = list(store_dirs)
    for row in rows:
        if row.store not in target_stores:
            target_stores.append(row.store)
    rows_by_store: Dict[str, List[ProductRow]] = {store: [] for store in target_stores}
    for row in rows:
        rows_by_store.setdefault(row.store, []).append(row)
    output_paths = []
    for store in target_stores:
        output_path = output_dir / f"{store}_每天转化数据.xlsx"
        write_store_workbook(store, rows_by_store.get(store, []), output_path)
        output_paths.append(output_path)
    return output_paths


def run_daily_conversion(config_path: str | Path = "config/config.json", env_path: str | Path = "config/.env", dry_run: bool = False) -> int:
    config = load_config(resolve_project_path(config_path))
    load_env(resolve_project_path(env_path))
    configured_data_dir = resolve_project_path(config.get("data_dir", "data"))
    data_dir = resolve_data_dir(configured_data_dir)
    store_dirs = discover_store_dirs(data_dir)
    mapping_file = resolve_project_path(config.get("mapping_file", "mapping/product_sheet_mapping.csv"))
    log_dir = resolve_project_path(config.get("log_dir", "outputs/logs"))
    output_dir = resolve_project_path(config.get("output_dir", "outputs/daily_convarsion_data"))
    output_mode = config.get("output_mode", "local_excel")
    default_year = int(config.get("default_year", 2026))
    date_output_format = config.get("date_output_format", "mdd")

    mapping = load_mapping(mapping_file)
    all_rows: List[ProductRow] = []
    warnings: List[str] = []
    skipped_products: List[SkippedProduct] = []

    if data_dir != configured_data_dir:
        print(f"[INFO] data_dir 已自动切换为：{data_dir}")
    if not store_dirs:
        print(f"[INFO] data 下没有店铺子目录，按扁平模式读取，店铺标记为“{DEFAULT_FLAT_STORE_NAME}”，映射按 商品ID/商品名 回退匹配。")

    for store, path in iter_input_files(data_dir, store_dirs):
        rows = read_xlsx_rows(path)
        file_date = parse_date_from_workbook(rows, default_year, date_output_format)
        if not file_date:
            file_date = parse_date_from_filename(path, default_year, date_output_format)
        parsed, file_warnings, skipped = parse_workbook_from_rows(path, store, file_date, mapping, rows)
        all_rows.extend(parsed)
        warnings.extend(file_warnings)
        skipped_products.extend(skipped)

    preview_path = write_preview(all_rows, log_dir)
    skipped_path = write_skipped(skipped_products, log_dir)
    diagnostics_path = write_mapping_diagnostics(mapping, log_dir)
    for warning in warnings:
        print(f"[WARN] {warning}")
    print(f"[INFO] 已解析可同步记录：{len(all_rows)} 条")
    if skipped_products:
        reason_counts: Dict[str, int] = {}
        for item in skipped_products:
            key = item.reason.split("：", 1)[0]
            reason_counts[key] = reason_counts.get(key, 0) + 1
        detail = "，".join(f"{key} {count} 条" for key, count in sorted(reason_counts.items(), key=lambda kv: -kv[1]))
        print(f"[INFO] 已跳过：{len(skipped_products)} 条（{detail}），明细：{skipped_path}")
        for key, count in reason_counts.items():
            if key.startswith("映射歧义"):
                print(f"[WARN] 有 {count} 条商品 {key}，请查看 {skipped_path} 中的候选明细。")
    matched_ids = {row.product_id for row in all_rows}
    print(f"[INFO] 映射表命中商品：{len(matched_ids)} 个（映射表共 {len(mapping.entries)} 条）")
    if diagnostics_path:
        print(f"[INFO] 映射表体检：{diagnostics_path}")
    print(f"[INFO] 本地预览：{preview_path}")

    if dry_run:
        print("[INFO] dry-run 模式，不写入飞书。")
        return 0

    if output_mode == "local_excel":
        output_paths = write_local_excels_by_store(all_rows, output_dir, store_dirs)
        for output_path in output_paths:
            print(f"[OK] 已写入本地 Excel：{output_path}")
        return 0

    if output_mode != "feishu":
        raise ValueError("config.json 的 output_mode 只支持 local_excel 或 feishu。")

    app_id = os.getenv("FEISHU_APP_ID", "").strip()
    app_secret = os.getenv("FEISHU_APP_SECRET", "").strip()
    if not app_id or not app_secret:
        raise ValueError("缺少飞书环境变量：FEISHU_APP_ID / FEISHU_APP_SECRET。请填写 .env。")

    clients: Dict[str, FeishuClient] = {}
    sheet_maps: Dict[str, Dict[str, str]] = {}
    skipped_stores: set[str] = set()

    for row in all_rows:
        if row.store in skipped_stores:
            continue
        try:
            spreadsheet_token = spreadsheet_token_for_store(row.store)
        except ValueError as exc:
            print(f"[WARN] {exc} 已跳过 {row.store} 的飞书写入。")
            skipped_stores.add(row.store)
            continue
        if spreadsheet_token not in clients:
            client = FeishuClient(app_id=app_id, app_secret=app_secret, spreadsheet_token=spreadsheet_token)
            client.authenticate()
            clients[spreadsheet_token] = client
            sheet_maps[spreadsheet_token] = client.get_sheet_id_by_name()
        client = clients[spreadsheet_token]
        sheet_map = sheet_maps[spreadsheet_token]
        try:
            sheet_id = sheet_map.get(row.sheet_name)
            if not sheet_id:
                sheet_id = client.create_sheet(row.sheet_name)
                sheet_map[row.sheet_name] = sheet_id
                print(f"[INFO] {row.store} 已创建 sheet：{row.sheet_name}")
            action = client.upsert_daily_row(sheet_id, row.values)
            print(f"[OK] {row.store} {row.product_id} -> {row.sheet_name}: {action}")
        except Exception as exc:
            print(f"[ERROR] {row.store} {row.product_id} -> {row.sheet_name}: {exc}")
            continue

    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="同步 TikTok Shop 商品每天转化数据到飞书表格")
    parser.add_argument("--config", default="config/config.json", help="配置文件路径，默认项目根目录下的 config/config.json")
    parser.add_argument("--env", default="config/.env", help="环境变量文件路径，默认项目根目录下的 config/.env")
    parser.add_argument("--dry-run", action="store_true", help="只解析并输出预览，不写入飞书")
    args = parser.parse_args()
    return run_daily_conversion(args.config, args.env, args.dry_run)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        raise SystemExit(1)
