"""202605.xlsx (テキストマスター一覧) から、支店別の顔写真レイアウトを自動生成する。

入力: 202605.xlsx 形式の「テキストマスター体制表」
処理: 各支店ブロックをパース → 支店ごとのシートを生成 → DB から写真を貼り込み
出力: 2605顔写真体制.xlsx と同等の「支店別グリッドレイアウト」xlsx

実行例:
    python3 -m src.cli.build_from_master file/202605.xlsx file/output.xlsx
"""
from __future__ import annotations
import argparse
import re
import sys
import warnings
from dataclasses import dataclass, field
from pathlib import Path

warnings.filterwarnings("ignore")

from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

from ..config import Settings, DataPaths, PLACEHOLDER_PHOTO_PATH
from ..db import Database, EmploymentStatus
from ..services import EmployeeService, PhotoService


# ============================================================
# 入力（202605.xlsx）の構造定義
# ============================================================
COL_F = 6   # ブロックヘッダ列（"集合XX支店" や "支店長" 等）
COL_G = 7   # 上記の対応値列（支店長の氏名等）
COL_I = 9   # セクション列 (●営業課, ●設計課)
COL_J = 10  # 課名列 (営業1課, 営業2課, 営業3課)

# 名前列と入社年列のペア（総合職）。L,M / N,O / P,Q / ...
TOTAL_PAIRS = [(c, c + 1) for c in range(12, 32, 2)]  # L-AE まで10ペア
JIMU_COLS = [32, 33]      # AF, AG: 実務職・事務
KEIYAKU_COLS = [34, 35]   # AH, AI: 契約社員
HAKEN_COLS = [36, 37, 38] # AJ, AK, AL: 派遣社員
HOLIDAY_COL = 39          # AM: 休職中（写真は不要だが情報として保持）

BRANCH_RE = re.compile(r"^(集合|《本部スタッフ》|技術部|事業推進部)")
EXCLUDE_BRANCH_RE = re.compile(r"^(本部長|本部長・部長|本部長・部長・支店長)")
SECTION_RE = re.compile(r"^[●○]\s*(.+)")
ROLE_MARK_CHARS = "＊*☆■◆○◯〇▽●◎▲△□"
ROLE_MARK_RE = re.compile(rf"^([{ROLE_MARK_CHARS}]+)?(.+)$")

# 名前末尾に付く役職/契約区分マーカー (M=マネージャー、C=契約 など)
TRAILING_MARK_RE = re.compile(r"[MＭCＣ]+$")

# 飾り文字列・非人物セルを除外するためのパターン
NOT_A_PERSON_RE = re.compile(r"^[《【〈\[].*[》】〉\]]$|^#REF!$|^＝|^=|^兼務除く$")


@dataclass
class PersonEntry:
    raw_text: str
    name: str
    marks: str
    year: str | None
    category: str = "総合職"   # 総合職 / 実務職 / 契約 / 派遣


@dataclass
class CourseRow:
    course_name: str   # "営業1課" / "" / "設計課"
    persons: list[PersonEntry]


@dataclass
class BranchData:
    branch_name: str
    top_positions: list[tuple[str, str]]   # (役職, 氏名)
    sections: dict[str, list[CourseRow]]   # "営業課" → [CourseRow, ...]
    jimu: list[PersonEntry]
    keiyaku: list[PersonEntry]
    haken: list[PersonEntry]


_CONCURRENT_RE = re.compile(r"^[\s　]*兼\s*[)）)）]\s*")


def _parse_person(text: str, year, category: str = "総合職") -> PersonEntry | None:
    """セルの文字列から人物データを抽出。人物セルでなければ None を返す。"""
    text = text.strip()
    if not text or NOT_A_PERSON_RE.match(text):
        return None
    # 兼務プレフィックス除去（marks に "兼" として保持）
    extra_marks = ""
    if _CONCURRENT_RE.match(text):
        text = _CONCURRENT_RE.sub("", text)
        extra_marks = "兼"
    # 先頭の役職記号 (☆, ＊, ○ など)
    m = ROLE_MARK_RE.match(text)
    if m and m.group(1):
        marks = m.group(1)
        name = m.group(2).strip()
    else:
        marks = ""
        name = text
    # 末尾の M/C 等（マネージャー、契約等）。マッチング用に除去し marks に追加
    tail = TRAILING_MARK_RE.search(name)
    if tail:
        marks = marks + tail.group(0)
        name = TRAILING_MARK_RE.sub("", name).strip()
    if extra_marks:
        marks = extra_marks + marks
    # プレースホルダー名（未確定枠）は人物として扱わない
    if name in ("新人", "新 人", "（新人）", "(新人)"):
        return None
    return PersonEntry(
        raw_text=text, name=name, marks=marks,
        year=(str(year).strip() if year is not None else None),
        category=category,
    )


def parse_master(path: Path) -> list[BranchData]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    branches: list[BranchData] = []
    current: BranchData | None = None
    current_section: str | None = None

    def flush():
        nonlocal current
        if current:
            branches.append(current)
        current = None

    for row in range(1, ws.max_row + 1):
        f = ws.cell(row=row, column=COL_F).value
        g = ws.cell(row=row, column=COL_G).value
        i = ws.cell(row=row, column=COL_I).value
        j = ws.cell(row=row, column=COL_J).value

        # ブロックヘッダ判定（branch切替）。同じ行のセクション/人物データは
        # この後続けて処理するため continue しない。
        if f and isinstance(f, str):
            text = f.strip()
            if EXCLUDE_BRANCH_RE.match(text):
                continue   # 集計行は完全スキップ
            if BRANCH_RE.match(text):
                flush()
                current = BranchData(
                    branch_name=text,
                    top_positions=[], sections={},
                    jimu=[], keiyaku=[], haken=[],
                )
                current_section = None

        if not current:
            continue

        # 役職 + 氏名（F+G）。ただし branch ヘッダ自身は除く。
        if f and isinstance(f, str) and g:
            text = f.strip()
            if not BRANCH_RE.match(text):
                current.top_positions.append((text, str(g).strip()))

        # セクションヘッダ
        if i and isinstance(i, str):
            m = SECTION_RE.match(i.strip())
            if m:
                current_section = m.group(1).strip()
                current.sections.setdefault(current_section, [])

        # 課行（総合職）
        if current_section is not None:
            course_name = ""
            if j and isinstance(j, str):
                course_name = j.strip()
            persons = []
            for name_col, year_col in TOTAL_PAIRS:
                n = ws.cell(row=row, column=name_col).value
                y = ws.cell(row=row, column=year_col).value
                if n and isinstance(n, str) and n.strip():
                    p = _parse_person(str(n), y, "総合職")
                    if p:
                        persons.append(p)
            if persons:
                current.sections[current_section].append(
                    CourseRow(course_name=course_name, persons=persons)
                )

        # 実務職・契約・派遣
        for col, bucket, cat in [
            *[(c, current.jimu, "実務職") for c in JIMU_COLS],
            *[(c, current.keiyaku, "契約") for c in KEIYAKU_COLS],
            *[(c, current.haken, "派遣") for c in HAKEN_COLS],
        ]:
            v = ws.cell(row=row, column=col).value
            if v and isinstance(v, str) and v.strip():
                p = _parse_person(v, None, cat)
                if p:
                    bucket.append(p)

    flush()
    return branches


# ============================================================
# 出力レイアウト
# ============================================================
PERSON_COL_W = 6     # 1人あたりの幅（列数）。区切り1列含む
PHOTO_ROWS = 8       # 写真の高さ（行数）
LABEL_ROWS = 2       # 写真直下のラベル行数 (kana+year / kanji)
LEADER_LABEL_ROWS = 1  # 課長/室長 役職ラベル用の上段行
PERSON_TOTAL_H = LEADER_LABEL_ROWS + PHOTO_ROWS + LABEL_ROWS + 1  # 課行の総高さ

LEFT_LABEL_COLS = 7  # 「マーケティング室」「不動産事業推進室」が収まる幅
SHEET_START_COL = 1
DATA_COL_END = 75   # 全体枠の右端列（左ラベル拡大に伴い若干広げる）

UNIFORM_COL_WIDTH = 2.6   # 1列の幅（小さく刻む）
PHOTO_ROW_HEIGHT = 15.0   # 写真領域の行高
LABEL_ROW_HEIGHT = 14.0
HEADER_ROW_HEIGHT = 24.0

# カラーパレット（スクリーンショット準拠）
COLOR_TITLE_BG = "1F4E79"   # 濃紺（タイトル）
COLOR_LABEL_BG = "2E7D5B"   # ティール緑（役職・セクションラベル）
COLOR_LABEL_BORDER = "1B5E40"  # ラベル枠
COLOR_OUTER_BORDER = "1F4E79"  # 全体外枠（濃紺）
COLOR_CELL_BORDER = "9DB7D9"   # 内部薄罫線
COLOR_TEXT_KANA = "555555"
COLOR_TEXT_YEAR = "555555"
COLOR_TEXT_KANJI = "111111"
COLOR_TEXT_UNREG = "CC0000"
COLOR_TEXT_LEAVE = "888888"

THICK_BORDER = Side(style="thick", color=COLOR_OUTER_BORDER)
THIN_BORDER = Side(style="thin", color=COLOR_CELL_BORDER)
LABEL_BORDER = Side(style="medium", color=COLOR_LABEL_BORDER)


def _norm(s: str) -> str:
    """名前正規化（兼務マーカー除去・スペース除去）。"""
    s = re.sub(r"^[\s　]*兼\s*[)）)）]\s*", "", s or "")
    return s.replace("　", "").replace(" ", "").strip()


def _shorten_branch(name: str) -> str:
    """シート名として使用可能な形式に整形。

    マスター左側に書かれた支店名・部署名をそのまま使う方針。
    Excel のシート名禁止文字 (《 》 : \\ / ? * [ ]) のみ除去する。
    """
    s = name
    for ch in ("《", "》", ":", "\\", "/", "?", "*", "[", "]"):
        s = s.replace(ch, "")
    s = s.strip()
    # 31文字上限
    return s[:31] if len(s) > 31 else s


def _to_year4(raw) -> str | None:
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    try:
        n = int(float(s))
    except ValueError:
        return s
    if n < 100:
        return str(1900 + n if n >= 50 else 2000 + n)
    return str(n)


def _setup_sheet(ws) -> None:
    for c in range(1, DATA_COL_END + 4):
        ws.column_dimensions[get_column_letter(c)].width = UNIFORM_COL_WIDTH


def _setup_page(ws) -> None:
    """印刷時のページレイアウトを整える（A3横、余白小さめ、用紙中央寄せ）。"""
    from openpyxl.worksheet.page import PageMargins, PrintOptions
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A3  # A3 横（広い体制図向け）
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # 幅優先、高さは何ページにわたっても可
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=0.4, right=0.4, top=0.4, bottom=0.4,
        header=0.2, footer=0.2,
    )
    ws.print_options = PrintOptions(
        horizontalCentered=True,
        verticalCentered=False,
    )
    # 印刷範囲を全データ領域に設定
    ws.print_area = f"A1:{get_column_letter(DATA_COL_END)}{max(ws.max_row, 100)}"


def _set_block_row_heights(ws, row_start: int) -> None:
    """1人ブロック分の行高を設定（写真領域＋ラベル2行）。"""
    for r in range(row_start, row_start + PHOTO_ROWS):
        ws.row_dimensions[r].height = PHOTO_ROW_HEIGHT
    ws.row_dimensions[row_start + PHOTO_ROWS].height = LABEL_ROW_HEIGHT
    ws.row_dimensions[row_start + PHOTO_ROWS + 1].height = LABEL_ROW_HEIGHT


def _label_box(ws, row: int, col_start: int, col_end: int, text: str,
               *, size: int = 11, height: float | None = 18):
    """役職ラベル/セクションヘッダ用のテーマ統一されたボックス。"""
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = Font(size=size, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=COLOR_LABEL_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(left=LABEL_BORDER, right=LABEL_BORDER,
                         top=LABEL_BORDER, bottom=LABEL_BORDER)
    if col_end > col_start:
        ws.merge_cells(start_row=row, start_column=col_start,
                       end_row=row, end_column=col_end)
        # 結合先セルにも罫線が必要
        for c in range(col_start + 1, col_end + 1):
            ws.cell(row=row, column=c).border = Border(
                left=LABEL_BORDER, right=LABEL_BORDER,
                top=LABEL_BORDER, bottom=LABEL_BORDER,
            )
    if height:
        ws.row_dimensions[row].height = height


def _label_text(ws, row: int, col_start: int, col_end: int, text: str,
                *, size: int = 9, bold: bool = False,
                color: str = "111111", align: str = "center"):
    """写真下のラベル行に統一書式でテキストを書く（複数セルをマージ）。"""
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = Font(size=size, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if col_end > col_start:
        ws.merge_cells(start_row=row, start_column=col_start,
                       end_row=row, end_column=col_end)


def _keep_side(side):
    """既存の Side が定義済みなら返す。なければ None。"""
    return side if (side is not None and side.style) else None


def _apply_outer_border(ws, row_end: int) -> None:
    """シート全体を太い濃紺枠で囲む（既存のセル罫線は保持）。"""
    col_start, col_end = 1, DATA_COL_END
    for c in range(col_start, col_end + 1):
        top_cell = ws.cell(row=1, column=c)
        top_cell.border = Border(
            top=THICK_BORDER,
            left=_keep_side(top_cell.border.left),
            right=_keep_side(top_cell.border.right),
            bottom=_keep_side(top_cell.border.bottom),
        )
        bot_cell = ws.cell(row=row_end, column=c)
        bot_cell.border = Border(
            bottom=THICK_BORDER,
            top=_keep_side(bot_cell.border.top),
            left=_keep_side(bot_cell.border.left),
            right=_keep_side(bot_cell.border.right),
        )
    for r in range(1, row_end + 1):
        l_cell = ws.cell(row=r, column=col_start)
        l_cell.border = Border(
            left=THICK_BORDER,
            top=_keep_side(l_cell.border.top),
            right=_keep_side(l_cell.border.right),
            bottom=_keep_side(l_cell.border.bottom),
        )
        r_cell = ws.cell(row=r, column=col_end)
        r_cell.border = Border(
            right=THICK_BORDER,
            top=_keep_side(r_cell.border.top),
            left=_keep_side(r_cell.border.left),
            bottom=_keep_side(r_cell.border.bottom),
        )


def _insert_photo(ws, col0_1: int, row0_1: int, w_cols: int, h_rows: int, path: Path) -> None:
    """セル範囲にぴったり貼り付ける。img.width/height はあえてセットしない
    （TwoCellAnchor がストレッチを担うため、px指定すると逆効果になる）。"""
    img = XLImage(str(path))
    anchor = TwoCellAnchor(
        editAs="oneCell",
        _from=AnchorMarker(col=col0_1 - 1, colOff=0, row=row0_1 - 1, rowOff=0),
        to=AnchorMarker(col=col0_1 - 1 + w_cols, colOff=0, row=row0_1 - 1 + h_rows, rowOff=0),
    )
    img.anchor = anchor
    ws.add_image(img)


def _draw_person_block(
    ws,
    col_start_1: int,
    row_start_1: int,
    person: PersonEntry,
    emp,
    photo_path: Path | None,
    is_on_leave: bool,
) -> None:
    """1人分（写真＋kana＋年＋kanji）を描画。

    レイアウト:
      [写真領域 (PHOTO_ROWS 行 × PERSON_COL_W-1 列)]
      [カナ氏名 (左寄せ) | 入社年 (右寄せ)]
      [漢字氏名 (中央)]
    """
    _set_block_row_heights(ws, row_start_1)

    col_end = col_start_1 + PERSON_COL_W - 1  # ブロック右端列（区切り余白含まず）
    photo_col_end = col_end - 1  # 区切り1列を残す

    # 写真領域＋ラベル領域を囲む枠（ブロック外枠）
    block_border = Side(style="thin", color="6E8FB7")
    for r in range(row_start_1, row_start_1 + PHOTO_ROWS + LABEL_ROWS):
        for c in range(col_start_1, photo_col_end + 1):
            existing = ws.cell(row=r, column=c).border
            is_top = (r == row_start_1)
            is_bottom = (r == row_start_1 + PHOTO_ROWS + LABEL_ROWS - 1)
            is_left = (c == col_start_1)
            is_right = (c == photo_col_end)
            ws.cell(row=r, column=c).border = Border(
                left=block_border if is_left else None,
                right=block_border if is_right else None,
                top=block_border if is_top else None,
                bottom=block_border if is_bottom else None,
            )

    # 写真領域: 本物の写真 → なければプレースホルダー画像を貼る
    img_to_use = photo_path if (photo_path and not is_on_leave) else PLACEHOLDER_PHOTO_PATH
    if img_to_use and img_to_use.is_file():
        # 写真はブロック領域いっぱいに広げる（左上に寄らないように）
        _insert_photo(ws, col_start_1, row_start_1,
                      PERSON_COL_W - 1, PHOTO_ROWS, img_to_use)
    if is_on_leave:
        # 休職中: プレースホルダー上に「休職中」ラベルをかぶせる
        cell = ws.cell(row=row_start_1 + PHOTO_ROWS - 1, column=col_start_1,
                       value="休職中")
        cell.font = Font(size=9, italic=True, color=COLOR_TEXT_LEAVE)
        cell.alignment = Alignment(horizontal="center", vertical="bottom")

    label_row_1 = row_start_1 + PHOTO_ROWS
    label_row_2 = label_row_1 + 1

    # カナ氏名 (左半分) + 入社年 (右半分・数値型で書き込み)
    half = max(1, (photo_col_end - col_start_1 + 1) // 2)
    kana = (emp.name_kana if emp else None) or ""
    year_int: int | None = None
    y4 = _to_year4(person.year)
    if y4 and y4.isdigit():
        year_int = int(y4)
    elif emp and emp.join_year:
        year_int = int(emp.join_year)

    _label_text(ws, label_row_1, col_start_1, col_start_1 + half - 1,
                kana, size=8, color=COLOR_TEXT_KANA, align="left")
    if year_int is not None:
        cell = ws.cell(row=label_row_1, column=col_start_1 + half, value=year_int)
        cell.font = Font(size=8, color=COLOR_TEXT_YEAR)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "0"
        if photo_col_end > col_start_1 + half:
            ws.merge_cells(start_row=label_row_1, start_column=col_start_1 + half,
                           end_row=label_row_1, end_column=photo_col_end)

    # 漢字氏名（役職記号付き、中央寄せ）
    name = emp.name if emp else person.name
    display = (person.marks or "") + name
    _label_text(ws, label_row_2, col_start_1, photo_col_end,
                display, size=10, bold=True, color=COLOR_TEXT_KANJI, align="center")


def _determine_leader_label(section_name: str, leader: PersonEntry) -> str:
    """先頭人物の役職ラベル（課長/課長代務/室長 等）を決定する。"""
    marks = leader.marks or ""
    is_acting = ("＊" in marks) or ("*" in marks)
    is_concurrent = "兼" in marks
    base = "室長" if ("室" in section_name and "課" not in section_name) else "課長"
    if is_acting:
        return f"{base}代務"
    if is_concurrent:
        return f"{base}（兼）"
    return base


def _build_section(
    ws,
    row_start: int,
    section_name: str,
    course_rows: list[CourseRow],
    lookup_fn,
) -> tuple[int, list[int]]:
    """1セクション（●営業課 or ●設計課）を描画する。

    Returns:
        (次の開始行, 各課行の写真開始行のリスト)
    """
    title = f"● {' '.join(section_name)}" if len(section_name) <= 4 else f"● {section_name}"
    _label_box(ws, row_start, 1, LEFT_LABEL_COLS, title, size=12, height=22)
    row = row_start + 1
    photo_row_starts: list[int] = []

    for idx, cr in enumerate(course_rows):
        # === リーダー役職ラベル行（写真の上）===
        if cr.persons and (cr.course_name or idx == 0):
            lbl = _determine_leader_label(section_name, cr.persons[0])
            if lbl:
                col_start = SHEET_START_COL + LEFT_LABEL_COLS
                col_end = col_start + PERSON_COL_W - 2
                _label_box(ws, row, col_start, col_end, lbl, size=10, height=16)
        row += LEADER_LABEL_ROWS

        # 写真の開始位置をここで記録
        photo_row_starts.append(row)

        # === 左に課ラベル（写真行の中段）===
        if cr.course_name:
            mid_row = row + (PHOTO_ROWS // 2)
            _label_box(ws, mid_row, 1, LEFT_LABEL_COLS - 1,
                       cr.course_name, size=10, height=16)

        # === 各人ブロック ===
        for i, person in enumerate(cr.persons):
            col_start = SHEET_START_COL + LEFT_LABEL_COLS + i * PERSON_COL_W
            emp, photo_path, is_on_leave = lookup_fn(person.name)
            _draw_person_block(ws, col_start, row, person, emp, photo_path, is_on_leave)
        row += PHOTO_ROWS + LABEL_ROWS + 1

    return row + 1, photo_row_starts


def _build_extra_persons(
    ws,
    row_start: int,
    title: str,
    persons: list[PersonEntry],
    lookup_fn,
) -> int:
    """実務職/契約/派遣のサイドリスト的ブロック。"""
    if not persons:
        return row_start
    _label_box(ws, row_start, 1, LEFT_LABEL_COLS, f"《{title}》", size=11, height=20)
    row = row_start + 1
    per_row = 8
    for i, person in enumerate(persons):
        if i > 0 and i % per_row == 0:
            row += PERSON_TOTAL_H
        col_start = SHEET_START_COL + LEFT_LABEL_COLS + (i % per_row) * PERSON_COL_W
        emp, photo_path, is_on_leave = lookup_fn(person.name)
        _draw_person_block(ws, col_start, row, person, emp, photo_path, is_on_leave)
    return row + PERSON_TOTAL_H + 1


def build_workbook(
    branches: list[BranchData],
    employee_service: EmployeeService,
    photo_service: PhotoService,
) -> tuple[Workbook, dict]:
    stats = {"placed": 0, "unmatched": [], "on_leave": []}

    def lookup(name: str):
        # repository の find_by_text を使う（完全一致 + 前方一致対応）
        candidates_active = employee_service.repo.find_by_text(name, only_active=True)
        if candidates_active:
            emp = candidates_active[0]
            photo_path = photo_service.resolve(emp.photo_path)
            if photo_path:
                stats["placed"] += 1
            return emp, photo_path, False

        # 在職で見つからなければ休職中含めて検索
        all_candidates = employee_service.repo.find_by_text(name, only_active=False)
        if all_candidates:
            stats["on_leave"].append(name)
            return all_candidates[0], None, True

        stats["unmatched"].append(name)
        return None, None, False

    wb = Workbook()
    wb.remove(wb.active)

    for branch in branches:
        sheet_name = _shorten_branch(branch.branch_name)[:31] or "Sheet"
        if sheet_name in wb.sheetnames:
            sheet_name = sheet_name + "_2"
        ws = wb.create_sheet(sheet_name)
        ws.sheet_view.showGridLines = False  # グリッド線を非表示
        _setup_sheet(ws)
        _setup_page(ws)

        # タイトル（濃紺で支店名を表示）
        tcell = ws.cell(row=1, column=1, value=branch.branch_name)
        tcell.font = Font(bold=True, size=14, color="FFFFFF")
        tcell.fill = PatternFill("solid", fgColor=COLOR_TITLE_BG)
        tcell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=DATA_COL_END)
        ws.row_dimensions[1].height = HEADER_ROW_HEIGHT
        row = 3

        # 上位役職（支店長/支店付/兼務）
        if branch.top_positions:
            persons_raw = [_parse_person(name, None) for _, name in branch.top_positions]
            pairs = [(lbl, p) for (lbl, _), p in zip(branch.top_positions, persons_raw) if p]
            for i, (lbl, p) in enumerate(pairs):
                col_start = SHEET_START_COL + LEFT_LABEL_COLS + i * PERSON_COL_W
                col_end = col_start + PERSON_COL_W - 2  # 区切り1列を残す
                _label_box(ws, row, col_start, col_end, lbl, size=10, height=18)
            row += 1
            for i, (_, p) in enumerate(pairs):
                col_start = SHEET_START_COL + LEFT_LABEL_COLS + i * PERSON_COL_W
                emp, photo_path, is_on_leave = lookup(p.name)
                _draw_person_block(ws, col_start, row, p, emp, photo_path, is_on_leave)
            row += PERSON_TOTAL_H + 1

        # 各セクション（●営業課/●設計課）
        section_photo_rows: dict[str, list[int]] = {}
        for section_name, course_rows in branch.sections.items():
            row, photo_starts = _build_section(ws, row, section_name, course_rows, lookup)
            section_photo_rows[section_name] = photo_starts

        # 派遣社員＋契約社員 → 設計課の2段目 (なければ最終課行) の右側にラベルなし配置
        # ただしメイン人物と重ならないよう右側に追いやる。あふれる場合は次の行に。
        extras = list(branch.haken) + list(branch.keiyaku)
        if extras:
            design_section_name = next(
                (sn for sn in section_photo_rows if "設計" in sn), None
            )
            target_row: int | None = None
            persons_on_target_row = 0
            if design_section_name:
                starts = section_photo_rows[design_section_name]
                design_courses = branch.sections.get(design_section_name, [])
                if len(starts) >= 2:
                    target_row = starts[1]
                    persons_on_target_row = (len(design_courses[1].persons)
                                             if len(design_courses) >= 2 else 0)
                elif starts:
                    target_row = starts[0]
                    persons_on_target_row = (len(design_courses[0].persons)
                                             if design_courses else 0)
            if target_row is None:
                all_starts = [(r, 0) for starts in section_photo_rows.values() for r in starts]
                if all_starts:
                    target_row = all_starts[-1][0]

            if target_row is not None:
                count = len(extras)
                # メイン人物が占めている右端の次の列
                main_right_edge = SHEET_START_COL + LEFT_LABEL_COLS + persons_on_target_row * PERSON_COL_W
                # 右寄せで配置したい開始列
                desired_start = DATA_COL_END - count * PERSON_COL_W + 1
                start_col = max(desired_start, main_right_edge + 1)
                # それでも右端を超える場合は、SA の直上に新しい行を立てて配置
                if start_col + count * PERSON_COL_W > DATA_COL_END + 1:
                    target_row = row
                    start_col = max(SHEET_START_COL + LEFT_LABEL_COLS,
                                    DATA_COL_END - count * PERSON_COL_W + 1)
                    row += PHOTO_ROWS + LABEL_ROWS + 1
                for i, person in enumerate(extras):
                    col_start = start_col + i * PERSON_COL_W
                    emp, photo_path, is_on_leave = lookup(person.name)
                    _draw_person_block(ws, col_start, target_row, person,
                                       emp, photo_path, is_on_leave)

        # SA（実務職）はラベル付きでメイン領域の下に配置
        row = _build_extra_persons(ws, row, "SA", branch.jimu, lookup)

        _apply_outer_border(ws, max(row, 5))

    return wb, stats


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("master", type=Path, help="入力 (202605.xlsx 形式のマスター)")
    parser.add_argument("output", type=Path, help="出力 xlsx")
    parser.add_argument("--data-dir", type=Path, default=None)
    args = parser.parse_args()

    if not args.master.is_file():
        print(f"[error] 入力ファイルが見つかりません: {args.master}", file=sys.stderr)
        return 1

    settings = Settings()
    data_dir = args.data_dir or settings.data_dir
    paths = DataPaths(data_dir)
    db = Database(paths.db_path)
    db.create_all()
    photos = PhotoService(paths.photos_dir)
    es = EmployeeService(db, photos)

    print(f"📂 入力: {args.master}")
    print(f"📦 DB  : {paths.db_path}")
    branches = parse_master(args.master)
    print(f"\n=== 解析結果 ===")
    for b in branches:
        total = sum(len(cr.persons) for cs in b.sections.values() for cr in cs)
        extra = len(b.jimu) + len(b.keiyaku) + len(b.haken)
        top = len(b.top_positions)
        print(f"  {b.branch_name}: 上位{top} / 総合職{total} / その他{extra}")

    wb, stats = build_workbook(branches, es, photos)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.output)

    print(f"\n=== 結果 ===")
    print(f"  写真挿入: {stats['placed']}枚")
    print(f"  未マッチ: {len(stats['unmatched'])}名")
    if stats['unmatched']:
        print(f"    {sorted(set(stats['unmatched']))[:15]}")
    print(f"  休職中  : {len(stats['on_leave'])}名")
    print(f"\n✅ 出力完了: {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
