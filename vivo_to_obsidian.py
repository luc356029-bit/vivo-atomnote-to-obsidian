# -*- coding: utf-8 -*-
"""
原子笔记 -> Obsidian 转换工具
读取 vivo 办公套件备份的原子笔记数据，生成 Obsidian 可直接打开的 Markdown 文件夹。
"""

import sqlite3
import os
import sys
import re
import shutil
import ctypes
from ctypes import wintypes
from datetime import datetime, timezone, timedelta
from html.parser import HTMLParser
from html import unescape
import struct

# Beijing timezone
BJT = timezone(timedelta(hours=8))
UNCATEGORIZED_FOLDER = '未分类'


# ═══════════════════════════════════════════════════════════
#  Windows API: 将文件夹移入回收站
# ═══════════════════════════════════════════════════════════
def send_to_recycle_bin(path):
    """使用 Windows Shell API 将文件/文件夹移入回收站。"""
    if not os.path.exists(path):
        return True

    # SHFileOperationW 结构体
    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [
            ("hwnd", wintypes.HWND),
            ("wFunc", ctypes.c_uint),
            ("pFrom", ctypes.c_wchar_p),
            ("pTo", ctypes.c_wchar_p),
            ("fFlags", ctypes.c_ushort),
            ("fAnyOperationsAborted", wintypes.BOOL),
            ("hNameMappings", ctypes.c_void_p),
            ("lpszProgressTitle", ctypes.c_wchar_p),
        ]

    FO_DELETE = 0x0003
    FOF_ALLOWUNDO = 0x0040        # 移入回收站而非永久删除
    FOF_NOCONFIRMATION = 0x0010   # 不弹确认框
    FOF_SILENT = 0x0004           # 不显示进度条

    # pFrom 需要双 null 结尾
    path_buf = path + '\0\0'
    op = SHFILEOPSTRUCTW()
    op.hwnd = None
    op.wFunc = FO_DELETE
    op.pFrom = path_buf
    op.pTo = None
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT
    op.fAnyOperationsAborted = False
    op.hNameMappings = None
    op.lpszProgressTitle = None

    result = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return result == 0


# ═══════════════════════════════════════════════════════════
#  Windows API: 设置文件创建时间和修改时间
# ═══════════════════════════════════════════════════════════
def set_file_timestamps(filepath, create_time_ms, update_time_ms):
    """设置文件的创建时间和修改时间（毫秒时间戳）。"""
    try:
        mtime_sec = (update_time_ms / 1000) if update_time_ms and update_time_ms > 0 else None
        ctime_sec = (create_time_ms / 1000) if create_time_ms and create_time_ms > 0 else None

        if not mtime_sec and not ctime_sec:
            return

        if not mtime_sec:
            mtime_sec = ctime_sec
        if not ctime_sec:
            ctime_sec = mtime_sec

        # 设置修改时间
        os.utime(filepath, (mtime_sec, mtime_sec))

        # 通过 Windows API 设置创建时间
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateFileW(
            filepath, 0x100, 0, None,  # FILE_WRITE_ATTRIBUTES
            3, 0x02000000, None  # OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS
        )
        if handle != -1:
            EPOCH_DIFF = 116444736000000000
            ctime_ft = int(ctime_sec * 10000000) + EPOCH_DIFF
            mtime_ft = int(mtime_sec * 10000000) + EPOCH_DIFF
            ctime_filetime = wintypes.FILETIME(ctime_ft & 0xFFFFFFFF, ctime_ft >> 32)
            mtime_filetime = wintypes.FILETIME(mtime_ft & 0xFFFFFFFF, mtime_ft >> 32)
            kernel32.SetFileTime(handle, ctypes.byref(ctime_filetime), None, ctypes.byref(mtime_filetime))
            kernel32.CloseHandle(handle)
    except Exception:
        pass  # 静默失败，不影响主流程


# ═══════════════════════════════════════════════════════════
#  HTML → Markdown 转换器
# ═══════════════════════════════════════════════════════════
class VivoHTMLToMarkdown(HTMLParser):
    """将 Vivo 便签 HTML 转换为 Obsidian 兼容的 Markdown。"""

    def __init__(self, resource_map):
        super().__init__()
        self.result = []
        self.current_line = []
        self.resource_map = resource_map
        self.in_list = False
        self.list_type = None
        self.list_counter = 0
        self.in_table = False
        self.table_rows = []
        self.current_row = []
        self.current_cell = []
        self.in_cell = False
        self.bold = False
        self.italic = False
        self.underline = False
        self.strikethrough = False
        self.in_heading = False
        self.heading_level = 0
        self.in_blockquote = False
        self.in_code = False
        self.in_pre = False
        self.tag_stack = []

    def _flush_line(self):
        line = ''.join(self.current_line).strip()
        if line:
            self.result.append(line)
        self.current_line = []

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        self.tag_stack.append(tag)

        if tag == 'p':
            if self.current_line:
                self._flush_line()
        elif tag == 'br':
            self._flush_line()
        elif tag == 'div':
            if self.current_line:
                self._flush_line()
        elif tag in ('b', 'strong'):
            self.bold = True
            self.current_line.append('**')
        elif tag in ('i', 'em'):
            self.italic = True
            self.current_line.append('*')
        elif tag == 'u':
            self.underline = True
            self.current_line.append('<u>')
        elif tag in ('s', 'strike', 'del'):
            self.strikethrough = True
            self.current_line.append('~~')
        elif tag == 'ul':
            if self.current_line:
                self._flush_line()
            self.in_list = True
            self.list_type = 'ul'
        elif tag == 'ol':
            if self.current_line:
                self._flush_line()
            self.in_list = True
            self.list_type = 'ol'
            self.list_counter = 0
        elif tag == 'li':
            if self.current_line:
                self._flush_line()
            if self.list_type == 'ol':
                self.list_counter += 1
                self.current_line.append(f'{self.list_counter}. ')
            else:
                self.current_line.append('- ')
        elif tag == 'blockquote':
            if self.current_line:
                self._flush_line()
            self.in_blockquote = True
        elif tag == 'code':
            self.in_code = True
            self.current_line.append('`')
        elif tag == 'pre':
            if self.current_line:
                self._flush_line()
            self.in_pre = True
            self.result.append('```')
        elif tag in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            if self.current_line:
                self._flush_line()
            self.in_heading = True
            self.heading_level = int(tag[1])
            self.current_line.append('#' * self.heading_level + ' ')
        elif tag == 'table':
            if self.current_line:
                self._flush_line()
            self.in_table = True
            self.table_rows = []
        elif tag == 'tr':
            self.current_row = []
        elif tag in ('td', 'th'):
            self.in_cell = True
            self.current_cell = []
        elif tag == 'vnote-image':
            guid = attrs_dict.get('guid', '')
            filename = attrs_dict.get('filename', '')
            if guid and guid in self.resource_map:
                actual_filename = self.resource_map[guid]
            elif filename:
                actual_filename = filename
            else:
                actual_filename = f'unknown_{guid}'
            self.current_line.append(f'![[attachments/{actual_filename}]]')
        elif tag == 'img':
            src = attrs_dict.get('src', '')
            alt = attrs_dict.get('alt', '')
            if src:
                self.current_line.append(f'![{alt}]({src})')
        elif tag == 'a':
            href = attrs_dict.get('href', '')
            if href:
                self.current_line.append('[')
        elif tag == 'vnote-todo':
            checked = attrs_dict.get('checked', '')
            if checked == 'true':
                self.current_line.append('- [x] ')
            else:
                self.current_line.append('- [ ] ')
        elif tag == 'input':
            input_type = attrs_dict.get('type', '')
            if input_type == 'checkbox':
                checked = 'checked' in attrs_dict
                self.current_line.append('- [x] ' if checked else '- [ ] ')

    def handle_endtag(self, tag):
        if self.tag_stack and self.tag_stack[-1] == tag:
            self.tag_stack.pop()

        if tag == 'p':
            self._flush_line()
        elif tag == 'div':
            self._flush_line()
        elif tag in ('b', 'strong'):
            self.bold = False
            self.current_line.append('**')
        elif tag in ('i', 'em'):
            self.italic = False
            self.current_line.append('*')
        elif tag == 'u':
            self.underline = False
            self.current_line.append('</u>')
        elif tag in ('s', 'strike', 'del'):
            self.strikethrough = False
            self.current_line.append('~~')
        elif tag in ('ul', 'ol'):
            self.in_list = False
            self.list_type = None
        elif tag == 'li':
            self._flush_line()
        elif tag == 'blockquote':
            self.in_blockquote = False
        elif tag == 'code':
            self.in_code = False
            self.current_line.append('`')
        elif tag == 'pre':
            self.in_pre = False
            self._flush_line()
            self.result.append('```')
        elif tag in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            self.in_heading = False
            self._flush_line()
        elif tag == 'table':
            self.in_table = False
            self._render_table()
        elif tag == 'tr':
            if self.current_row:
                self.table_rows.append(self.current_row)
        elif tag in ('td', 'th'):
            self.in_cell = False
            self.current_row.append(''.join(self.current_cell).strip())
            self.current_cell = []
        elif tag == 'a':
            self.current_line.append(']')

    def handle_data(self, data):
        if self.in_cell:
            self.current_cell.append(data)
        elif self.in_pre:
            self.result.append(data)
        elif self.in_blockquote:
            for line in data.split('\n'):
                self.current_line.append('> ' + line)
                self._flush_line()
        else:
            self.current_line.append(data)

    def handle_entityref(self, name):
        char = unescape(f'&{name};')
        (self.current_cell if self.in_cell else self.current_line).append(char)

    def handle_charref(self, name):
        char = unescape(f'&#{name};')
        (self.current_cell if self.in_cell else self.current_line).append(char)

    def _render_table(self):
        if not self.table_rows:
            return
        max_cols = max(len(row) for row in self.table_rows)
        for row in self.table_rows:
            while len(row) < max_cols:
                row.append('')
        col_widths = [3] * max_cols
        for row in self.table_rows:
            for i, cell in enumerate(row):
                col_widths[i] = max(col_widths[i], len(cell))
        header = self.table_rows[0]
        self.result.append('| ' + ' | '.join(cell.ljust(col_widths[i]) for i, cell in enumerate(header)) + ' |')
        self.result.append('| ' + ' | '.join('-' * col_widths[i] for i in range(max_cols)) + ' |')
        for row in self.table_rows[1:]:
            self.result.append('| ' + ' | '.join(cell.ljust(col_widths[i]) for i, cell in enumerate(row)) + ' |')

    def get_markdown(self):
        if self.current_line:
            self._flush_line()
        cleaned = []
        prev_empty = False
        for line in self.result:
            is_empty = line.strip() == ''
            if is_empty and prev_empty:
                continue
            cleaned.append(line)
            prev_empty = is_empty
        return '\n'.join(cleaned)


# ═══════════════════════════════════════════════════════════
#  工具函数
# ═══════════════════════════════════════════════════════════
def html_to_markdown(html_content, resource_map):
    if not html_content:
        return ''
    parser = VivoHTMLToMarkdown(resource_map)
    try:
        parser.feed(html_content)
    except Exception:
        text = re.sub(r'<[^>]+>', '\n', html_content)
        return unescape(text)
    return parser.get_markdown()


def sanitize_filename(name):
    for ch in '<>:"/\\|?*':
        name = name.replace(ch, '_')
    name = re.sub(r'[\x00-\x1f\x7f]', '', name)
    name = name.strip('. ')
    if len(name) > 100:
        name = name[:100]
    return name if name else 'untitled'


def extract_first_text(html_content, max_chars=8):
    if not html_content:
        return ''
    text = re.sub(r'<[^>]+>', '', html_content)
    text = unescape(text).strip()
    for line in text.split('\n'):
        line = line.strip()
        if line:
            return line[:max_chars]
    return ''


def ms_to_date_str(ms_timestamp):
    if not ms_timestamp or ms_timestamp <= 0:
        return 'unknown_date'
    try:
        dt = datetime.fromtimestamp(ms_timestamp / 1000, tz=BJT)
        return dt.strftime('%Y-%m-%d')
    except (ValueError, OSError, OverflowError):
        return 'unknown_date'


def get_desktop_path():
    """获取当前用户的桌面路径。"""
    # 使用 Windows API 获取桌面路径，兼容中英文系统
    buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, 0x0000, None, 0, buf)  # CSIDL_DESKTOP = 0
    return buf.value


# ═══════════════════════════════════════════════════════════
#  主转换流程
# ═══════════════════════════════════════════════════════════
def convert(data_dir, database_dir):
    """执行完整的转换流程。"""

    # 定位关键路径
    db_path = os.path.join(database_dir, 'NoteSync.db')
    resource_dir = os.path.join(data_dir, 'Note', 'Resource')
    desktop = get_desktop_path()
    output_dir = os.path.join(desktop, 'obsidian转换结果')
    attachments_dir = os.path.join(output_dir, 'attachments')

    # ── 验证输入路径 ──
    if not os.path.isfile(db_path):
        print(f"\n  [错误] 找不到数据库文件 NoteSync.db")
        print(f"         检查路径: {db_path}")
        return False

    if not os.path.isdir(os.path.join(data_dir, 'Note')):
        print(f"\n  [错误] data 文件夹中找不到 Note 子目录")
        print(f"         检查路径: {data_dir}")
        return False

    # ── 处理已有输出文件夹：移入回收站 ──
    if os.path.exists(output_dir):
        print(f"\n  [回收] 检测到桌面已有「obsidian转换结果」文件夹，正在移入回收站...")
        success = send_to_recycle_bin(output_dir)
        if not success or os.path.exists(output_dir):
            print(f"  [错误] 无法移入回收站，请手动删除后重试: {output_dir}")
            return False
        print(f"  [完成] 旧文件夹已移入回收站")

    # ── 连接数据库（只读） ──
    db_uri = f"file:{db_path}?mode=ro"
    conn = sqlite3.connect(db_uri, uri=True)
    conn.text_factory = lambda b: b.decode('utf-8', errors='replace')
    cursor = conn.cursor()

    # ── Step 1: 加载文件夹 ──
    print(f"\n  [1/5] 加载文件夹信息...")
    cursor.execute("SELECT guid, nameNew FROM NoteBook")
    notebooks = {}
    for row in cursor.fetchall():
        name = row[1].strip() if row[1] else row[0]
        notebooks[row[0]] = name
    print(f"        找到 {len(notebooks)} 个文件夹")

    # ── Step 2: 加载图片资源 ──
    print(f"  [2/5] 加载图片资源...")
    cursor.execute("SELECT guid, noteGuid, name, mime FROM Resource")
    resource_map = {}
    note_resources = {}
    resource_count = 0
    for row in cursor.fetchall():
        res_guid, note_guid, res_name, res_mime = row
        if res_mime in ('table',):
            continue
        resource_map[res_guid] = res_name
        if note_guid not in note_resources:
            note_resources[note_guid] = []
        note_resources[note_guid].append({'guid': res_guid, 'name': res_name})
        resource_count += 1
    print(f"        找到 {resource_count} 个图片资源")

    # ── Step 3: 创建输出目录 ──
    print(f"  [3/5] 创建输出目录...")
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(attachments_dir, exist_ok=True)

    folder_paths = {}
    for guid, name in notebooks.items():
        folder_path = os.path.join(output_dir, sanitize_filename(name))
        os.makedirs(folder_path, exist_ok=True)
        folder_paths[guid] = folder_path

    uncategorized_path = os.path.join(output_dir, UNCATEGORIZED_FOLDER)
    os.makedirs(uncategorized_path, exist_ok=True)
    print(f"        创建了 {len(folder_paths)} 个文件夹 + 未分类文件夹")

    # ── Step 4: 转换笔记 ──
    print(f"  [4/5] 转换笔记中...")
    cursor.execute("""
        SELECT guid, title, contentNote, noteBookGuid, createTime, updateTime, type
        FROM Note ORDER BY createTime ASC
    """)
    notes = cursor.fetchall()
    print(f"        总共 {len(notes)} 条笔记")

    used_filenames = {}
    success_count = 0
    skip_count = 0
    image_notes = 0

    for note in notes:
        guid, title, content_html, notebook_guid, create_time, update_time, note_type = note

        # 确定输出文件夹
        folder = folder_paths.get(notebook_guid, uncategorized_path)

        # 确定文件名
        if title and title.strip():
            base_name = title.strip()
        else:
            first_text = extract_first_text(content_html, max_chars=8)
            base_name = first_text if first_text else ms_to_date_str(create_time) + '_untitled'

        safe_name = sanitize_filename(base_name) or 'untitled'

        # 重复标题加序号
        if folder not in used_filenames:
            used_filenames[folder] = {}
        if safe_name in used_filenames[folder]:
            used_filenames[folder][safe_name] += 1
            final_name = f"{safe_name}（{used_filenames[folder][safe_name]}）"
        else:
            used_filenames[folder][safe_name] = 0
            final_name = safe_name

        # 转换内容
        md_content = html_to_markdown(content_html, resource_map)

        # 写入文件
        md_filepath = os.path.join(folder, f"{final_name}.md")
        try:
            with open(md_filepath, 'w', encoding='utf-8') as f:
                f.write(md_content)
            success_count += 1
        except Exception as e:
            print(f"        [失败] 写入失败: {final_name} - {e}")
            skip_count += 1
            continue

        # 设置时间戳
        set_file_timestamps(md_filepath, create_time, update_time)

        if guid in note_resources:
            image_notes += 1

    print(f"        [完成] 成功: {success_count} 条 | 跳过: {skip_count} 条 | 含图片: {image_notes} 条")

    # ── Step 5: 复制图片 ──
    print(f"  [5/5] 复制图片资源...")
    copied = 0
    missing = 0

    for res_guid, res_name in resource_map.items():
        res_folder = os.path.join(resource_dir, res_guid)
        if not os.path.isdir(res_folder):
            missing += 1
            continue

        # 遍历 Resource/{guid}/{hash_dir}/{actual_file}
        source_file = None
        for item in os.listdir(res_folder):
            item_path = os.path.join(res_folder, item)
            if os.path.isdir(item_path):
                for f in os.listdir(item_path):
                    if not f.startswith('thumb_'):
                        source_file = os.path.join(item_path, f)
                        break
            elif os.path.isfile(item_path) and not item.startswith('thumb_'):
                source_file = item_path

        if not source_file:
            missing += 1
            continue

        dest_file = os.path.join(attachments_dir, res_name)
        if os.path.exists(dest_file):
            base, ext = os.path.splitext(res_name)
            counter = 1
            while os.path.exists(dest_file):
                dest_file = os.path.join(attachments_dir, f"{base}_{counter}{ext}")
                counter += 1

        try:
            shutil.copy2(source_file, dest_file)
            copied += 1
        except Exception:
            missing += 1

    print(f"        [完成] 复制成功: {copied} 个 | 缺失: {missing} 个")

    conn.close()

    # ── 汇总 ──
    print(f"\n{'=' * 56}")
    print(f"  转换完成!")
    print(f"  输出位置: {output_dir}")
    print(f"  笔记总数: {success_count}")
    print(f"  文件夹数: {len(folder_paths) + 1}")
    print(f"  图片总数: {copied}")
    print(f"{'=' * 56}")
    return True


# ═══════════════════════════════════════════════════════════
#  入口：交互式提示
# ═══════════════════════════════════════════════════════════
def main():
    # 确保控制台支持 UTF-8
    if sys.platform == 'win32':
        os.system('chcp 65001 >nul 2>&1')
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass

    print()
    print("=" * 56)
    print("      原子笔记 -> Obsidian 转换工具 v1.0")
    print()
    print("  将 vivo 办公套件备份的原子笔记转换为 Obsidian Markdown 文件夹")
    print("  转换结果将保存到桌面「obsidian转换结果」文件夹")
    print("=" * 56)
    print()

    # 提示用户输入路径
    print("  请输入以下两个路径（从 vivo 备份中获取）：")
    print()
    print("  提示: data 文件夹通常位于 vivo 办公套件备份目录:")
    print("        C:\\Users\\用户名\\AppData\\Roaming\\pcsuite\\data")
    print()

    data_dir = input("  1. 请输入 data 文件夹路径: ").strip().strip('"').strip("'")
    database_dir = input("  2. 请输入 database 文件夹路径: ").strip().strip('"').strip("'")

    # 基本路径检查
    if not data_dir or not os.path.isdir(data_dir):
        print(f"\n  [错误] data 文件夹路径无效或不存在: {data_dir}")
        input("\n  按回车键退出...")
        return

    if not database_dir or not os.path.isdir(database_dir):
        print(f"\n  [错误] database 文件夹路径无效或不存在: {database_dir}")
        input("\n  按回车键退出...")
        return

    print()
    print("-" * 56)
    print("  开始转换...")
    print("-" * 56)

    success = convert(data_dir, database_dir)

    if success:
        print(f"\n  请用 Obsidian 打开桌面上的「obsidian转换结果」文件夹。")
    else:
        print(f"\n  转换过程中出现错误，请检查路径后重试。")

    print()
    input("  按回车键退出...")


if __name__ == '__main__':
    main()
