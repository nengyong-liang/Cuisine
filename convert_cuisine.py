"""
美食知识库批量转换脚本
将 WPS 云盘中的美食记录转换为本地 Markdown/CSV 知识库

支持格式：
- .docx → .md (pandoc)
- .doc → .md (Word COM + pandoc)
- .xlsx/.xls → .csv (pandas)
- .pdf → .md (pdfplumber)
- .txt → .md (直接复制 + frontmatter)
- .jpg/.png → 原样复制
"""

import os
import shutil
import subprocess
import tempfile
import time
from datetime import datetime
from pathlib import Path

# ==================== 配置区域 ====================
SOURCE_ROOT = r"C:\Users\Administrator\WPSDrive\1599884454\WPS云盘\1_LNY\2_生活笔记整理\1生活\2饭_做饭\2食堂_外面吃饭_零食\2重庆-吃"
TARGET_ROOT = r"d:\1_LNY\code\Cuisine\重庆-吃"

# 统计信息
stats = {
    "docx": 0, "doc": 0, "xlsx": 0, "xls": 0,
    "pdf": 0, "txt": 0, "image": 0, "wpsonline": 0,
    "skipped": 0, "failed": 0, "total": 0
}

# 日志文件
log_entries = []

def log(message, level="INFO"):
    """记录日志"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{timestamp}] [{level}] {message}"
    log_entries.append(entry)
    print(entry)

def generate_frontmatter(src_path):
    """生成 YAML frontmatter"""
    rel_src = os.path.relpath(src_path, SOURCE_ROOT)
    return f"""---
source_file: "{rel_src}"
converted_date: "{datetime.now().strftime('%Y-%m-%d')}"
---

"""

# ==================== 转换函数 ====================

def convert_docx_to_md(src_path, dst_path):
    """使用 pandoc 转换 docx 为 md（首选）"""
    try:
        # 提取图片到 media 子目录
        media_dir = os.path.join(os.path.dirname(dst_path), "media")
        os.makedirs(media_dir, exist_ok=True)
        
        result = subprocess.run(
            ["pandoc", "-f", "docx", "-t", "markdown", "--wrap=none",
             "--extract-media", media_dir,
             str(src_path)],
            capture_output=True, text=True, encoding="utf-8", timeout=60
        )
        
        if result.returncode == 0 and result.stdout.strip():
            frontmatter = generate_frontmatter(src_path)
            with open(dst_path, "w", encoding="utf-8") as f:
                f.write(frontmatter + result.stdout)
            stats["docx"] += 1
            return True
        else:
            log(f"pandoc 转换失败: {result.stderr}", "WARN")
            return False
    except Exception as e:
        log(f"pandoc 异常: {str(e)}", "ERROR")
        return False

def convert_doc_to_md(src_path, dst_path):
    """使用 Word COM 转换 doc 为 md"""
    try:
        import win32com.client
        import pythoncom
        
        pythoncom.CoInitialize()
        word = None
        doc = None
        temp_docx = None
        
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            doc = word.Documents.Open(os.path.abspath(src_path))
            temp_docx = os.path.join(tempfile.gettempdir(), f"convert_{int(time.time())}.docx")
            doc.SaveAs2(temp_docx, FileFormat=16)  # wdFormatXMLDocument
            doc.Close(False)
            doc = None
            word.Quit()
            word = None
            
            success = convert_docx_to_md(temp_docx, dst_path)
            if success:
                stats["doc"] += 1
            return success
        except Exception as e:
            log(f".doc 转换失败: {str(e)}", "ERROR")
            return False
        finally:
            try:
                if doc:
                    doc.Close(False)
            except:
                pass
            try:
                if word:
                    word.Quit()
            except:
                pass
            pythoncom.CoUninitialize()
            if temp_docx and os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except:
                    pass
    except ImportError:
        log("pywin32 未安装，无法转换 .doc 文件", "ERROR")
        return False

def convert_excel_to_csv(src_path, dst_dir, base_name, engine):
    """转换 Excel 为 CSV"""
    try:
        import pandas as pd
        
        xls = pd.ExcelFile(src_path, engine=engine)
        sheet_count = len(xls.sheet_names)
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # 多 sheet 时添加后缀
            suffix = f"_{sheet_name}" if sheet_count > 1 else ""
            safe_suffix = suffix.replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_")
            
            csv_path = os.path.join(dst_dir, f"{base_name}{safe_suffix}.csv")
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        
        if ".xlsx" in src_path.lower():
            stats["xlsx"] += 1
        else:
            stats["xls"] += 1
        return True
    except Exception as e:
        log(f"Excel 转换失败: {str(e)}", "ERROR")
        return False

def convert_pdf_to_md(src_path, dst_path):
    """转换 PDF 为 md"""
    try:
        import pdfplumber
        
        parts = []
        with pdfplumber.open(src_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    parts.append(text)
                if i < len(pdf.pages) - 1:
                    parts.append("\n---\n")
        
        frontmatter = generate_frontmatter(src_path)
        with open(dst_path, "w", encoding="utf-8") as f:
            f.write(frontmatter + "\n".join(parts))
        
        stats["pdf"] += 1
        return True
    except Exception as e:
        log(f"PDF 转换失败: {str(e)}", "ERROR")
        return False

def convert_txt_to_md(src_path, dst_path):
    """转换 txt 为 md"""
    try:
        with open(src_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        frontmatter = generate_frontmatter(src_path)
        with open(dst_path, "w", encoding="utf-8") as f:
            f.write(frontmatter + content)
        
        stats["txt"] += 1
        return True
    except Exception as e:
        log(f"TXT 转换失败: {str(e)}", "ERROR")
        return False

# ==================== 主流程 ====================

def main():
    log("=" * 80)
    log("美食知识库批量转换开始")
    log(f"源目录: {SOURCE_ROOT}")
    log(f"目标目录: {TARGET_ROOT}")
    log("=" * 80)
    
    # 检查源目录是否存在
    if not os.path.exists(SOURCE_ROOT):
        log(f"源目录不存在: {SOURCE_ROOT}", "ERROR")
        return
    
    # 遍历所有文件
    for root, dirs, files in os.walk(SOURCE_ROOT):
        dirs.sort()  # 保持目录顺序
        
        for filename in sorted(files):
            src_path = os.path.join(root, filename)
            rel_path = os.path.relpath(src_path, SOURCE_ROOT)
            rel_dir = os.path.dirname(rel_path)
            
            # 计算目标路径
            dst_dir = os.path.join(TARGET_ROOT, rel_dir) if rel_dir else TARGET_ROOT
            os.makedirs(dst_dir, exist_ok=True)
            
            # 跳过临时文件
            if filename.startswith("~$"):
                stats["skipped"] += 1
                continue
            
            ext = os.path.splitext(filename)[1].lower()
            base_name = os.path.splitext(filename)[0]
            stats["total"] += 1
            
            log(f"处理: {rel_path}")
            
            # 根据扩展名选择转换方式
            if ext == ".docx":
                dst_path = os.path.join(dst_dir, base_name + ".md")
                if convert_docx_to_md(src_path, dst_path):
                    log(f"✅ .docx → .md: {base_name}.md")
                else:
                    stats["failed"] += 1
                    log(f"❌ .docx 转换失败: {base_name}", "ERROR")
            
            elif ext == ".doc":
                dst_path = os.path.join(dst_dir, base_name + ".md")
                if convert_doc_to_md(src_path, dst_path):
                    log(f"✅ .doc → .md: {base_name}.md")
                else:
                    stats["failed"] += 1
                    log(f"❌ .doc 转换失败: {base_name}", "ERROR")
            
            elif ext == ".xlsx":
                if convert_excel_to_csv(src_path, dst_dir, base_name, "openpyxl"):
                    log(f"✅ .xlsx → .csv: {base_name}.csv")
                else:
                    stats["failed"] += 1
                    log(f"❌ .xlsx 转换失败: {base_name}", "ERROR")
            
            elif ext == ".xls":
                if convert_excel_to_csv(src_path, dst_dir, base_name, "xlrd"):
                    log(f"✅ .xls → .csv: {base_name}.csv")
                else:
                    stats["failed"] += 1
                    log(f"❌ .xls 转换失败: {base_name}", "ERROR")
            
            elif ext == ".pdf":
                dst_path = os.path.join(dst_dir, base_name + ".md")
                if convert_pdf_to_md(src_path, dst_path):
                    log(f"✅ .pdf → .md: {base_name}.md")
                else:
                    stats["failed"] += 1
                    log(f"❌ .pdf 转换失败: {base_name}", "ERROR")
            
            elif ext == ".txt":
                dst_path = os.path.join(dst_dir, base_name + ".md")
                if convert_txt_to_md(src_path, dst_path):
                    log(f"✅ .txt → .md: {base_name}.md")
                else:
                    stats["failed"] += 1
                    log(f"❌ .txt 转换失败: {base_name}", "ERROR")
            
            elif ext in (".jpg", ".jpeg", ".png", ".gif", ".bmp"):
                dst_path = os.path.join(dst_dir, filename)
                shutil.copy2(src_path, dst_path)
                stats["image"] += 1
                log(f"✅ 图片复制: {filename}")
            
            elif filename.endswith(".pof.wpsonline"):
                # WPS 云文档占位符
                placeholder = base_name[:-4] if base_name.endswith(".pof") else base_name
                dst_path = os.path.join(dst_dir, placeholder + ".md")
                with open(dst_path, "w", encoding="utf-8") as f:
                    f.write(f"# {placeholder}\n\n> ⚠️ WPS 云文档占位符\n> \n> 原始文件: `{rel_path}`\n> \n> 请从 WPS 云端手动补充内容到此文件。\n")
                stats["wpsonline"] += 1
                log(f"⚠️ WPS 云文档占位符: {placeholder}.md")
            
            else:
                stats["skipped"] += 1
                log(f"⏭️ 跳过不支持的格式: {ext}")
    
    # 输出统计报告
    print("\n" + "=" * 80)
    log("转换完成！统计报告:")
    log("=" * 80)
    log(f"总文件数:     {stats['total']}")
    log(f".docx → .md:  {stats['docx']}")
    log(f".doc → .md:   {stats['doc']}")
    log(f".xlsx → .csv: {stats['xlsx']}")
    log(f".xls → .csv:  {stats['xls']}")
    log(f".pdf → .md:   {stats['pdf']}")
    log(f".txt → .md:   {stats['txt']}")
    log(f"图片复制:     {stats['image']}")
    log(f"WPS 云文档:   {stats['wpsonline']}")
    log(f"跳过:         {stats['skipped']}")
    log(f"失败:         {stats['failed']}")
    log("=" * 80)
    
    # 保存日志文件
    log_file = os.path.join(TARGET_ROOT, "conversion_log.md")
    with open(log_file, "w", encoding="utf-8") as f:
        f.write("# 转换日志\n\n")
        f.write(f"**转换时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write("## 统计摘要\n\n")
        f.write(f"- 总文件数: {stats['total']}\n")
        f.write(f"- 成功转换: {stats['docx'] + stats['doc'] + stats['xlsx'] + stats['xls'] + stats['pdf'] + stats['txt'] + stats['image'] + stats['wpsonline']}\n")
        f.write(f"- 失败: {stats['failed']}\n")
        f.write(f"- 跳过: {stats['skipped']}\n\n")
        f.write("## 详细日志\n\n```\n")
        f.write("\n".join(log_entries))
        f.write("\n```\n")
    
    log(f"\n日志已保存到: {log_file}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log("\n用户中断转换", "WARN")
    except Exception as e:
        log(f"程序异常: {str(e)}", "ERROR")
        raise
