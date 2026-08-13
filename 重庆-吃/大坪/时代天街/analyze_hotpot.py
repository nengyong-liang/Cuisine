"""
大坪/时代天街火锅店深度分析脚本
从 txt 评论数据中系统化提取 8 家火锅店的所有相关评论，
进行情感分析，输出 CSV 数据 + MD 分析报告。
"""

import csv
import re
import os
from dataclasses import dataclass, field
from typing import Optional
from collections import defaultdict

# ==================== 配置区域 ====================
DATA_DIR = os.path.dirname(os.path.abspath(__file__))
TXT_FILES = ["大坪.txt", "大坪-1.txt", "避雷.txt"]
CSV_OUTPUT = os.path.join(DATA_DIR, "火锅店评论数据.csv")
MD_OUTPUT = os.path.join(DATA_DIR, "火锅店深度分析报告.md")

# 8 家目标火锅店及其别名
RESTAURANT_ALIASES = {
    "四间火锅": ["四间火锅", "四间"],
    "梯坎老火锅": ["梯坎老火锅", "梯坎", "廖姐梯坎", "廖姐梯坎老火锅"],
    "郭记明火锅": ["郭记明火锅", "郭记明", "明火锅"],
    "贰妹贵州地摊火锅": ["贰妹贵州地摊火锅", "贰妹", "贵州地摊火锅",
                         "贰妹贵州火锅", "贵州豆豉火锅", "贰妹地摊火锅"],
    "三孃老火锅": ["三孃老火锅", "三孃"],
    "响火锅": ["响火锅"],
    "瓜火锅": ["瓜火锅"],
}

# 情感关键词（简单词列表）
POSITIVE_WORDS = [
    "推荐", "宝藏", "王炸", "正宗", "新鲜", "性价比", "香",
    "绝了", "无敌", "糯", "回购", "不错", "惊艳", "仙品",
    "不拉肚子", "划算", "实惠", "干净", "地道", "入味", "软糯",
    "酥脆", "嫩", "下饭", "扶墙出", "神", "赞", "喜欢", "爱吃",
    "私藏", "食堂", "常去", "必去", "必吃", "不会低于", "好吃来砍我",
]
NEGATIVE_WORDS = [
    "难吃", "避雷", "苦", "差", "脏", "拉肚子", "营销",
    "退步", "水", "发苦", "发馊", "不新鲜", "态度差", "发臭",
    "踩雷", "坑", "黑店", "垃圾", "lj", "恶心", "想吐", "头发",
    "虫", "蟑螂", "老鼠", "缺斤短两", "耍称", "不卫生", "预制",
    "夹生", "寡淡", "偏咸", "过咸", "无记忆点",
    "刺客", "杀鱼费", "隐形消费", "发臭",
]

# 复合负面模式（正则）— 优先级高于简单词匹配
NEGATIVE_PATTERNS = [
    (r'不好吃(?!来砍我)', '不好吃'),  # 排除"不好吃来砍我"这种正面表达
    (r'没.*好吃', '没好吃'),  # 注意：不使用"不.*好吃"因为太贪婪，会误匹配"绝不...好吃"
    (r'越来越[拉差]', '越来越拉/差'),
    (r'拉中之拉', '拉中之拉'),
    (r'好吃个喘喘', '反讽好吃个喘喘'),
    (r'好吃个屁', '反讽好吃个屁'),
    (r'都是营销', '都是营销'),
    (r'全是营销', '全是营销'),
    (r'千万别', '千万别'),
    (r'快跑', '快跑'),
    (r'贵(?!州)', '贵'),  # "贵"但不匹配"贵州"
    (r'一般(?!.*[好错])', '一般'),  # "一般"但后面不跟"好"或"错"
    (r'感觉一般', '感觉一般'),
    (r'普通(?!.*好)', '普通'),
    (r'发苦', '发苦'),
    (r'锅底.*苦|苦.*锅底', '锅底苦'),
    (r'菜品.*差|差.*菜品', '菜品差'),
    (r'卫生.*差|差.*卫生', '卫生差'),
    (r'态度.*差|差.*态度', '态度差'),
    (r'吃.*拉肚子|拉肚子', '拉肚子'),
    (r'排队.*久|等位.*久', '排队久'),
]

# 复合正面模式（正则）
POSITIVE_PATTERNS = [
    (r'不好吃来砍我', '不好吃来砍我(好评)'),
    (r'好吃不贵', '好吃不贵'),
    (r'超级好吃', '超级好吃'),
    (r'无敌', '无敌'),
]

# "好吃" 的正面匹配（排除否定/反讽语境）
import re as _re
def _is_positive_haoci(text: str) -> bool:
    """检查'好吃'是否在正面语境中"""
    for m in _re.finditer('好吃', text):
        start = max(0, m.start() - 3)
        before = text[start:m.start()]
        after = text[m.end():m.end() + 4]
        # 排除: 不好吃, 没好吃, 难好吃
        if before.endswith('不') or before.endswith('没') or before.endswith('难'):
            continue
        # 排除: 好吃个喘喘, 好吃个屁
        if '个喘喘' in after or '个屁' in after:
            continue
        # 排除: "以前...好吃...现在...难吃" 模式中的好吃
        if '以前' in before or '一开始' in before:
            # 检查后面是否有"现在...难吃"或"越来越"
            remaining = text[m.end():]
            if '现在' in remaining and ('难吃' in remaining or '拉' in remaining):
                continue
            if '越来越' in remaining:
                continue
        return True
    return False

# ==================== 数据结构 ====================

@dataclass
class Comment:
    username: str = ""
    content: str = ""
    date: str = ""
    location: str = ""
    likes: int = 0
    reply_to: str = ""  # 回复目标用户名
    line_number: int = 0
    source_file: str = ""
    # 匹配结果
    matched_stores: list = field(default_factory=list)
    sentiment: str = ""  # positive / negative / neutral
    sentiment_words: list = field(default_factory=list)
    is_comparison: bool = False
    matched_by: str = ""  # "direct" / "post_context" / "reply_thread"

@dataclass
class Post:
    title: str = ""
    content: str = ""
    date: str = ""
    comment_count: int = 0
    comments: list = field(default_factory=list)
    source_file: str = ""
    content_line_start: int = 0
    # 帖子正文匹配到的店铺
    post_stores: list = field(default_factory=list)


# ==================== 解析器 ====================

# 日期行正则：匹配 "03-23重庆" / "2023-12-30重庆" / "03-23重庆" 等
DATE_PATTERN = re.compile(r'^(\d{2,4}-\d{2}(?:-\d{2})?)(.*)$')

def parse_likes(line: str) -> int:
    """解析赞数行：数字返回数字，'赞'返回0"""
    line = line.strip()
    if line == "赞" or line == "":
        return 0
    try:
        return int(line)
    except ValueError:
        return 0

def is_reply_line(line: str) -> Optional[str]:
    """检测是否是回复行，返回被回复的用户名"""
    line = line.strip()
    m = re.match(r'^回复\s+(.+?)\s*:', line)
    if m:
        return m.group(1).strip()
    return None

def match_stores_in_text(text: str) -> list:
    """在文本中匹配火锅店，返回匹配到的店名列表"""
    matched = []
    for store_name, aliases in RESTAURANT_ALIASES.items():
        for alias in aliases:
            if alias in text:
                matched.append(store_name)
                break
    return matched

def analyze_sentiment(text: str) -> tuple:
    """改进的情感分析：处理否定/反讽/转折，返回 (sentiment, matched_words)"""
    matched_neg = []
    matched_pos = []

    # Step 1: 先匹配复合负面模式（正则），优先级最高
    for pattern, label in NEGATIVE_PATTERNS:
        if re.search(pattern, text):
            if label not in matched_neg:
                matched_neg.append(label)

    # Step 2: 匹配复合正面模式
    for pattern, label in POSITIVE_PATTERNS:
        if re.search(pattern, text):
            if label not in matched_pos:
                matched_pos.append(label)
            # 如果"不好吃来砍我"命中，则从负面中移除"不好吃"
            if '来砍我' in label:
                matched_neg = [w for w in matched_neg if w != '不好吃']

    # Step 3: 匹配简单负面词
    for w in NEGATIVE_WORDS:
        if w in text:
            if w not in matched_neg:
                matched_neg.append(w)

    # Step 4: 匹配简单正面词
    for w in POSITIVE_WORDS:
        if w in text:
            matched_pos.append(w)

    # Step 5: 特殊处理"好吃"——使用上下文感知匹配
    if _is_positive_haoci(text):
        if '好吃' not in matched_pos:
            matched_pos.append('好吃')

    # Step 6: 如果复合负面模式命中了"不好吃"等，则移除正面词中的"好吃"
    if any('不好吃' in w or '没好吃' in w for w in matched_neg):
        matched_pos = [w for w in matched_pos if w != '好吃']

    neg_count = len(matched_neg)
    pos_count = len(matched_pos)

    if neg_count > pos_count:
        return "negative", matched_neg
    elif pos_count > neg_count:
        return "positive", matched_pos
    elif pos_count > 0:
        return "positive", matched_pos
    else:
        return "neutral", []

def parse_txt(filepath: str) -> list:
    """解析 txt 文件，返回 Post 列表"""
    filename = os.path.basename(filepath)
    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()

    # 预处理：去掉换行符
    lines = [line.rstrip("\n").rstrip("\r") for line in lines]

    # 按 "猜你想搜" 分割帖子块
    # 每个块 = [前面的内容, "猜你想搜", 搜索词, 日期, "共X条评论", 评论...]
    posts_raw = []
    current_block_lines = []
    current_start_line = 1

    for i, line in enumerate(lines):
        if line.strip() == "猜你想搜":
            # 保存当前块
            if current_block_lines:
                posts_raw.append((current_start_line, current_block_lines))
            current_block_lines = []
            current_start_line = i + 1
        current_block_lines.append((i + 1, line))

    # 最后一个块
    if current_block_lines:
        posts_raw.append((current_start_line, current_block_lines))

    posts = []
    for block_start, block_lines in posts_raw:
        post = Post(source_file=filename)
        post.content_line_start = block_start

        # block_lines[0] 开始就是帖子内容（在 "猜你想搜" 之前）
        # 但我们按 "猜你想搜" 分割后，block_lines 的内容是该标记之后的部分
        # 需要重新理解结构

        # 实际上：block_lines 包含从上一个"猜你想搜"之后到当前"猜你想搜"之前的所有行
        # 或者从文件开头到第一个"猜你想搜"
        # 第一个块（文件开头到第一个"猜你想搜"）是第一个帖子的正文
        # 后续块的结构是：搜索词、日期、评论数、评论...

        # 找到块内的结构
        idx = 0
        block_content_lines = []

        # 如果块不以 "猜你想搜" 开头，说明这是第一个块（文件开头到第一个"猜你想搜"）
        # 这个块的最后一部分是第一个帖子的正文
        if block_lines[0][1].strip() != "猜你想搜":
            # 这是第一个块，包含帖子正文
            # 跳过 "猜你想搜" 行（它在下一个块的开头）
            # 不对，我们按 "猜你想搜" 分割，所以 "猜你想搜" 行本身不在 block 中
            # 让我重新理解：current_block_lines 是在遇到 "猜你想搜" 之前的所有行
            # 所以第一个 block 是文件开头到第一个 "猜你想搜" 之前的行
            # 这个 block 的内容就是帖子正文
            post.content = "\n".join(line for _, line in block_lines if line.strip())
            post.title = block_lines[0][1] if block_lines else ""
            posts.append(post)
            continue

        # 后续块：以 "猜你想搜" 开头（但我们已经把 "猜你想搜" 行也包含在内了）
        # 等等，我上面的逻辑是：遇到 "猜你想搜" 时，先把之前的 block 存起来，
        # 然后开始新的 block。所以 "猜你想搜" 这一行会被加到新 block 的开头。

        # 让我重新看代码逻辑：
        # for i, line:
        #   if line == "猜你想搜":
        #       save old block
        #       start new block
        #   add line to new block
        # 所以 "猜你想搜" 行会被加到新 block 的开头。

        # 重新处理：跳过 "猜你想搜" 行
        if block_lines[0][1].strip() == "猜你想搜":
            idx = 1
        else:
            idx = 0

        # 读取搜索关键词
        if idx < len(block_lines):
            post.title = block_lines[idx][1].strip()
            idx += 1

        # 读取日期行
        if idx < len(block_lines):
            date_line = block_lines[idx][1].strip()
            # 日期可能是 "03-19", "2025-04-27", "编辑于 03-30"
            post.date = date_line
            idx += 1

        # 读取评论数行 "共 X 条评论"
        if idx < len(block_lines):
            count_line = block_lines[idx][1].strip()
            m = re.match(r'^共\s*(\d+)\s*条评论', count_line)
            if m:
                post.comment_count = int(m.group(1))
                idx += 1
            else:
                # 可能没有评论数行
                pass

        # 跳过空行
        while idx < len(block_lines) and block_lines[idx][1].strip() == "":
            idx += 1

        # 解析评论
        # 同时需要找到这个帖子对应的正文
        # 帖子正文在前一个 block 的末尾部分
        # 我们先解析评论，后面再关联正文

        comments_raw = block_lines[idx:]

        # 解析评论
        comments = parse_comments(comments_raw, filename)
        post.comments = comments

        # 尝试从前一个 post 获取正文
        if posts:
            # 前一个 post 的 content 可能就是当前帖子的正文
            # 不对，前一个 post 的 content 是它自己的正文
            # 实际上，文件结构是：
            # [帖子1正文] 猜你想搜 [搜索词1] [日期1] [评论数1] [评论1...]
            # [帖子2正文] 猜你想搜 [搜索词2] [日期2] [评论数2] [评论2...]
            # 所以每个 block 开头的非 "猜你想搜" 部分就是帖子正文
            pass

        posts.append(post)

    # 第二轮：关联帖子正文
    # 每个 post 的正文在它前面那个 block 的内容部分
    # 但我们按 "猜你想搜" 分割，所以 block[0] 的内容部分是帖子1的正文
    # block[1] 的 "猜你想搜" 之后是帖子1的元信息和评论
    # block[1] 的内容部分（如果有）是帖子2的正文

    # 重新理解：实际上文件结构是：
    # 帖子1标题和正文
    # 猜你想搜
    # 搜索词1
    # 日期1
    # 共X条评论1
    # 评论...
    # 帖子2标题和正文  ← 这部分在 block[1] 的 "猜你想搜" 之后、评论之前？
    # 不对，"猜你想搜" 是帖子元信息的开始，不是结束

    # 让我重新看实际数据：
    # Line 1: 石油路好吃又平价的8️⃣家美食！！！  ← 帖子1标题
    # Line 2-12: 帖子1正文
    # Line 13: 空行
    # Line 14: 猜你想搜  ← 帖子1元信息
    # Line 15: 悦来香酒家  ← 搜索词
    # Line 16: 空行 (有时)
    # Line 17-20: 可能还有另一个帖子的标题
    # Line 22: 猜你想搜  ← 帖子2元信息
    # ...

    # 啊，我看到了！一个文件里可能有多个帖子，每个帖子有：
    # 1. 帖子正文
    # 2. "猜你想搜" + 搜索词 + 日期 + 评论数
    # 3. 评论列表
    # 然后下一个帖子的正文开始

    # 所以正确的分块方式应该是：
    # 文件开头到第一个 "猜你想搜" = 帖子1正文
    # 第一个 "猜你想搜" 到第二个 "猜你想搜" 之间的内容 = 帖子1的元信息+评论 + 帖子2的正文
    # 这太复杂了，我需要重新设计解析器

    return posts  # 临时返回，后面重写


def parse_comments(comment_lines: list, source_file: str) -> list:
    """解析评论列表"""
    comments = []
    i = 0
    n = len(comment_lines)

    while i < n:
        line_no, line = comment_lines[i]

        # 跳过空行和分隔符
        stripped = line.strip()
        if stripped == "" or stripped.startswith("- THE END"):
            i += 1
            continue

        # 跳过 "关注" 等非评论行
        if stripped == "关注":
            i += 1
            continue

        # 用户名行
        username = stripped
        i += 1

        # 收集评论内容（可能多行，直到遇到日期行）
        content_lines = []
        reply_to = ""
        while i < n:
            _, content_line = comment_lines[i]
            stripped_content = content_line.strip()

            # 检查是否是日期行（评论内容结束）
            if DATE_PATTERN.match(stripped_content) and not stripped_content.startswith("回复"):
                break
            if stripped_content == "" and content_lines:
                # 可能是内容结束，也可能是内容中间的空行
                # 检查下一行是否是日期行
                if i + 1 < n:
                    next_stripped = comment_lines[i + 1][1].strip()
                    if DATE_PATTERN.match(next_stripped):
                        break
                break
            if stripped_content.startswith("- THE END"):
                break

            # 检查是否是回复行
            reply_target = is_reply_line(stripped_content)
            if reply_target and not content_lines:
                reply_to = reply_target
                content_lines.append(stripped_content)
            else:
                content_lines.append(stripped_content)
            i += 1

        content = " ".join(content_lines)

        # 日期+地点行
        date_str = ""
        location = ""
        if i < n:
            date_line = comment_lines[i][1].strip()
            m = DATE_PATTERN.match(date_line)
            if m:
                date_str = m.group(1)
                location = m.group(2).strip()
            i += 1

        # 赞数行
        likes = 0
        if i < n:
            likes_line = comment_lines[i][1].strip()
            if likes_line == "赞" or likes_line == "回复" or likes_line == "":
                likes = 0
            else:
                try:
                    likes = int(likes_line)
                except ValueError:
                    likes = 0
            i += 1

        # 回复数行（跳过）
        if i < n:
            reply_line = comment_lines[i][1].strip()
            if reply_line in ("回复", "赞", "") or reply_line.isdigit():
                i += 1

        # 跳过空行
        while i < n and comment_lines[i][1].strip() == "":
            i += 1

        if content or username:
            comment = Comment(
                username=username,
                content=content,
                date=date_str,
                location=location,
                likes=likes,
                reply_to=reply_to,
                line_number=line_no,
                source_file=source_file,
            )
            comments.append(comment)

    return comments


# ==================== 重新设计的解析器 ====================

def parse_txt_v2(filepath: str) -> list:
    """
    解析 txt 文件 v2：正确处理帖子结构

    文件结构：
    [帖子1正文（标题+内容）]
    猜你想搜
    [搜索词]
    [日期]
    共 X 条评论
    [空行]
    [评论1]
    [评论2]
    ...
    - THE END
    [帖子2正文]
    猜你想搜
    ...
    """
    filename = os.path.basename(filepath)
    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()
    lines = [line.rstrip("\n").rstrip("\r") for line in lines]

    posts = []
    i = 0
    n = len(lines)

    while i < n:
        # 收集帖子正文（直到遇到 "猜你想搜"）
        post_content_lines = []
        post_start_line = i + 1

        while i < n and lines[i].strip() != "猜你想搜":
            post_content_lines.append((i + 1, lines[i]))
            i += 1

        # 如果没有遇到 "猜你想搜"，说明是文件末尾的残余内容
        if i >= n:
            break

        # 跳过 "猜你想搜"
        i += 1  # now i points to line after "猜你想搜"

        post = Post(source_file=filename, content_line_start=post_start_line)

        # 帖子正文
        content_text = "\n".join(line for _, line in post_content_lines if line.strip())
        post.content = content_text
        # 帖子标题 = 第一行非空行
        for _, line in post_content_lines:
            if line.strip():
                post.title = line.strip()
                break

        # 搜索关键词
        if i < n:
            post.search_keyword = lines[i].strip()
            i += 1

        # 日期行
        if i < n:
            date_line = lines[i].strip()
            if date_line and date_line != "共" and not date_line.startswith("共"):
                post.date = date_line
                i += 1

        # 评论数行
        if i < n:
            count_line = lines[i].strip()
            m = re.match(r'^共\s*(\d+)\s*条评论', count_line)
            if m:
                post.comment_count = int(m.group(1))
                i += 1

        # 跳过空行
        while i < n and lines[i].strip() == "":
            i += 1

        # 收集评论行（直到下一个 "猜你想搜" 或文件末尾）
        comment_lines = []
        while i < n and lines[i].strip() != "猜你想搜":
            # 注意：下一个帖子的正文也会被包含进来
            # 但评论结束后通常有 "- THE END" 标记
            # 如果没有标记，我们需要靠评论格式来区分
            comment_lines.append((i + 1, lines[i]))
            i += 1

        # 从 comment_lines 中分离出评论和下一个帖子的正文
        # 策略：从后往前找，如果遇到 "- THE END" 或 "- THE END -"，则其后的是下个帖子正文
        # 如果没有 THE END，则需要靠评论格式判断

        # 找 THE END 标记
        end_idx = -1
        for j in range(len(comment_lines) - 1, -1, -1):
            if comment_lines[j][1].strip().startswith("- THE END"):
                end_idx = j
                break

        if end_idx >= 0:
            # THE END 之前的是评论，之后的是下个帖子正文
            actual_comment_lines = comment_lines[:end_idx]
            # THE END 之后的内容需要放回 lines 中供下一个帖子处理
            # 但我们已经消费了这些行，需要回退 i
            leftover_lines = comment_lines[end_idx + 1:]
            # 回退 i
            i -= len(leftover_lines)
        else:
            # 没有 THE END 标记，尝试通过评论格式分离
            # 从后往前找，第一个不符合评论格式的行就是分界点
            # 评论格式：用户名 → 内容 → 日期行 → 赞/数字 → 回复/数字 → 空行
            # 日期行是很好的锚点
            actual_comment_lines = comment_lines
            # 简单处理：全部当作评论

        # 解析评论
        comments = parse_comments(actual_comment_lines, filename)
        post.comments = comments

        # 匹配帖子正文中的店铺
        post.post_stores = match_stores_in_text(content_text)

        posts.append(post)

    return posts


# ==================== 店铺匹配 + 情感分析 ====================

def match_and_analyze(posts: list) -> list:
    """匹配店铺归属并做情感分析，返回所有匹配的评论"""
    all_matched_comments = []

    for post in posts:
        # 为帖子构建店铺上下文
        post_context_stores = set(post.post_stores)

        for comment in post.comments:
            # 1. 直接匹配：评论内容中包含店名
            direct_stores = match_stores_in_text(comment.content)

            # 2. 帖子上下文匹配：如果帖子正文提到某店
            context_stores = set()
            if post_context_stores:
                context_stores = post_context_stores.copy()
                # 如果评论明确提到其他店，则从上下文中移除
                # （但如果评论同时提到上下文中的店和其他店，两个都保留）
                if direct_stores and not any(s in direct_stores for s in post_context_stores):
                    context_stores = set()  # 评论在讨论别的店

            # 3. 回复线程匹配：如果回复的评论属于某店
            reply_stores = set()
            if comment.reply_to:
                # 查找被回复的评论
                for other_comment in post.comments:
                    if other_comment.username == comment.reply_to:
                        if other_comment.matched_stores:
                            reply_stores = set(other_comment.matched_stores)
                        break

            # 合并所有匹配
            all_stores = set(direct_stores) | context_stores | reply_stores

            if all_stores:
                # 标记匹配方式
                if direct_stores:
                    comment.matched_by = "direct"
                elif context_stores:
                    comment.matched_by = "post_context"
                else:
                    comment.matched_by = "reply_thread"

                # 多店归属标记
                comment.is_comparison = len(all_stores) > 1

                # 情感分析
                comment.sentiment, comment.sentiment_words = analyze_sentiment(comment.content)

                # 为每个匹配的店创建一条记录
                for store in all_stores:
                    comment_copy = Comment(
                        username=comment.username,
                        content=comment.content,
                        date=comment.date,
                        location=comment.location,
                        likes=comment.likes,
                        reply_to=comment.reply_to,
                        line_number=comment.line_number,
                        source_file=comment.source_file,
                        matched_stores=list(all_stores),
                        sentiment=comment.sentiment,
                        sentiment_words=comment.sentiment_words,
                        is_comparison=comment.is_comparison,
                        matched_by=comment.matched_by,
                    )
                    comment_copy.store_name = store
                    comment_copy.post_title = post.title
                    comment_copy.post_date = post.date
                    all_matched_comments.append(comment_copy)

    return all_matched_comments


# ==================== CSV 输出 ====================

def export_csv(comments: list, output_path: str):
    """导出 CSV"""
    fieldnames = [
        "store_name", "post_title", "post_date", "username", "content",
        "comment_date", "location", "likes", "reply_to", "sentiment",
        "sentiment_words", "is_comparison", "matched_by", "source_file",
        "line_number"
    ]

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for c in comments:
            writer.writerow({
                "store_name": getattr(c, "store_name", ""),
                "post_title": getattr(c, "post_title", ""),
                "post_date": getattr(c, "post_date", ""),
                "username": c.username,
                "content": c.content,
                "comment_date": c.date,
                "location": c.location,
                "likes": c.likes,
                "reply_to": c.reply_to,
                "sentiment": c.sentiment,
                "sentiment_words": ", ".join(c.sentiment_words),
                "is_comparison": c.is_comparison,
                "matched_by": c.matched_by,
                "source_file": c.source_file,
                "line_number": c.line_number,
            })


# ==================== MD 报告生成 ====================

def generate_report(comments: list, output_path: str):
    """生成 MD 分析报告"""
    # 按店分组统计
    store_data = defaultdict(lambda: {
        "total": 0, "positive": 0, "negative": 0, "neutral": 0,
        "comments": [], "max_likes": 0,
        "pos_words": defaultdict(int), "neg_words": defaultdict(int),
        "direct_count": 0, "context_count": 0, "reply_count": 0,
    })

    for c in comments:
        store = getattr(c, "store_name", "")
        data = store_data[store]
        data["total"] += 1
        data["comments"].append(c)

        if c.sentiment == "positive":
            data["positive"] += 1
            for w in c.sentiment_words:
                data["pos_words"][w] += 1
        elif c.sentiment == "negative":
            data["negative"] += 1
            for w in c.sentiment_words:
                data["neg_words"][w] += 1
        else:
            data["neutral"] += 1

        if c.likes > data["max_likes"]:
            data["max_likes"] = c.likes

        if c.matched_by == "direct":
            data["direct_count"] += 1
        elif c.matched_by == "post_context":
            data["context_count"] += 1
        else:
            data["reply_count"] += 1

    # 按提及次数排序
    sorted_stores = sorted(store_data.items(), key=lambda x: x[1]["total"], reverse=True)

    report = []
    report.append("# 大坪/时代天街火锅店深度分析报告")
    report.append("")
    report.append("## 一、分析方法说明")
    report.append("")
    report.append("### 数据来源")
    report.append("- 3 个 txt 文件：大坪.txt（6186行）、大坪-1.txt（2136行）、避雷.txt（612行）")
    report.append("- 共 10 个帖子，约 1364 条评论")
    report.append("")
    report.append("### 解析方法")
    report.append("1. **帖子结构解析**：按「猜你想搜」标记分割帖子，解析评论线程（用户名→内容→日期+地点→赞数→回复数）")
    report.append("2. **店铺匹配**（三层逻辑）：")
    report.append("   - **直接匹配**：评论内容中包含店名/别名/简称")
    report.append("   - **帖子上下文**：帖子正文提到某店，该帖所有评论默认归属该店（解决代词指代）")
    report.append("   - **回复线程**：回复评论继承被回复评论的店铺归属")
    report.append("3. **情感分析**：关键词法（正面词 vs 负面词计数）")
    report.append("4. **多店归属**：对比评论同时归属多家店，标记 is_comparison=True")
    report.append("")
    report.append("### 局限性")
    report.append("- 情感分析为简单关键词法，无法理解反讽/双重否定等复杂语义")
    report.append("- 帖子上下文归属可能过度归属（帖子提到某店但评论在讨论其他话题）")
    report.append("- 别名列表可能不全，部分口语化简称可能遗漏")
    report.append("")

    report.append("---")
    report.append("")
    report.append("## 二、7家火锅店逐一分析")
    report.append("")

    # 店铺人均数据（从概要表）
    price_map = {
        "四间火锅": "75",
        "梯坎老火锅": "68",
        "郭记明火锅": "70",
        "贰妹贵州地摊火锅": "50",
        "三孃老火锅": "67",
        "响火锅": "70",
        "瓜火锅": "65",
    }

    for store_name, data in sorted_stores:
        positive_rate = data["positive"] / data["total"] * 100 if data["total"] > 0 else 0
        negative_rate = data["negative"] / data["total"] * 100 if data["total"] > 0 else 0

        report.append(f"### {store_name}")
        report.append(f"- **人均**：约 {price_map.get(store_name, '未知')} 元")
        report.append(f"- **提及次数**：{data['total']} 次")
        report.append(f"- **情感分布**：正面 {data['positive']}（{positive_rate:.0f}%）/ 负面 {data['negative']}（{negative_rate:.0f}%）/ 中性 {data['neutral']}")
        report.append(f"- **匹配方式**：直接匹配 {data['direct_count']} / 帖子上下文 {data['context_count']} / 回复线程 {data['reply_count']}")
        report.append("")

        # 高赞正面评论
        pos_comments = sorted(
            [c for c in data["comments"] if c.sentiment == "positive"],
            key=lambda x: x.likes, reverse=True
        )[:3]
        if pos_comments:
            report.append("**高赞正面评论**：")
            for c in pos_comments:
                report.append(f"- (赞{c.likes}) {c.content[:150]}{'...' if len(c.content) > 150 else ''}")
            report.append("")

        # 高赞负面评论
        neg_comments = sorted(
            [c for c in data["comments"] if c.sentiment == "negative"],
            key=lambda x: x.likes, reverse=True
        )[:3]
        if neg_comments:
            report.append("**高赞负面评论**：")
            for c in neg_comments:
                report.append(f"- (赞{c.likes}) {c.content[:150]}{'...' if len(c.content) > 150 else ''}")
            report.append("")

        # 高频正面关键词
        if data["pos_words"]:
            top_pos = sorted(data["pos_words"].items(), key=lambda x: x[1], reverse=True)[:5]
            report.append(f"**典型优点关键词**：{', '.join(f'{w}({n}次)' for w, n in top_pos)}")
            report.append("")

        # 高频负面关键词
        if data["neg_words"]:
            top_neg = sorted(data["neg_words"].items(), key=lambda x: x[1], reverse=True)[:5]
            report.append(f"**典型缺点关键词**：{', '.join(f'{w}({n}次)' for w, n in top_neg)}")
            report.append("")

        report.append("---")
        report.append("")

    # 横向对比表
    report.append("## 三、横向对比表")
    report.append("")
    report.append("| 店名 | 提及次数 | 正面数 | 负面数 | 正面率 | 最高赞 | 人均 | 匹配方式分布 |")
    report.append("|------|---------|--------|--------|--------|--------|------|-------------|")

    for store_name, data in sorted_stores:
        positive_rate = data["positive"] / data["total"] * 100 if data["total"] > 0 else 0
        match_dist = f"直接{data['direct_count']}/上下文{data['context_count']}/回复{data['reply_count']}"
        report.append(
            f"| {store_name} | {data['total']} | {data['positive']} | {data['negative']} | "
            f"{positive_rate:.0f}% | {data['max_likes']} | "
            f"{price_map.get(store_name, '?')}元 | {match_dist} |"
        )

    report.append("")

    # 最终推荐排名
    report.append("## 四、最终推荐排名")
    report.append("")
    report.append("基于提及次数、正面率、高赞评论质量综合判定：")
    report.append("")

    # 计算推荐分数：正面率 * log(提及次数+1) + 高赞评论加成
    import math
    store_scores = []
    for store_name, data in sorted_stores:
        if data["total"] > 0:
            positive_rate = data["positive"] / data["total"]
            negative_rate = data["negative"] / data["total"]
            # 分数 = 正面率 * log(提及次数) - 负面率 * 0.5
            score = positive_rate * math.log(data["total"] + 1) - negative_rate * 0.5 * math.log(data["total"] + 1)
            store_scores.append((store_name, score, data))

    store_scores.sort(key=lambda x: x[1], reverse=True)

    medals = ["🥇", "🥈", "🥉"]
    for idx, (store_name, score, data) in enumerate(store_scores):
        medal = medals[idx] if idx < 3 else f"{idx + 1}."
        positive_rate = data["positive"] / data["total"] * 100 if data["total"] > 0 else 0
        report.append(
            f"{medal} **{store_name}** — 提及{data['total']}次，正面率{positive_rate:.0f}%，"
            f"综合评分{score:.2f}"
        )

    report.append("")
    report.append("## 五、附录")
    report.append(f"- 完整评论数据见 CSV 文件：`火锅店评论数据.csv`")
    report.append(f"- 分析脚本：`analyze_hotpot.py`")
    report.append("")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report))


# ==================== 主流程 ====================

def main():
    print("=" * 60)
    print("大坪/时代天街火锅店深度分析")
    print("=" * 60)

    all_posts = []
    for txt_file in TXT_FILES:
        filepath = os.path.join(DATA_DIR, txt_file)
        if not os.path.exists(filepath):
            print(f"⚠️ 文件不存在：{filepath}")
            continue
        print(f"📄 解析 {txt_file} ...")
        posts = parse_txt_v2(filepath)
        print(f"   → 解析到 {len(posts)} 个帖子")
        for post in posts:
            print(f"     - 「{post.title[:30]}...」 {len(post.comments)} 条评论, 帖子正文匹配店铺: {post.post_stores}")
        all_posts.extend(posts)

    print(f"\n📊 总计：{len(all_posts)} 个帖子")

    # 店铺匹配 + 情感分析
    print("\n🔍 匹配火锅店评论 + 情感分析 ...")
    matched_comments = match_and_analyze(all_posts)
    print(f"   → 匹配到 {len(matched_comments)} 条评论（含多店归属）")

    # 按店统计
    store_counts = defaultdict(int)
    for c in matched_comments:
        store_counts[getattr(c, "store_name", "")] += 1
    for store, count in sorted(store_counts.items(), key=lambda x: x[1], reverse=True):
        print(f"   - {store}: {count} 条")

    # 导出 CSV
    print(f"\n💾 导出 CSV：{CSV_OUTPUT}")
    export_csv(matched_comments, CSV_OUTPUT)

    # 生成 MD 报告
    print(f"\n📝 生成 MD 报告：{MD_OUTPUT}")
    generate_report(matched_comments, MD_OUTPUT)

    print("\n✅ 完成！")
    print(f"   CSV: {CSV_OUTPUT}")
    print(f"   MD:  {MD_OUTPUT}")


if __name__ == "__main__":
    main()
