import pandas as pd
from deep_translator import GoogleTranslator
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm  # 用于显示进度条，需要安装：pip install tqdm

# --- 文件和列设置 (保持不变) ---
file_name = "test.xlsx"
column_name = '商品名称'
OUTPUT_FILE = "翻译后的_商品数据_加速版.xlsx"
MAX_THREADS = 10  # 可以根据您的网络情况调整，通常 5 到 20 效果最好

# 1. 读取文件
try:
    df = pd.read_excel(file_name)
except Exception as e:
    print(f"错误: 无法读取文件 {file_name}。{e}")
    exit()
if column_name not in df.columns:
    print(f"错误: 缺少 '{column_name}' 列。")
    exit()

# 2. 初始化翻译器
try:
    translator = GoogleTranslator(source='vi', target='zh-CN')
except Exception as e:
    print(f"错误: 初始化翻译服务失败。{e}")
    exit()


# 3. 定义多线程翻译函数
def translate_batch(texts):
    """接收一批文本，使用 GoogleTranslator 批量翻译"""
    # 过滤空值/非字符串，并只保留需要翻译的文本
    texts_to_translate = [str(t) for t in texts if pd.notna(t) and str(t).strip()]

    if not texts_to_translate:
        return [""] * len(texts)  # 如果批次为空，返回空列表

    try:
        # **关键加速点：一次性发送批量请求**
        translated_list = translator.translate_batch(texts_to_translate)

        # 结果与原始输入对齐（简化处理，如果翻译失败，这里可能需要更复杂的逻辑）
        translated_results = []
        translated_index = 0
        for original_text in texts:
            if pd.notna(original_text) and str(original_text).strip():
                # 对应已翻译的结果
                translated_results.append(translated_list[translated_index])
                translated_index += 1
            else:
                # 原始空值
                translated_results.append("")
        return translated_results

    except Exception as e:
        print(f"警告: 批量翻译时发生错误 ({e})。该批次将返回空值。")
        return [""] * len(texts)  # 失败时返回空字符串列表


# 4. 执行加速翻译
print(f"\n--- 开始使用 {MAX_THREADS} 个线程加速翻译 {len(df)} 条数据 ---")
CHUNK_SIZE = 100  # 每个线程处理 100 条数据

all_translated = []
data_chunks = [df[column_name].iloc[i:i + CHUNK_SIZE] for i in range(0, len(df), CHUNK_SIZE)]

# 使用 ThreadPoolExecutor 并行处理数据块
with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
    # 使用 tqdm 包装，显示进度条
    # executor.map 会并行地将 translate_batch 应用于 data_chunks 中的每个批次
    translated_chunks = list(tqdm(executor.map(translate_batch, data_chunks),
                                  total=len(data_chunks),
                                  desc="翻译进度"))

# 5. 组合结果
for chunk in translated_chunks:
    all_translated.extend(chunk)

# 确保结果长度一致
if len(all_translated) == len(df):
    df['商品名称 (中文翻译)'] = all_translated
else:
    print("警告: 翻译结果数量与原始数据不匹配，请检查代码逻辑。")
    exit()

# 6. 导出结果
df.to_excel(OUTPUT_FILE, index=False)

print("\n--- 翻译加速完成 ---")
print(f"完整数据（包含新的中文翻译列）已保存到文件：'{OUTPUT_FILE}'")
print(df[[column_name, '商品名称 (中文翻译)']].head())