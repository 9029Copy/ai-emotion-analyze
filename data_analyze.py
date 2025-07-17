import pandas as pd
from openpyxl import load_workbook
import argparse, httpx, sys, json, os
import json
import io

CONFIG_FILE = "config.json"
with open(CONFIG_FILE, encoding="utf-8") as f:
    cfg = json.load(f)

def chat_once(q, key, url):
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "text/plain"}
    r = httpx.post(f"{url}/chat", content=q, headers=headers, timeout=60)
    r.raise_for_status()
    return r.text

try:
    p = argparse.ArgumentParser()
    p.add_argument("-k", "--key", required=True)
    p.add_argument("--host", default=cfg["host"])
    p.add_argument("--port", type=int, default=cfg["port"])
    p.add_argument("--scheme", default=cfg["scheme"])
    args = p.parse_args()

    # 读取comment.xlsx文件，A列数据从第2行开始（索引为1）
    df_original = pd.read_excel('input.xlsx', usecols='A', header=None, skiprows=1)
    comment_list = df_original[0].tolist()
    input_str = json.dumps(comment_list, ensure_ascii=False)

    question = "接下来我会给你json格式的一组文本，内容是某件商品的评论，请根据其内容进行情感分类，分类结果为下列之一：正面、负面、中性，然后输出相同格式的json文本给我（不输出其他任何内容）。字符串内容如下：" + input_str
    # question = "hello"
    answer_str = chat_once(question, args.key, f"{args.scheme}://{args.host}:{args.port}")
    # print("[debug]回答：" + answer_str)


    df_processed = pd.read_json(io.StringIO(answer_str))

    # 设置列名
    df_original.columns = ['商品评论']
    df_processed.columns = ['情感分类']

    # 合并两列
    merged_df = pd.concat([df_original, df_processed], axis=1)
    merged_df.to_excel("output.xlsx", index=False, engine='openpyxl')

    # 设置第一列列宽
    wb = load_workbook("output.xlsx")
    ws = wb.active
    ws.column_dimensions['A'].width = 50  # 设置第一列(A列)宽度为50
    wb.save("output.xlsx")
    wb.close()

except Exception as e:
    print(f"发生错误：{str(e)}")
