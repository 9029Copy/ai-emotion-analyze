import pandas as pd
from openpyxl import load_workbook
import argparse, httpx, sys, json, os
import json
import io

CONFIG_FILE = "config.json"
with open(CONFIG_FILE, encoding="utf-8") as f:
    cfg = json.load(f)

HOST = cfg.get("host", "127.0.0.1")
PORT = cfg.get("port", 8000)
SCHEME = cfg.get("scheme", "http")
MODEL = cfg.get("model", "THUDM/GLM-4-9B-0414")

def chat_once(q, model, key, url):
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json"
    }
    data = {
        "question": q,
        "model": model
    }
    try:
        r = httpx.post(f"{url}/chat", json=data, headers=headers, timeout=60)
        r.raise_for_status()
    except httpx.HTTPStatusError as e:
        if(e.response.status_code == 401):
            print("密钥错误。请检查你输入的密钥是否正确。")
        else:
            print(f"HTTP服务异常: {e}\n请检查你输入的模型名是否正确。")
        sys.exit(1)
    except httpx.HTTPError as e:
        print(f"HTTP请求异常: {e}\n请检查你输入的协议、主机名、端口号是否正确，或检查你的网络连接。")
        sys.exit(1)
        
    return r.text

def main():
    try:
        p = argparse.ArgumentParser()
        p.add_argument("-k", "--key", required=True)
        p.add_argument("-m", "--model", default=MODEL)
        p.add_argument("--host", default=HOST)
        p.add_argument("--port", type=int, default=PORT)
        p.add_argument("--scheme", default=SCHEME)
        args = p.parse_args()

        # 读取comment.xlsx文件，A列数据从第2行开始（索引为1）
        df_original = pd.read_excel('input.xlsx', usecols='A', header=None, skiprows=1)
        comment_list = df_original[0].tolist()

        # 初始化空列表
        sentiment_list = []
        token_list = []

        i = 1
        for text in comment_list:
            text = text.replace("\n", "")
            question = f"请对以下商品评论进行情感分类（正面/负面/中性，其中既表现出正面也表现出负面为中性，不输出其他任何内容）：{text}"

            # 调用chat_once函数并解析返回结果
            do = True
            while do:
                answer = chat_once(question, args.model, args.key, f"{args.scheme}://{args.host}:{args.port}")
                answer_json = json.loads(answer)
                content = answer_json["content"].replace("\n", "")
                tokens = answer_json["total_tokens"]
                if content in ["正面", "负面", "中性"]:
                    do = False
            
            # 插入列表
            sentiment_list.append(content)
            token_list.append(tokens)
            print(f"第{i}条评论处理完成")
            i += 1

        # 合并列，设置列名
        df_result = pd.DataFrame({
            '商品评论': comment_list,
            '情感分类': sentiment_list,
            'token消耗': token_list
        })

        # 写入Excel文件
        with pd.ExcelWriter('output.xlsx') as writer:
            df_result.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            worksheet.column_dimensions['A'].width = 50  # 商品评论列
            worksheet.column_dimensions['B'].width = 15  # 情感分类列
            worksheet.column_dimensions['C'].width = 15  # token消耗列

        print("全部处理完成。")

    except Exception as e:
        print(f"发生错误：{str(e)}")

if __name__ == "__main__":
    main()
