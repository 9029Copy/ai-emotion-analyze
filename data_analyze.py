import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import argparse, httpx, sys, json, os
import json
import io
from jsonschema import validate

CONFIG_FILE = "config.json"
with open(CONFIG_FILE, encoding="utf-8") as f:
    cfg = json.load(f)

HOST = cfg.get("host", "127.0.0.1")
PORT = cfg.get("port", 8000)
SCHEME = cfg.get("scheme", "http")
MODEL = cfg.get("model", "THUDM/GLM-4-9B-0414")

SCHEMA = {
    "type": "object",
    "required": ["sentiment", "tags", "score"],
    "properties": {
        "sentiment": {"type": "string", "enum": ["正面", "负面", "中性"]},
        "tags": {
            "type": "array",
            "items": {"type": "string", "minLength": 1},
            "minItems": 1
        },
        "score": {"type": "number", "minimum": 1, "maximum": 10}
    },
    "additionalProperties": False   # 禁止多余字段
}

def validating(raw: str) -> dict:
    data = json.loads(raw)
    validate(instance=data, schema=SCHEMA)
    return data

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
        tags_list = []
        score_list = []
        status_list = []

        i = 1
        for text in comment_list:
            text = text.replace("\n", "")
            question = f'''请对一句商品评论进行情感分类（正面/负面/中性，
                        其中既表现出正面也表现出负面为中性），并进行标签提取（如[手机, 屏幕]），
                        并输出评分（1-10分）。输出结果必须是一个可直接解析的json字符串，
                        一定不使用markdown，不要参考对话历史记录。\n
                        示例：\n
                        输入1：\n
                        这个手机很好用，拍照也很清晰。\n
                        输出1：\n
                        {{"sentiment": "正面", "tags": ["手机", "拍照"], "score": 9}}\n
                        输入2：\n
                        这个手机拍照很清晰，但是屏幕太小了。\n
                        输出2：\n
                        {{"sentiment": "中性", "tags": ["手机", "拍照", "屏幕"], "score": 6}}\n
                        待处理的句子是：{text}
                        '''
            # 调用chat_once函数并解析返回结果
            do = True
            while do:
                answer = chat_once(question, args.model, args.key, f"{args.scheme}://{args.host}:{args.port}")
                answer_json = json.loads(answer)
                content = answer_json["content"].replace("\n", "")
                tokens = answer_json["total_tokens"]
                try:
                    content_json = validating(content)
                    do = False
                except Exception as e:
                    print(f"第{i}条评论处理失败\nAI返回结果为：{content}\n报错：{str(e)}\n正在重试...")
            
            # 插入列表
            sentiment_list.append(content_json["sentiment"])
            tags_list.append(content_json["tags"])
            score_list.append(content_json["score"])
            token_list.append(tokens)
            status_list.append("成功")
            print(f"第{i}条评论处理完成")
            i += 1

        # 合并列，设置列名
        df_result = pd.DataFrame({
            '商品评论': comment_list,
            '情感分类': sentiment_list,
            '标签': tags_list,
            '评分': score_list,
            '处理状态': status_list,
            'token消耗': token_list
        })

        # 写入Excel文件
        with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            # 设置列宽
            worksheet.column_dimensions['A'].width = 50  # 商品评论列
            worksheet.column_dimensions['B'].width = 10  # 情感分类列
            worksheet.column_dimensions['C'].width = 40  # 标签列
            worksheet.column_dimensions['D'].width = 10  # 评分列
            worksheet.column_dimensions['E'].width = 10  # 处理状态列
            worksheet.column_dimensions['F'].width = 10  # token消耗列

        print(f"全部处理成功，共处理{i-1}条评论。")

    except Exception as e:
        print(f"发生错误：{str(e)}")

if __name__ == "__main__":
    main()
