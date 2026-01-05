#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文字雲生成器
讀取 Excel 檔案中的文字內容並生成精美的文字雲圖片
"""

import pandas as pd
import jieba
import matplotlib.pyplot as plt
from wordcloud import WordCloud
import re
from collections import Counter
import numpy as np
from PIL import Image
import os
import glob
import sys

def read_excel_text(file_path):
    """
    讀取 Excel 檔案中的所有文字內容
    
    Args:
        file_path (str): Excel 檔案路徑
    
    Returns:
        str: 合併後的文字內容
    """
    try:
        # 讀取 Excel 檔案，嘗試所有工作表
        excel_file = pd.ExcelFile(file_path)
        all_text = []
        
        print(f"發現 {len(excel_file.sheet_names)} 個工作表: {excel_file.sheet_names}")
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"處理工作表: {sheet_name} (形狀: {df.shape})")
            
            # 將所有文字型態的欄位內容合併
            for column in df.columns:
                text_values = df[column].dropna().astype(str)
                all_text.extend(text_values.tolist())
        
        combined_text = ' '.join(all_text)
        print(f"總共讀取了 {len(combined_text)} 個字符")
        return combined_text
        
    except Exception as e:
        print(f"讀取 Excel 檔案時發生錯誤: {e}")
        return ""

def preprocess_text(text):
    """
    預處理文字內容
    
    Args:
        text (str): 原始文字
    
    Returns:
        str: 處理後的文字
    """
    # 移除特殊字符，但保留中文、英文、數字
    text = re.sub(r'[^\u4e00-\u9fff\w\s]', ' ', text)
    
    # 移除多餘的空白
    text = re.sub(r'\s+', ' ', text)
    
    # 移除過短的字詞（少於2個字符）
    words = text.split()
    filtered_words = [word for word in words if len(word) >= 2]
    
    return ' '.join(filtered_words)

def load_stopwords(stopwords_file='stopwords.txt'):
    """
    從檔案載入停用詞
    
    Args:
        stopwords_file (str): 停用詞檔案路徑
    
    Returns:
        set: 停用詞集合
    """
    stop_words = set()
    
    # 嘗試從檔案讀取停用詞
    try:
        with open(stopwords_file, 'r', encoding='utf-8') as f:
            for line in f:
                word = line.strip()
                if word and not word.startswith('#'):  # 忽略空行和註解
                    stop_words.add(word)
        print(f"從 {stopwords_file} 載入了 {len(stop_words)} 個停用詞")
    except FileNotFoundError:
        print(f"找不到停用詞檔案 {stopwords_file}，使用內建停用詞")
    
    # 內建的停用詞（作為備份）
    default_stop_words = {
        # 基本停用詞
        '的', '了', '在', '是', '我', '有', '和', '就', 
        '不', '人', '都', '一', '一個', '上', '也', '很', 
        '到', '說', '要', '去', '你', '會', '著', '沒有',
        '看', '好', '自己', '這', '那', '但是', '然後',
        '可以', '還是', '因為', '所以', '如果', '這樣',
        '什麼', '怎麼', '為什麼', '這個', '那個', '現在',
        '已經', '可能', '應該', '覺得', '知道', '想要',
        
        # 代詞
        '他', '她', '它', '我們', '你們', '他們', '她們',
        '自己', '本人', '某', '某個', '某些', '另', '另外',
        '別', '別人', '別的', '各', '各個', '各自',
        
        # 連詞和語氣詞
        '而且', '並且', '或者', '以及', '還有', '另外',
        '此外', '除了', '無論', '不管', '儘管', '雖然',
        '但', '卻', '只是', '只有', '除非', '如果',
        '啊', '呀', '嗎', '呢', '吧', '嘛', '哦', '哎',
        
        # 方位詞
        '前', '後', '左', '右', '東', '南', '西', '北',
        '中', '內', '外', '裡', '面', '邊', '旁', '間',
        
        # 數量詞
        '二', '三', '四', '五', '六', '七', '八', '九', '十',
        '百', '千', '萬', '億', '第一', '第二', '第三',
        '一些', '一點', '一下', '一直', '一起', '一樣',
        
        # 時間詞
        '今天', '昨天', '明天', '今年', '去年', '明年',
        '早上', '下午', '晚上', '現在', '以前', '以後',
        '剛才', '馬上', '立刻', '忽然', '突然', '最近',
        
        # 程度副詞
        '非常', '十分', '特別', '尤其', '格外', '更加',
        '比較', '相當', '頗', '挺', '蠻', '太', '最',
        '更', '還', '再', '又', '越來越', '漸漸',
        
        # 常用動詞
        '做', '搞', '弄', '來', '走', '跑', '飛', '坐',
        '站', '躺', '睡', '吃', '喝', '穿', '戴', '拿',
        '給', '送', '帶', '買', '賣', '用', '玩', '學',
        
        # 助詞和語氣
        '著', '了', '過', '起來', '下來', '出來', '進來',
        '回來', '過來', '起', '下', '出', '進', '回', '過',
        
        # 無意義詞
        'nan', 'NaN', '', ' ', '　', '\n', '\t', '\r',
        '哈', '呵', '嗯', '嗚', '咦', '咯', '喔', '噢',
        
        # 網路用語
        '笑死', '真的', '超級', '根本', '完全', '絕對',
        '肯定', '一定', '必須', '必要', '需要', '想', '感覺',

        #自訂停用詞
        '你媽'
    }
    
    # 合併檔案中的停用詞和內建停用詞
    stop_words.update(default_stop_words)
    
    return stop_words

def segment_chinese_text(text):
    """
    使用 jieba 對中文文字進行分詞
    
    Args:
        text (str): 輸入文字
    
    Returns:
        list: 分詞結果列表
    """
    # 使用 jieba 分詞
    words = jieba.lcut(text)
    
    # 載入停用詞
    stop_words = load_stopwords()
    
    # 過濾停用詞和短詞
    filtered_words = [
        word for word in words 
        if word not in stop_words 
        and len(word) >= 2 
        and word.strip() != ''
        and not word.isdigit()  # 過濾純數字
    ]
    
    return filtered_words

def create_wordcloud(text, output_path='wordcloud.png'):
    """
    創建文字雲圖片
    
    Args:
        text (str): 輸入文字
        output_path (str): 輸出圖片路徑
    """
    if not text.strip():
        print("沒有有效的文字內容可以生成文字雲")
        return
    
    # 預處理文字
    processed_text = preprocess_text(text)
    print(f"預處理後的文字長度: {len(processed_text)}")
    
    # 中文分詞
    words = segment_chinese_text(processed_text)
    print(f"分詞結果: {len(words)} 個詞語")
    
    if len(words) < 5:
        print("詞語數量太少，無法生成有效的文字雲")
        return
    
    # 統計詞頻
    word_freq = Counter(words)
    print(f"最常見的10個詞語: {word_freq.most_common(10)}")
    
    # 將詞語重新組合成字符串
    text_for_wordcloud = ' '.join(words)
    
    # 設定文字雲參數
    wordcloud = WordCloud(
        font_path='C:/Windows/Fonts/msjh.ttc',  # 微軟正黑體
        width=1920,
        height=1080,
        background_color='white',
        max_words=200,
        colormap='viridis',
        relative_scaling=0.5,
        random_state=42,
        collocations=False
    )
    
    try:
        # 生成文字雲
        wordcloud.generate(text_for_wordcloud)
        
        # 創建圖表
        plt.figure(figsize=(19.2, 10.8), dpi=100)
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.tight_layout(pad=0)
        
        # 保存圖片
        plt.savefig(output_path, dpi=300, bbox_inches='tight', 
                   facecolor='white', edgecolor='none')
        plt.show()
        
        print(f"文字雲已成功保存到: {output_path}")
        
    except Exception as e:
        print(f"生成文字雲時發生錯誤: {e}")

def main():
    """
    主函數
    """
    # 若有命令列參數，優先使用傳入的檔案路徑
    excel_file = None
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]

    # 若未提供參數，則自動搜尋當前資料夾中的 .xlsx 檔，取第一個
    if not excel_file:
        xlsx_files = sorted(glob.glob(os.path.join(os.getcwd(), "*.xlsx")))
        if not xlsx_files:
            print("找不到任何 .xlsx 檔案，請將 Excel 檔放在本程式同一資料夾，或傳入檔案路徑作為參數")
            print(f"目前資料夾 ({os.getcwd()}) 的檔案清單: {os.listdir(os.getcwd())}")
            return
        excel_file = xlsx_files[0]

    # 根據 Excel 檔名自動產生輸出圖檔名稱
    base_name = os.path.splitext(os.path.basename(excel_file))[0]
    output_file = f"{base_name}_文字雲.png"
    
    print("="*50)
    print("文字雲生成器啟動")
    print("="*50)
    
    # 讀取 Excel 檔案
    text_content = read_excel_text(excel_file)
    
    if not text_content:
        print("無法讀取到有效的文字內容")
        return
    
    print(f"讀取到的文字內容前100字符: {text_content[:100]}...")
    
    # 生成文字雲
    create_wordcloud(text_content, output_file)
    
    print("="*50)
    print("程式執行完成")
    print("="*50)

if __name__ == "__main__":
    main()
