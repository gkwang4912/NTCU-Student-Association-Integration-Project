#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
情緒分析程式
讀取 Excel 檔案並進行情緒分析
"""

import json
import requests
import pandas as pd
from typing import Dict, Any
import time
import glob
import os
from datetime import datetime

class EmotionAnalyzer:
    """情緒分析類別"""
    
    def __init__(self, config_file: str = "api.json"):
        """
        初始化情緒分析器
        
        Args:
            config_file: 設定檔路徑，包含 API 金鑰
        """
        self.base_url = "https://generativelanguage.googleapis.com/v1beta"
        self.config = self.load_config(config_file)
        self.api_key = self.config.get("api_key")
        
        if not self.api_key:
            raise ValueError("API 金鑰未找到，請檢查 api.json 檔案")
        
        self.current_model = "gemini-2.5-flash"  # 預設使用 Flash 模型
        
        # 情緒分析提示詞
        self.emotion_prompt = """你是一個簡單的情緒判斷助理，僅根據整體輸入內容判斷其情緒為正向、中性或負向。  
請將整個輸入視為一個完整內容，而非逐點或逐句分析。
判斷時，需有明確正向、負向詞彙才判斷成正負向，否則皆視為中性

輸出結果僅為單一數字：  
'1'表示正向，'0'表示中性，'-1'表示負向。
禁止生成任何其他字符、逐點回應或額外內容。

你所需要知道的資訊：
1、小詠、大詠是宿舍的名字
2、學餐是指學生餐廳，最近因為學校政策而關閉
3、操場最近換了顏色

請分析以下文本：
"""
    
    def load_config(self, config_file: str) -> Dict[str, Any]:
        """載入設定檔"""
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"設定檔 {config_file} 未找到")
            return {}
        except json.JSONDecodeError:
            print(f"設定檔 {config_file} 格式錯誤")
            return {}
    
    def send_message(self, message: str) -> Dict[str, Any]:
        """
        發送文字訊息給 Gemini API
        
        Args:
            message: 要發送的文字訊息
            
        Returns:
            Dict: API 回應
        """
        url = f"{self.base_url}/models/{self.current_model}:generateContent"
        
        headers = {
            "Content-Type": "application/json"
        }
        
        # 構建請求內容 - 只支援文字
        payload = {
            "contents": [{"parts": [{"text": message}]}],
            "generationConfig": {
                "temperature": 0.7,
                "topK": 40,
                "topP": 0.95,
                "maxOutputTokens": 8192,
            }
        }
        
        # 添加 API 金鑰到 URL
        url_with_key = f"{url}?key={self.api_key}"
        
        try:
            response = requests.post(url_with_key, headers=headers, json=payload)
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.RequestException as e:
            return {"error": f"請求錯誤: {e}"}
        except json.JSONDecodeError as e:
            return {"error": f"回應解析錯誤: {e}"}
    
    def extract_text_from_response(self, response: Dict[str, Any]) -> str:
        """從 API 回應中提取文字內容"""
        try:
            candidates = response.get("candidates", [])
            if candidates:
                content = candidates[0].get("content", {})
                parts = content.get("parts", [])
                if parts:
                    return parts[0].get("text", "無回應內容")
            return "無有效回應"
        except (KeyError, IndexError):
            return "回應格式錯誤"
    
    def analyze_emotion(self, text: str) -> str:
        """
        分析單一文本的情緒，遇到錯誤會持續重試直到成功
        
        Args:
            text: 要分析的文本
            
        Returns:
            str: 情緒分數 ('1', '0', '-1')
        """
        full_message = self.emotion_prompt + text
        retry_count = 0
        
        while True:  # 無限重試直到成功
            retry_count += 1
            response = self.send_message(full_message)
            
            if "error" in response:
                if "429" in str(response["error"]):  # 頻率限制錯誤
                    wait_time = min(60, 5 * retry_count)  # 最多等待60秒
                    print(f"遇到頻率限制 (第{retry_count}次重試)，等待 {wait_time} 秒後重試...")
                    time.sleep(wait_time)
                    continue
                elif "quota" in str(response["error"]).lower() or "limit" in str(response["error"]).lower():
                    print(f"遇到配額限制 (第{retry_count}次重試)，等待 30 秒後重試...")
                    time.sleep(30)
                    continue
                else:
                    print(f"API 錯誤 (第{retry_count}次重試): {response['error']}")
                    print("等待 10 秒後重試...")
                    time.sleep(10)
                    continue
            
            # 成功獲得回應，處理結果
            result = self.extract_text_from_response(response)
            # 清理結果，只保留數字
            cleaned_result = result.strip().replace("'", "").replace('"', "")
            
            # 驗證結果是否為有效的情緒分數
            if cleaned_result in ['1', '0', '-1']:
                if retry_count > 1:
                    print(f"重試成功！(共重試 {retry_count-1} 次)")
                return cleaned_result
            else:
                print(f"無效的API回應 (第{retry_count}次重試): {result}")
                print("等待 5 秒後重試...")
                time.sleep(5)
                continue
    
    def process_excel_file(self, input_file: str, output_file: str, text_column: str = None):
        """
        處理 Excel 檔案進行情緒分析，支援斷點續傳
        
        Args:
            input_file: 輸入的 Excel 檔案路徑
            output_file: 輸出的 Excel 檔案路徑
            text_column: 包含文本的欄位名稱，如果為 None 則自動偵測
        """
        print(f"讀取檔案: {input_file}")
        
        try:
            # 讀取 Excel 檔案
            df = pd.read_excel(input_file)
            print(f"成功讀取 {len(df)} 筆資料")
            print(f"可用欄位: {df.columns.tolist()}")
            
            # 自動偵測文本欄位
            if text_column is None:
                text_columns = []
                for col in df.columns:
                    if df[col].dtype == 'object':  # 字串類型欄位
                        # 檢查是否包含文字內容
                        sample_texts = df[col].dropna().head(3)
                        if len(sample_texts) > 0:
                            avg_length = sample_texts.astype(str).str.len().mean()
                            if avg_length > 5:  # 平均長度大於5的欄位可能是文本
                                text_columns.append(col)
                
                if not text_columns:
                    raise ValueError("未找到合適的文本欄位")
                
                # 選擇最可能的文本欄位（通常是長度最長的）
                text_column = max(text_columns, key=lambda col: df[col].astype(str).str.len().mean())
                print(f"自動選擇文本欄位: {text_column}")
            
            if text_column not in df.columns:
                raise ValueError(f"指定的欄位 '{text_column}' 不存在")
            
            # 檢查是否已有部分結果（斷點續傳）
            start_index = 0
            try:
                existing_df = pd.read_excel(output_file)
                if '情緒分析' in existing_df.columns and '情緒標籤' in existing_df.columns:
                    # 找到最後一個已處理的索引
                    processed_mask = (existing_df['情緒分析'] != '') & (existing_df['情緒標籤'] != '')
                    if processed_mask.any():
                        start_index = processed_mask.sum()
                        print(f"發現已處理的資料，從第 {start_index + 1} 筆開始繼續...")
                        df = existing_df.copy()  # 使用已有的資料
            except FileNotFoundError:
                print("輸出檔案不存在，將從頭開始處理")
            except Exception as e:
                print(f"讀取已有結果時發生錯誤: {e}，將從頭開始處理")
            
            # 如果是新開始，建立新欄位
            if start_index == 0:
                df['情緒分析'] = ""
                df['情緒標籤'] = ""
            
            # 逐行分析情緒
            total_rows = len(df)
            
            try:
                for idx in range(start_index, total_rows):
                    row = df.iloc[idx]
                    text = str(row[text_column])
                    
                    print(f"\n第 {idx+1}/{total_rows} 行:")
                    print(f"文本: {text[:100]}{'...' if len(text) > 100 else ''}")
                    
                    if pd.isna(row[text_column]) or text.strip() == "" or text == "nan":
                        emotion_score = "0"
                        emotion_label = "中性"
                        print("空白文本 -> 中性")
                    else:
                        print("分析中...")
                        emotion_score = self.analyze_emotion(text)
                        
                        # 轉換分數為標籤
                        if emotion_score == "1":
                            emotion_label = "正向"
                        elif emotion_score == "-1":
                            emotion_label = "負向"
                        else:
                            emotion_label = "中性"
                        
                        print(f"結果: {emotion_label} ({emotion_score})")
                    
                    df.at[idx, '情緒分析'] = emotion_score
                    df.at[idx, '情緒標籤'] = emotion_label
                    
                    # 每處理5筆就保存一次（避免意外中斷損失進度）
                    if (idx + 1) % 5 == 0 or idx == total_rows - 1:
                        df.to_excel(output_file, index=False)
                        print(f"已保存進度到第 {idx + 1} 筆")
                    
                    # 避免 API 頻率限制
                    time.sleep(3)  # 增加延遲到3秒
                    
            except KeyboardInterrupt:
                print(f"\n\n程式被中斷，已保存進度到第 {idx + 1} 筆")
                df.to_excel(output_file, index=False)
                print(f"結果已保存到: {output_file}")
                return
            
            # 最終保存
            df.to_excel(output_file, index=False)
            print(f"\n分析完成！結果已儲存到: {output_file}")
            
            # 統計結果
            processed_data = df[df['情緒標籤'] != '']
            emotion_counts = processed_data['情緒標籤'].value_counts()
            print("\n情緒分析統計:")
            for emotion, count in emotion_counts.items():
                percentage = (count / len(processed_data)) * 100
                print(f"  {emotion}: {count} 筆 ({percentage:.1f}%)")
                
        except Exception as e:
            print(f"處理檔案時發生錯誤: {e}")
            raise

def generate_statistics_report(result_file: str) -> None:
    """
    生成情緒分析統計報告
    
    Args:
        result_file: 情緒分析結果檔案路徑
    """
    try:
        print(f"\n正在生成統計報告: {result_file}")
        
        # 讀取 Excel 檔案
        df = pd.read_excel(result_file)
        
        # 檢查是否有情緒分析欄位
        if '情緒標籤' not in df.columns and '情緒分析' not in df.columns:
            print(f"  ✗ {result_file} 中沒有找到情緒分析欄位")
            return
        
        # 使用情緒標籤欄位進行統計
        if '情緒標籤' in df.columns:
            emotion_col = '情緒標籤'
        else:
            emotion_col = '情緒分析'
        
        # 過濾掉空白資料
        valid_data = df[df[emotion_col].notna() & (df[emotion_col] != '')]
        
        # 統計各情緒數量
        emotion_counts = valid_data[emotion_col].value_counts()
        
        # 計算總數和百分比
        total = len(valid_data)
        positive = emotion_counts.get('正向', 0) if '正向' in emotion_counts else emotion_counts.get('1', 0)
        negative = emotion_counts.get('負向', 0) if '負向' in emotion_counts else emotion_counts.get('-1', 0)
        neutral = emotion_counts.get('中性', 0) if '中性' in emotion_counts else emotion_counts.get('0', 0)
        
        # 計算百分比
        positive_pct = (positive / total * 100) if total > 0 else 0
        negative_pct = (negative / total * 100) if total > 0 else 0
        neutral_pct = (neutral / total * 100) if total > 0 else 0
        
        # 生成報告內容
        report = []
        report.append("情緒分析統計:")
        report.append(f"  中性: {neutral} 筆 ({neutral_pct:.1f}%)")
        report.append(f"  正向: {positive} 筆 ({positive_pct:.1f}%)")
        report.append(f"  負向: {negative} 筆 ({negative_pct:.1f}%)")
        
        # 將報告轉為字串
        report_text = "\n".join(report)
        
        # 生成輸出檔案名稱
        base_name = result_file.replace('_情緒分析結果.xlsx', '')
        output_file = f"{base_name}_統計報告.txt"
        
        # 寫入 txt 檔案
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(report_text)
        
        print(f"  ✓ 統計報告已儲存到: {output_file}")
        print(f"  中性: {neutral} 筆 | 正向: {positive} 筆 | 負向: {negative} 筆")
        
    except Exception as e:
        print(f"  ✗ 生成統計報告時發生錯誤: {e}")

def main():
    """主程式入口"""
    try:
        # 初始化情緒分析器
        analyzer = EmotionAnalyzer()
        
        # 開始分析
        print("=== 情緒分析程式 ===")
        print("特點:")
        print("- 遇到 API 錯誤會自動重試，直到成功為止")
        print("- 每5筆資料自動保存進度")
        print("- 支援斷點續傳，可從中斷處繼續執行")
        print("- 使用 Ctrl+C 可以安全中斷並保存進度")
        print("- 自動處理當前目錄中所有的 xlsx 檔案")
        print("-" * 50)
        
        # 搜尋當前目錄中所有的 xlsx 檔案
        import glob
        import os
        
        xlsx_files = glob.glob("*.xlsx")
        
        # 過濾掉已處理的結果檔案（包含「情緒分析結果」或以「~$」開頭的暫存檔）
        xlsx_files = [f for f in xlsx_files if not f.startswith("~$") and "情緒分析結果" not in f]
        
        if not xlsx_files:
            print("錯誤: 當前目錄中沒有找到任何 xlsx 檔案")
            print(f"當前目錄: {os.getcwd()}")
            return
        
        print(f"\n找到 {len(xlsx_files)} 個 Excel 檔案:")
        for i, file in enumerate(xlsx_files, 1):
            print(f"  {i}. {file}")
        print()
        
        # 處理每個 xlsx 檔案
        for input_file in xlsx_files:
            try:
                # 生成輸出檔案名稱
                base_name = os.path.splitext(input_file)[0]
                output_file = f"{base_name}_情緒分析結果.xlsx"
                
                print(f"\n{'='*60}")
                print(f"開始處理: {input_file}")
                print(f"輸出檔案: {output_file}")
                print(f"{'='*60}\n")
                
                analyzer.process_excel_file(input_file, output_file)
                
                print(f"\n✓ {input_file} 處理完成！")
                
                # 生成統計報告
                generate_statistics_report(output_file)
                print()
                
            except KeyboardInterrupt:
                print(f"\n程式被使用者中斷，{input_file} 的進度已保存")
                raise  # 重新拋出以停止處理其他檔案
            except Exception as e:
                print(f"\n✗ 處理 {input_file} 時發生錯誤: {e}")
                print("繼續處理下一個檔案...\n")
                continue
        
        print("\n" + "="*60)
        print("所有檔案處理完成！")
        print("="*60)
        
    except KeyboardInterrupt:
        print("\n程式被使用者中斷")
    except Exception as e:
        print(f"程式執行失敗: {e}")
        print("請檢查 api.json 檔案和 Excel 檔案是否存在")

if __name__ == "__main__":
    main()
