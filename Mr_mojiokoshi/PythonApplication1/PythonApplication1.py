import whisper
import os
import re
import openpyxl
from tkinter import filedialog, Tk

# 文の分割：句点・読点に基づいて調整（7文字±4文字）
def split_text_with_balanced_length(text, base_len=7, margin=4):
    # まず「。」「？」「！」などで大きく分割
    raw_sentences = re.split(r'(?<=[。！？])\s*', text)
    all_chunks = []

    for sentence in raw_sentences:
        sentence = sentence.strip()
        if not sentence:
            continue

        # 長い文は調整して分割
        while len(sentence) > base_len + margin:
            split_pos = -1
            # できれば「、」「。」などで切りたい（句読点優先）
            for i in range(base_len + margin, base_len - margin - 1, -1):
                if i < len(sentence) and sentence[i] in '、。':
                    split_pos = i + 1
                    break
            if split_pos == -1:
                # 見つからなければ無理やり base_len で切る
                split_pos = base_len

            chunk = sentence[:split_pos].strip()
            all_chunks.append(chunk)
            sentence = sentence[split_pos:]
        if sentence:
            all_chunks.append(sentence.strip())
    return all_chunks

# ② Excelへの保存
def save_to_excel(sentences, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "文字起こし"

    for i, sentence in enumerate(sentences):
        ws.cell(row=i + 1, column=1, value=sentence)

    wb.save(output_path)

# ③ ファイル選択＆出力先フォルダ選択
def main():
    # tkinterウィンドウ非表示
    root = Tk()
    root.withdraw()

    # mp3ファイル選択
    file_path = filedialog.askopenfilename(
        title="文字起こしするMP3ファイルを選んでください",
        filetypes=[("MP3 files", "*.mp3")]
    )
    if not file_path:
        print("ファイルが選ばれませんでした。")
        return

    # 保存先フォルダ選択
    save_dir = filedialog.askdirectory(title="出力先フォルダを選んでください")
    if not save_dir:
        print("保存先が選ばれませんでした。")
        return

    print("音声の文字起こしを開始します...")

    # whisperモデル読み込み
    model = whisper.load_model("base")

    # 音声認識
    result = model.transcribe(file_path, language="ja")
    text = result["text"]

    # 文の分割（バランス取りあり）
    sentences = split_text_with_balanced_length(text)

    # 出力ファイル名の生成
    base_name = os.path.basename(file_path).replace(".mp3", "")
    output_path = os.path.join(save_dir, f"{base_name}_transcribed.xlsx")

    # Excel出力
    save_to_excel(sentences, output_path)

    print("完了しました！保存先：", output_path)

if __name__ == "__main__":
    main()
