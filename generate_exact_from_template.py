import os
import glob
from pptx import Presentation

# 1. 元のファイルを直接ベース（テンプレート）として読み込む
template_files = glob.glob("templates/チョキ_テーマ.pptx") or glob.glob("templates/*.pptx")
if not template_files:
    print("templates/チョキ_テーマ.pptx が見つかりません。")
    exit()

template_path = template_files[0]
print(f"ベーステンプレート: {template_path}")
prs = Presentation(template_path)

# 元のファイルの1枚目のスライドをそのまま取得（デザイン・色が100%保持されます）
slide = prs.slides[0]

# スライド内の図形（テキスト枠、表）を特定して内容を差し替え
for shape in slide.shapes:
    # タイトル部分
    if shape.has_text_frame and "予算" in shape.text_frame.text:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                # 書式（色・フォント）を壊さないよう、文字だけを更新
                if "予算" in r.text or "振り返り" in r.text:
                    r.text = "2027年度ChoQi予算 2026年度の振り返り"

    # 表（テーブル）部分
    if shape.has_table:
        table = shape.table
        print("テーブルが見つかりました。元のセルの色とスタイルを完全維持したまま値を更新します。")
        # 1行目ヘッダー
        headers = ["", "予算", "実績(9月まで)", "実績見込", "差異"]
        for c_idx, h_text in enumerate(headers):
            if c_idx < len(table.columns):
                # 既存のテキストランの書式を維持して文字だけ差し替え
                cell = table.cell(0, c_idx)
                if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
                    cell.text_frame.paragraphs[0].runs[0].text = h_text
                else:
                    cell.text = h_text

os.makedirs("outputs", exist_ok=True)
out_file = "outputs/exact_cloned_presentation.pptx"
prs.save(out_file)
print(f"作成完了: {out_file}")
