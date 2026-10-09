import os
from pptx import Presentation

# 1. テンプレートファイルの読み込み
template_path = 'templates/template.pptx'

if os.path.exists(template_path):
    prs = Presentation(template_path)
    print(f"テンプレート '{template_path}' を読み込みました。")
else:
    prs = Presentation()
    print("テンプレートが見つからないため、新規プレゼンテーションを作成します。")

# 2. タイトルスライドの追加
title_slide_layout = prs.slide_layouts[0]
slide1 = prs.slides.add_slide(title_slide_layout)

title = slide1.shapes.title
subtitle = slide1.placeholders[1]

title.text = "事業計画提案書"
subtitle.text = "ガイドラインに基づく自動生成スライド\n作成日: 2026年10月"

# 3. 本文スライドの追加
bullet_slide_layout = prs.slide_layouts[1]
slide2 = prs.slides.add_slide(bullet_slide_layout)

shapes = slide2.shapes
title_shape = shapes.title
body_shape = shapes.placeholders[1]

title_shape.text = "プロジェクト概要"

tf = body_shape.text_frame
tf.text = "ガイドラインに沿ったメインカラー（グリーン）の適用方針"

p1 = tf.add_paragraph()
p1.text = "メインカラー: ディープグリーン (#006633)"
p1.level = 1

p2 = tf.add_paragraph()
p2.text = "アクセントカラー: ミントグリーン (#20B2AA)"
p2.level = 1

# 4. 完成したファイルの保存
os.makedirs('outputs', exist_ok=True)
output_path = 'outputs/generated_presentation.pptx'
prs.save(output_path)

print(f"スライドを作成し、'{output_path}' に保存しました！")
