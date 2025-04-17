# ALTIDES (アルタイデス)
**Alternative Text Insertion and Dynamic Extraction System**

## 概要
ALTIDESは、LM Studio上のγ（gamma）モデルを利用して、PowerPoint（PPTX）、Word（DOCX）、PDFファイル内の画像（写真、図表、グラフなど）に対して自動的に代替テキストを生成・埋め込みするツールです。  
テキストボックスや吹き出しなど、すでにテキストが含まれるオブジェクトは対象外としています。

## 特徴
- **多形式対応:** PPTX、DOCX、PDFの各種ドキュメントに対応
- **LM Studio連携:** LM Studio上のγモデルを利用し、画像認識とAlt Text生成を実施
- **GUI搭載:** Tkinterによるファイル／フォルダ選択UIで一括処理が可能
- **柔軟な設定:** コード上部にパラメータをまとめ、環境に合わせた調整が容易

## インストール
### 依存ライブラリのインストール
以下のコマンドで必要なPythonパッケージをインストールしてください。

```bash
pip install python-pptx python-docx PyMuPDF requests Pillow

## Exeコマンド
pyinstaller --noconfirm --onefile --windowed --icon=icon\icon.ico src\altides.py