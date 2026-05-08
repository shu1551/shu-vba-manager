import sys
import io

# エンコーディング設定（Windows環境での文字化け防止）
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# select_macro_gui.py を呼び出してGUIを起動するエントリーポイント
if __name__ == "__main__":
    try:
        import select_macro_gui
        select_macro_gui.main()
    except Exception as e:
        print(f"GUIの起動に失敗しました: {e}")
        input("Enterキーを押して終了してください...")
