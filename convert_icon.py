from PIL import Image
import os

def convert_png_to_ico(png_path, ico_path):
    """PNGファイルをICOファイルに変換します"""
    try:
        img = Image.open(png_path)
        
        # ICOファイルは通常正方形なので、正方形にリサイズ
        size = max(img.size)
        square_img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        
        # 元の画像を中央に配置
        paste_x = (size - img.size[0]) // 2
        paste_y = (size - img.size[1]) // 2
        square_img.paste(img, (paste_x, paste_y))
        
        # ICOで使用される一般的なサイズに変換
        icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
        square_img.save(ico_path, format='ICO', sizes=icon_sizes)
        print(f"アイコンの変換が完了しました: {ico_path}")
        return True
    except Exception as e:
        print(f"アイコンの変換中にエラーが発生しました: {str(e)}")
        return False

if __name__ == "__main__":
    png_path = "achievement.png"
    ico_path = "app_icon.ico"
    
    if not os.path.exists(png_path):
        print(f"エラー: {png_path} が見つかりません。")
    else:
        convert_png_to_ico(png_path, ico_path) 