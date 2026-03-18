import sys
import traceback

try:
    import win32com.client
    import pythoncom
    import urllib.request
    import tempfile
    import os
    import base64
    from flask import Flask, request, jsonify
    from flask_cors import CORS
except ImportError as e:
    print("\n[错误] 缺少必要的 Python 运行库！")
    print(f"详情: {e}")
    print("\n请运行以下命令进行安装：")
    print("pip install flask flask-cors pywin32 requests")
    input("\n按回车键退出本窗口...")
    sys.exit(1)

app = Flask(__name__)
# 允许跨域请求，这样你的网页就可以向本地服务发消息了
CORS(app)

@app.route('/send_to_ps', methods=['POST'])
def send_to_ps():
    print("Received request to send to Photoshop...")
    pythoncom.CoInitialize()
    try:
        data = request.json
        image_data = data.get('image_url')
        
        if not image_data:
            return jsonify({"success": False, "error": "没有提供图片数据"}), 400

        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, 'ps_temp_image.png')

        # 如果是 Base64 数据
        if image_data.startswith('data:image'):
            # 提取 base64 部分
            header, encoded = image_data.split(",", 1)
            with open(temp_path, "wb") as fh:
                fh.write(base64.b64decode(encoded))
        # 如果是 URL 链接 (http/https)
        elif image_data.startswith('http'):
            import requests
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}
            response = requests.get(image_data, headers=headers)
            if response.status_code == 200:
                with open(temp_path, 'wb') as f:
                    f.write(response.content)
            else:
                return jsonify({"success": False, "error": f"下载图片失败，状态码: {response.status_code}"}), 400
        else:
            return jsonify({"success": False, "error": "不支持的图片格式（仅支持 http/https 链接或 base64）"}), 400

        print(f"Image saved to temp path: {temp_path}")

        # 尝试连接到当前正在运行的 Photoshop
        try:
            psApp = win32com.client.Dispatch("Photoshop.Application")
            print("Connected to Photoshop.")
        except Exception as e:
            print("Could not connect to Photoshop:", e)
            return jsonify({"success": False, "error": "无法连接到 Photoshop。请确保 Photoshop 已打开。详细错误：" + str(e)}), 500
        
        # 检查 PS 里有没有打开的文档，如果没有，建一个
        try:
            target_doc = psApp.ActiveDocument
        except Exception:
            print("No active document, creating a new one.")
            # 创建一个默认文档作为底板 (1920x1080)
            target_doc = psApp.Documents.Add(1920, 1080, 72, "New Document")
            
        print("Duplicating image layer at original resolution...")
        # 1. 以新文档形式在后台打开刚刚下载的高清原图
        temp_doc = psApp.Open(temp_path)
        # 2. 获取原图层并直接跨文档复制 (避免了剪贴板和图层蒙版导致的“粘贴不可用”报错)
        temp_layer = temp_doc.ActiveLayer
        temp_layer.Duplicate(target_doc)
        # 3. 关闭临时图片文档 (2 = 不保存修改)
        temp_doc.Close(2)
        
        # 4. 切换回用户的目标文档
        psApp.ActiveDocument = target_doc
        
        print("Image pasted successfully.")

        return jsonify({"success": True, "message": "已成功发送到 PS！"})

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    try:
        print("==================================================")
        print("🎨 Nano Banana - Photoshop 桥接服务运行中！")
        print("请保持这个黑色窗口打开（可以最小化）。")
        print("在网页上点击【发送到 PS】即可。")
        print("==================================================")
        app.run(port=5000)
    except Exception as e:
        print("\n发生意外错误导致程序停止：")
        traceback.print_exc()
        input("\n按回车键退出本窗口...")

