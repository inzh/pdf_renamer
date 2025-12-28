import webview
import os
import base64
import shutil
import threading
import requests
import hashlib
import json
import random
import time
import fitz  # PyMuPDF
from PIL import Image, ImageOps
from io import BytesIO
from pathlib import Path
from typing import Optional, Tuple

# 显示用的最大尺寸（用于压缩预览图）
MAX_DISPLAY_SIZE = 4200

# 数据持久化文件路径
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'image_data.json')


def load_persistent_data() -> dict:
    """加载持久化数据"""
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"加载数据失败: {e}")
    return {}


def save_persistent_data(data: dict):
    """保存持久化数据"""
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存数据失败: {e}")


class Api:
    def __init__(self):
        self.images: list[str] = []  # 图片路径列表
        self.current_index: int = 0  # 当前图片索引
        self.save_path: str = ""  # 保存路径
        self.rotation: int = 0  # 当前旋转角度
        self.window: Optional[webview.Window] = None
        self._dialog_result = None
        self._dialog_event = threading.Event()
        # 记录原图尺寸（旋转后），用于OCR坐标计算
        self.original_width: int = 0
        self.original_height: int = 0
        # 记录显示图尺寸
        self.display_width: int = 0
        self.display_height: int = 0
        # 当前打开的文件夹路径
        self.current_folder: str = ""

    def set_window(self, window: webview.Window):
        self.window = window

    def open_file(self) -> dict:
        """打开单个图片或PDF文件"""
        try:
            file_types = ('Description (*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.webp;*.pdf;*.tiff;*.tif)',)
            result = self.window.create_file_dialog(
                webview.FileDialog.OPEN,
                allow_multiple=False,
                file_types=file_types
            )
            if result and len(result) > 0:
                file_path = result[0]
                self.current_folder = os.path.dirname(file_path)
                self.images = [file_path]
                self.current_index = 0
                self.rotation = 0
                data = self._get_current_image_data()
                data["folderPath"] = self.current_folder
                return data
            return {"success": False, "message": "未选择文件"}
        except Exception as e:
            return {"success": False, "message": f"打开文件失败: {str(e)}"}

    def open_folder(self) -> dict:
        """打开文件夹并读取所有图片和PDF"""
        try:
            result = self.window.create_file_dialog(webview.FileDialog.FOLDER)
            if result and len(result) > 0:
                folder_path = result[0]
                self.current_folder = folder_path
                supported_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp', '.pdf'}
                self.images = []
                for file in sorted(Path(folder_path).iterdir()):
                    if file.suffix.lower() in supported_extensions:
                        self.images.append(str(file))
                if self.images:
                    self.current_index = 0
                    self.rotation = 0
                    data = self._get_current_image_data()
                    data["folderPath"] = self.current_folder
                    return data
                return {"success": False, "message": "文件夹中没有图片或PDF"}
            return {"success": False, "message": "未选择文件夹"}
        except Exception as e:
            return {"success": False, "message": f"打开文件夹失败: {str(e)}"}

    def _is_pdf(self, file_path: str) -> bool:
        """判断是否为 PDF 文件"""
        return os.path.splitext(file_path)[1].lower() == '.pdf'

    def _load_image(self, image_path: str) -> Image.Image:
        """
        加载图片，支持普通图片和 PDF（只取第一页）
        """
        if self._is_pdf(image_path):
            # 加载 PDF 第一页
            doc = fitz.open(image_path)
            page = doc[0]
            # 使用较高的分辨率渲染
            mat = fitz.Matrix(2.0, 2.0)  # 2x 缩放
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
            return img
        else:
            # 普通图片
            img = Image.open(image_path)
            # 处理 EXIF 方向
            img = ImageOps.exif_transpose(img)
            return img

    def _get_current_image_data(self) -> dict:
        """获取当前图片的数据（压缩后的预览图）"""
        if not self.images or self.current_index >= len(self.images):
            return {"success": False, "message": "没有可显示的图片"}
        
        try:
            image_path = self.images[self.current_index]
            # 使用统一的图片加载方法（支持 PDF）
            img = self._load_image(image_path)
            
            # 应用用户手动旋转
            if self.rotation != 0:
                img = img.rotate(-self.rotation, expand=True)
            
            # 记录原图尺寸（旋转后）
            self.original_width = img.width
            self.original_height = img.height
            
            # 转换颜色模式
            if img.mode in ('RGBA', 'LA', 'P'):
                img = img.convert('RGB')
            
            # 压缩图片用于显示（保持宽高比）
            display_img = img.copy()
            if img.width > MAX_DISPLAY_SIZE or img.height > MAX_DISPLAY_SIZE:
                # 计算缩放比例
                ratio = min(MAX_DISPLAY_SIZE / img.width, MAX_DISPLAY_SIZE / img.height)
                new_width = int(img.width * ratio)
                new_height = int(img.height * ratio)
                display_img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # 记录显示图尺寸
            self.display_width = display_img.width
            self.display_height = display_img.height
            
            # 转换为 base64
            buffer = BytesIO()
            display_img.save(buffer, format='JPEG', quality=85)
            img_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
            
            # 关闭图片
            img.close()
            
            return {
                "success": True,
                "image": f"data:image/jpeg;base64,{img_base64}",
                "filename": os.path.basename(image_path),
                "filepath": image_path,
                "index": self.current_index,
                "total": len(self.images),
                "width": self.display_width,  # 返回显示图尺寸
                "height": self.display_height,
                "originalWidth": self.original_width,  # 原图尺寸
                    "originalHeight": self.original_height
                }
        except Exception as e:
            return {"success": False, "message": f"读取图片失败: {str(e)}"}

    def next_image(self) -> dict:
        """切换到下一张图片"""
        if not self.images:
            return {"success": False, "message": "没有加载图片"}
        self.current_index = (self.current_index + 1) % len(self.images)
        self.rotation = 0
        return self._get_current_image_data()

    def prev_image(self) -> dict:
        """切换到上一张图片"""
        if not self.images:
            return {"success": False, "message": "没有加载图片"}
        self.current_index = (self.current_index - 1) % len(self.images)
        self.rotation = 0
        return self._get_current_image_data()

    def goto_image(self, index: int) -> dict:
        """跳转到指定索引的图片（索引从1开始）"""
        if not self.images:
            return {"success": False, "message": "没有加载图片"}
        # 用户输入的是从1开始的索引，转换为从0开始
        actual_index = index - 1
        if actual_index < 0 or actual_index >= len(self.images):
            return {"success": False, "message": f"无效的页码，请输入 1-{len(self.images)}"}
        self.current_index = actual_index
        self.rotation = 0
        return self._get_current_image_data()

    def rotate_left(self) -> dict:
        """左旋转90度"""
        self.rotation = (self.rotation - 90) % 360
        return self._get_current_image_data()

    def rotate_right(self) -> dict:
        """右旋转90度"""
        self.rotation = (self.rotation + 90) % 360
        return self._get_current_image_data()

    def save_image_data(self, folder_path: str, filename: str, name: str, number: str) -> dict:
        """
        保存图片的图名图号数据到本地
        
        参数:
            folder_path: 文件夹绝对路径
            filename: 图片文件名
            name: 图名
            number: 图号
        """
        try:
            data = load_persistent_data()
            if folder_path not in data:
                data[folder_path] = {}
            data[folder_path][filename] = {"name": name, "number": number}
            save_persistent_data(data)
            return {"success": True}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def load_folder_data(self, folder_path: str) -> dict:
        """
        加载指定文件夹的所有图片数据
        
        参数:
            folder_path: 文件夹绝对路径
        
        返回:
            该文件夹下所有图片的图名图号数据
        """
        try:
            data = load_persistent_data()
            folder_data = data.get(folder_path, {})
            return {"success": True, "data": folder_data}
        except Exception as e:
            return {"success": False, "message": str(e), "data": {}}

    def clear_folder_data(self, folder_path: str) -> dict:
        """
        清空指定文件夹的所有图片数据
        
        参数:
            folder_path: 文件夹绝对路径
        """
        try:
            data = load_persistent_data()
            if folder_path in data:
                del data[folder_path]
                save_persistent_data(data)
            return {"success": True, "message": "历史数据已清空"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def set_save_path(self) -> dict:
        """设置保存路径"""
        try:
            result = self.window.create_file_dialog(webview.FileDialog.FOLDER)
            if result and len(result) > 0:
                self.save_path = result[0]
                return {"success": True, "path": self.save_path}
            return {"success": False, "message": "未选择保存路径"}
        except Exception as e:
            return {"success": False, "message": f"设置路径失败: {str(e)}"}

    def get_save_path(self) -> str:
        """获取当前保存路径"""
        return self.save_path

    def crop_and_ocr(self, x: float, y: float, width: float, height: float, 
                     img_width: float, img_height: float) -> dict:
        """
        裁剪图片区域并执行OCR（从原图裁剪，支持PDF）
        
        参数:
            x, y: 选择框左上角坐标（相对于显示的图片在界面上的尺寸）
            width, height: 选择框的宽高
            img_width, img_height: 图片在界面上显示的尺寸（经过CSS缩放后）
        """
        if not self.images or self.current_index >= len(self.images):
            return {"success": False, "message": "没有可处理的图片"}
        
        try:
            image_path = self.images[self.current_index]
            # 使用统一的图片加载方法（支持 PDF）
            img = self._load_image(image_path)
            
            # 应用用户手动旋转
            if self.rotation != 0:
                img = img.rotate(-self.rotation, expand=True)
            
            if img.mode in ('RGBA', 'LA', 'P'):
                img = img.convert('RGB')
            
            # 原图实际尺寸
            actual_width, actual_height = img.size
            
            # 计算从界面显示尺寸到原图尺寸的缩放比例
            scale_x = actual_width / img_width
            scale_y = actual_height / img_height
            
            # 转换坐标到原图坐标
            crop_x = int(x * scale_x)
            crop_y = int(y * scale_y)
            crop_width = int(width * scale_x)
            crop_height = int(height * scale_y)
            
            # 确保坐标在有效范围内
            crop_x = max(0, min(crop_x, actual_width))
            crop_y = max(0, min(crop_y, actual_height))
            crop_x2 = max(0, min(crop_x + crop_width, actual_width))
            crop_y2 = max(0, min(crop_y + crop_height, actual_height))
            
            # 从原图裁剪
            cropped = img.crop((crop_x, crop_y, crop_x2, crop_y2))
            
            # 关闭原图
            img.close()
            
            # 执行OCR
            text = perform_ocr(cropped)
            
            return {"success": True, "text": text}
        except Exception as e:
            return {"success": False, "message": f"处理失败: {str(e)}"}

    def rename_and_save(self, image_name: str, image_number: str) -> dict:
        """
        重命名并保存图片/PDF到保存文件夹
        
        参数:
            image_name: 图名
            image_number: 图号
        """
        if not self.images or self.current_index >= len(self.images):
            return {"success": False, "message": "没有可保存的图片"}
        
        if not self.save_path:
            return {"success": False, "message": "请先设置保存路径"}
        
        if not image_name and not image_number:
            return {"success": False, "message": "图名和图号不能都为空"}
        
        # 检查 Windows 不允许的文件名字符
        invalid_chars = r'\/:*?"<>|'
        invalid_found = []
        for char in invalid_chars:
            if char in image_name or char in image_number:
                invalid_found.append(char)
        
        if invalid_found:
            chars_display = ' '.join(invalid_found)
            return {
                "success": False, 
                "message": f"文件名包含非法字符: {chars_display}\n请修改图名或图号后重试"
            }
        
        try:
            image_path = self.images[self.current_index]
            is_pdf = self._is_pdf(image_path)
            
            # 获取原文件扩展名（保持原格式）
            original_ext = os.path.splitext(image_path)[1].lower()
            if not original_ext:
                original_ext = '.jpg'
            
            # 构建新文件名
            if image_name and image_number:
                base_name = f"{image_name}-{image_number}"
            elif image_name:
                base_name = image_name
            else:
                base_name = image_number
            
            # 构建保存路径，检查是否存在同名文件
            save_file_path = os.path.join(self.save_path, f"{base_name}{original_ext}")
            
            # 如果文件已存在，添加 (1), (2), ... 后缀
            counter = 1
            while os.path.exists(save_file_path):
                save_file_path = os.path.join(self.save_path, f"{base_name}({counter}){original_ext}")
                counter += 1
            
            if is_pdf:
                # PDF: 直接复制原文件
                shutil.copy2(image_path, save_file_path)
            else:
                # 普通图片
                # 检查是否需要处理图片（有EXIF方向信息或用户旋转）
                needs_processing = False
                if self.rotation != 0:
                    needs_processing = True
                else:
                    try:
                        with Image.open(image_path) as img:
                            exif = img._getexif()
                            if exif:
                                exif_orientation = exif.get(274)
                                if exif_orientation is not None and exif_orientation != 1:
                                    needs_processing = True
                    except:
                        pass
                
                if not needs_processing:
                    # 无需处理，直接复制原文件
                    shutil.copy2(image_path, save_file_path)
                else:
                    # 需要处理：应用EXIF方向和用户旋转后保存
                    with Image.open(image_path) as img:
                        img = ImageOps.exif_transpose(img)
                        
                        if self.rotation != 0:
                            img = img.rotate(-self.rotation, expand=True)
                        
                        if img.mode in ('RGBA', 'LA', 'P'):
                            img = img.convert('RGB')
                        
                        # 保存时使用最高质量
                        img.save(save_file_path, quality=100, subsampling=0)
            
            # 重置旋转状态
            self.rotation = 0
            
            # 切换到下一张图片
            total = len(self.images)
            if total == 1:
                # 只有一张图片，保持当前
                return {
                    "success": True,
                    "message": f"图片已保存",
                    "path": save_file_path,
                    "finished": False,
                    "nextImage": self._get_current_image_data()
                }
            else:
                # 切换到下一张
                self.current_index = (self.current_index + 1) % total
                next_image_data = self._get_current_image_data()
                
                return {
                    "success": True,
                    "message": f"图片已保存，当前第 {self.current_index + 1}/{total} 张",
                    "path": save_file_path,
                    "finished": False,
                    "nextImage": next_image_data
                }
        except Exception as e:
            return {"success": False, "message": f"保存失败: {str(e)}"}


def _make_ocr_body_and_sign(image_base64: str) -> Tuple[str, str]:
    """
    构建 OCR 请求体和签名
    
    参数:
        image_base64: 图片的 base64 编码字符串
    
    返回:
        (请求体JSON字符串, 签名)
    """
    body_obj = {
        "images": [
            {
                "data": image_base64,
                "dataId": "1",
                "type": 2
            }
        ],
        "nonce": random.randint(0, int(1e5)),
        "secretId": "Inner_40731a6efece4c2e992c0d670222e6da",
        "timestamp": int(time.time() * 1000)
    }
    body = json.dumps(body_obj)
    text = body + '43e7a66431b14c8f856a8e889070c19b'
    sign = hashlib.md5(text.encode('utf-8')).hexdigest()
    return body, sign


def perform_ocr(image: Image.Image) -> str:
    """
    执行OCR识别（使用超星OCR API）
    
    参数:
        image: PIL Image对象，需要进行OCR识别的图片
    
    返回:
        识别出的文字内容
    """
    try:
        # 将 PIL Image 转换为 base64
        buffer = BytesIO()
        # 确保是 RGB 模式
        if image.mode in ('RGBA', 'LA', 'P'):
            image = image.convert('RGB')
        image.save(buffer, format='PNG')
        image_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        # 构建请求体和签名
        body, sign = _make_ocr_body_and_sign(image_base64)
        
        # 发送请求
        headers = {
            'CX-Signature': sign,
            'Content-Type': 'application/json;charset=utf-8'
        }
        
        response = requests.post(
            'http://ai.chaoxing.com/api/v1/ocr/common/sync',
            data=body,
            headers=headers,
            timeout=30
        )
        
        # 解析响应
        if response.status_code == 200:
            data = response.json()
            if 'data' in data and len(data['data']) > 0:
                result = ''.join(item['text'] for item in data['data'][0])
                return result.strip()
            else:
                return ""
        else:
            print(f"OCR请求失败: {response.status_code}")
            return ""
            
    except requests.exceptions.Timeout:
        print("OCR请求超时")
        return ""
    except Exception as e:
        print(f"OCR识别错误: {str(e)}")
        return ""


def main():
    api = Api()
    
    # 获取HTML文件路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    html_path = os.path.join(current_dir, 'index.html')
    
    # 检查HTML文件是否存在
    if not os.path.exists(html_path):
        print(f"错误: 找不到HTML文件: {html_path}")
        return
    
    print(f"正在加载: {html_path}")
    
    # 创建窗口
    window = webview.create_window(
        '图纸重命名工具',
        html_path,
        js_api=api,
        width=1400,
        height=985,
        min_size=(1000, 700),
        resizable=True
    )
    
    api.set_window(window)
    
    # 启动应用
    # 尝试不同的后端: 'edgechromium' (默认), 'mshtml', 'cef'
    # 如果假死，可以尝试将 gui 参数改为 'mshtml'
    webview.start(debug=False, gui='qt')


if __name__ == "__main__":
    main()
