import os
import sys
import base64
import time
import json
import pandas as pd
from volcenginesdkarkruntime import Ark
from pathlib import Path
from dotenv import load_dotenv

# 加载.env文件
load_dotenv()

def get_project_root():
    """获取项目根目录，兼容开发环境和PyInstaller打包环境"""
    try:
        # 检查是否有从main.py传递过来的project_root变量
        if 'project_root' in globals():
            return globals()['project_root']
        # PyInstaller创建一个临时文件夹并将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
        # 在打包环境中，可执行文件通常位于dist目录中
        # 我们需要获取可执行文件所在目录的上级目录作为项目根目录
        executable_dir = os.path.dirname(sys.executable)
        project_root = os.path.dirname(executable_dir)
    except Exception:
        # 检查是否有从main.py传递过来的project_root变量
        if 'project_root' in globals():
            return globals()['project_root']
        # 开发环境中，向上两级到达项目根目录
        # Release/Image Title Generation/Title Generation.py
        script_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(os.path.dirname(script_dir))
    return project_root

# 获取项目根目录
project_root = get_project_root()
print(f"Title Generation 脚本 - 项目根目录: {project_root}")
print(f"当前工作目录: {os.getcwd()}")
print(f"__file__ 属性: {__file__ if '__file__' in globals() else '未设置'}")
print(f"project_root 全局变量: {'存在' if 'project_root' in globals() else '不存在'}")

# 读取Release目录下的配置文件获取图片文件夹路径和结果文件夹路径
# config_file = Path(project_root) / "Release" / "config.txt"
config_file = Path(project_root)  / "config.txt"
print(f"配置文件路径: {config_file}")
IMAGE_FOLDER_PATH = ""
RESULT_FOLDER_PATH = ""
PARENT_CLASS_GROUP_SIZE = 2  # 默认值为2
MODEL_NAME = "doubao-seed-1-6-251015"  # 默认模型名称
if config_file.exists():
    print("配置文件存在，正在读取...")
    with open(config_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                key, value = line.split('=', 1)
                if key.strip() == 'IMAGE_FOLDER_PATH':
                    IMAGE_FOLDER_PATH = value.strip()
                    print(f"读取到图片文件夹路径: {IMAGE_FOLDER_PATH}")
                elif key.strip() == 'RESULT_FOLDER_PATH':
                    RESULT_FOLDER_PATH = value.strip()
                    print(f"读取到结果文件夹路径: {RESULT_FOLDER_PATH}")
                elif key.strip() == 'PARENT_CLASS_GROUP_SIZE':
                    try:
                        PARENT_CLASS_GROUP_SIZE = int(value.strip())
                        print(f"读取到父类分组大小: {PARENT_CLASS_GROUP_SIZE}")
                    except ValueError:
                        print(f"警告: PARENT_CLASS_GROUP_SIZE配置值无效，使用默认值2")
                elif key.strip() == 'MODEL_NAME':
                    MODEL_NAME = value.strip()
                    print(f"读取到模型名称: {MODEL_NAME}")

if not IMAGE_FOLDER_PATH or not os.path.exists(IMAGE_FOLDER_PATH):
    raise ValueError(f"无效的图片文件夹路径: {IMAGE_FOLDER_PATH}")

# 如果配置文件中没有指定结果路径，则使用默认路径
if not RESULT_FOLDER_PATH:
    RESULT_FOLDER_PATH = str(Path(project_root) / "Result")
    print(f"未在配置文件中找到结果文件夹路径，使用默认路径: {RESULT_FOLDER_PATH}")
else:
    print(f"使用配置文件中的结果文件夹路径: {RESULT_FOLDER_PATH}")

# 读取提示词
prompt_file = Path(project_root) / "prompt.txt"
if not prompt_file.exists():
    raise FileNotFoundError(f"找不到提示词文件: {prompt_file}")

with open(prompt_file, 'r', encoding='utf-8') as f:
    prompt_content = f.read().strip()

# 配置API密钥
DOUBAO_API_KEY = os.getenv('DOUBAO_API_KEY')
if not DOUBAO_API_KEY:
    raise ValueError("请在.env文件中设置DOUBAO_API_KEY")

# 初始化豆包模型客户端
client = Ark(
    api_key=DOUBAO_API_KEY,
    base_url="https://ark.cn-beijing.volces.com/api/v3",
)

# 设置结果目录和失败目录
result_dir = Path(RESULT_FOLDER_PATH)
failure_dir = Path(project_root) / "Failure"

# 创建必要的目录
result_dir.mkdir(parents=True, exist_ok=True)
failure_dir.mkdir(parents=True, exist_ok=True)

# 将本地图片转换为base64编码
def image_to_base64(image_path):
    with open(image_path, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
        # 获取文件扩展名
        _, ext = os.path.splitext(image_path)
        # 根据文件扩展名确定MIME类型
        mime_types = {
            '.jpg': 'jpeg',
            '.jpeg': 'jpeg',
            '.png': 'png',
            '.bmp': 'bmp',
            '.gif': 'gif',
            '.tiff': 'tiff',
            '.webp': 'webp'
        }
        mime_type = mime_types.get(ext.lower(), 'jpeg')  # 默认为jpeg
        return f"data:image/{mime_type};base64,{encoded_string}"

def extract_json_from_response(response_text):
    """从响应文本中提取JSON内容"""
    # 查找第一个 '{' 和最后一个 '}'
    start = response_text.find('{')
    end = response_text.rfind('}') + 1
    
    if start != -1 and end > start:
        json_str = response_text[start:end]
        try:
            # 验证是否为有效的JSON
            parsed = json.loads(json_str)
            return json_str, parsed
        except json.JSONDecodeError:
            pass
    
    return None, None

def main():
    """主函数"""
    print("开始处理图片...")
    
    # 支持的图片格式
    image_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')
    
    # 获取所有图片文件
    image_files = []
    added_files = set()
    for ext in image_extensions:
        # 添加小写扩展名的文件
        for file_path in Path(IMAGE_FOLDER_PATH).glob(f"*{ext}"):
            if file_path.name.lower() not in added_files:
                image_files.append(file_path)
                added_files.add(file_path.name.lower())
        
        # 添加大写扩展名的文件
        for file_path in Path(IMAGE_FOLDER_PATH).glob(f"*{ext.upper()}"):
            if file_path.name.lower() not in added_files:
                image_files.append(file_path)
                added_files.add(file_path.name.lower())
    
    if not image_files:
        print(f"在 {IMAGE_FOLDER_PATH} 中没有找到支持的图片文件")
        return
    
    print(f"找到 {len(image_files)} 个图片文件")
    
    # 存储结果和失败记录
    results = []
    failed_images = []
    
    # 处理每个图片
    success_count = 0
    for idx, image_file in enumerate(image_files, 1):
        try:
            print(f"[{idx}/{len(image_files)}] 正在处理: {image_file.name}")
            
            # 将本地图片转换为base64编码
            image_base64 = image_to_base64(image_file)
            
            # 调用豆包模型
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": image_base64
                                },
                            },
                            {"type": "text", "text": prompt_content},
                        ],
                    }
                ],
            )
            
            # 获取生成的内容
            generated_content = response.choices[0].message.content.strip()
            
            # 解析JSON响应
            try:
                # 从响应中提取JSON（假设它被包装在一些文本中）
                # 找到第一个'{'和最后一个'}'来提取JSON部分
                start_idx = generated_content.find('{')
                end_idx = generated_content.rfind('}') + 1
                
                if start_idx != -1 and end_idx > start_idx:
                    json_str = generated_content[start_idx:end_idx]
                    parsed_data = json.loads(json_str)
                    
                    # 提取四个字段
                    amazon_title = parsed_data.get("amazon_title", "")
                    amazon_title_translation = parsed_data.get("amazon_title_translation", "")
                    short_title = parsed_data.get("short_title", "")
                    short_title_translation = parsed_data.get("short_title_translation", "")
                else:
                    # 如果未找到JSON，则将整个内容作为amazon_title使用
                    amazon_title = generated_content
                    amazon_title_translation = ""
                    short_title = ""
                    short_title_translation = ""
            except json.JSONDecodeError:
                # 如果JSON解析失败，则将整个内容作为amazon_title使用
                amazon_title = generated_content
                amazon_title_translation = ""
                short_title = ""
                short_title_translation = ""
            
            # 移除文件扩展名
            image_name_without_extension = image_file.stem
            
            # 添加到结果列表
            results.append({
                '图片名称': image_name_without_extension,
                '亚马逊产品标题': amazon_title,
                '亚马逊产品标题翻译': amazon_title_translation,
                '短标题': short_title,
                '短标题翻译': short_title_translation
            })
            
            print(f"✓ 处理成功: {image_file.name}")
            success_count += 1
            
        except Exception as e:
            print(f"✗ 处理失败 {image_file.name}: {str(e)}")
            # 移除文件扩展名
            image_name_without_extension = image_file.stem
            
            # 记录失败的图片到失败列表
            failed_images.append({
                '图片名称': image_name_without_extension,
                '错误信息': str(e)
            })
    
    # 创建DataFrame
    df = pd.DataFrame(results)

    # 计算父类编号
    parent_class_ids = []
    image_names = df['图片名称'].tolist()
    
    # 根据PARENT_CLASS_GROUP_SIZE分组计算父类编号
    for i in range(len(image_names)):
        # 计算当前行所属的组索引
        group_index = i // PARENT_CLASS_GROUP_SIZE
        # 计算该组的第一个元素索引
        first_index_in_group = group_index * PARENT_CLASS_GROUP_SIZE
        # 使用该组第一个图片的名称作为父类编号
        parent_class_id = image_names[first_index_in_group]
        parent_class_ids.append(parent_class_id)
    
    # 将父类编号添加到DataFrame中
    df['父类编号'] = parent_class_ids

    # 重新排列列顺序，将父类编号列放在F列位置（即索引为1的位置）
    df = df[['图片名称', '父类编号', '亚马逊产品标题', '亚马逊产品标题翻译', '短标题', '短标题翻译']]

    # 输出Excel文件路径
    output_excel = result_dir / "Image_Titles_Doubao.xlsx"

    # 检查文件是否存在，如果存在则重命名
    final_output_excel = output_excel
    counter = 1
    while final_output_excel.exists():
        # 获取文件名和扩展名
        name, ext = final_output_excel.stem, final_output_excel.suffix
        # 在文件名中添加计数器
        final_output_excel = result_dir / f"{name}_{counter}{ext}"
        counter += 1

    # 保存到Excel文件
    df.to_excel(final_output_excel, index=False, engine='openpyxl')
    print(f"所有图片处理完成。结果已保存到: {final_output_excel}")
    print(f"总共处理了: {len(results)} 张图片")

    # 如果有失败的图片，保存到失败记录文件
    if failed_images:
        failure_excel = failure_dir / "Failed_Images.xlsx"
        failure_df = pd.DataFrame(failed_images)
        failure_df.to_excel(failure_excel, index=False, engine='openpyxl')
        print(f"发现 {len(failed_images)} 张图片处理失败，失败记录已保存到: {failure_excel}")
    else:
        print("所有图片都处理成功，没有失败记录。")
    
    # 返回成功和失败的数量
    return len(results), len(failed_images)

if __name__ == "__main__":
    main()