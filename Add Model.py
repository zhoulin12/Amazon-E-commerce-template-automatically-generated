import os
import sys
from pathlib import Path
import pandas as pd

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
        # Release/Image Title Generation/Add Model.py
        script_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(os.path.dirname(script_dir))
    return project_root

def get_result_folder_from_config(config_file):
    """从配置文件中读取结果文件夹路径"""
    result_folder = None
    if os.path.exists(config_file):
        with open(config_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    if key.strip() == 'RESULT_FOLDER_PATH':
                        result_folder = value.strip()
                        break
    
    # 如果配置文件中没有指定结果文件夹路径，则使用默认路径
    if not result_folder:
        project_root = get_project_root()
        result_folder = os.path.join(project_root, "Result")
    
    # 确保结果文件夹存在
    os.makedirs(result_folder, exist_ok=True)
    return result_folder

def main():
    # 获取项目根目录
    project_root = get_project_root()
    print(f"Add Model 脚本 - 项目根目录: {project_root}")
    
    # 从配置文件读取结果文件夹路径
    # config_file = Path(project_root) / "Release" / "config.txt"
    config_file = Path(project_root)  / "config.txt"
    result_dir = get_result_folder_from_config(config_file)
    
    # 输入文件和模型文件路径
    input_excel = Path(result_dir) / "Image_Titles_Doubao.xlsx"
    model_excel = Path(project_root) / "需要的excel文件" / "型号.xlsx"
    
    # 输出文件路径
    output_excel = Path(result_dir) / "Image_Titles_Add_Model.xlsx"
    
    # 检查输入文件是否存在
    if not input_excel.exists():
        print(f"输入文件不存在: {input_excel}")
        return
    
    # 检查模型文件是否存在
    if not model_excel.exists():
        print(f"模型文件不存在: {model_excel}")
        return
    
    print(f"正在读取输入文件: {input_excel}")
    print(f"正在读取模型文件: {model_excel}")
    
    try:
        # 读取两个Excel文件
        df_input = pd.read_excel(input_excel)
        df_model = pd.read_excel(model_excel)
        
        print(f"输入文件包含 {len(df_input)} 行数据")
        print(f"模型文件包含 {len(df_model)} 行数据")
        
        # 检查必要的列是否存在
        required_input_columns = ['图片名称', '父类编号', '亚马逊产品标题', '亚马逊产品标题翻译', '短标题', '短标题翻译']
        missing_input_columns = [col for col in required_input_columns if col not in df_input.columns]
        if missing_input_columns:
            print(f"输入文件缺少必要的列: {missing_input_columns}")
            return
        
        required_model_columns = ['手机型号', '尺寸']
        missing_model_columns = [col for col in required_model_columns if col not in df_model.columns]
        if missing_model_columns:
            print(f"模型文件缺少必要的列: {missing_model_columns}")
            return
        
        # 创建结果列表
        results = []
        
        total_combinations = len(df_input) * len(df_model)
        current_combination = 0
        
        # 对于输入文件中的每一行，与模型文件中的所有行组合
        for idx_input, row_input in df_input.iterrows():
            image_name = row_input['图片名称']
            parent_class_id = row_input['父类编号']  # 读取父类编号
            amazon_title = row_input['亚马逊产品标题']
            amazon_title_translation = row_input['亚马逊产品标题翻译']
            short_title = row_input['短标题']
            short_title_translation = row_input['短标题翻译']
            
            # 与模型文件中的每一行组合
            for idx_model, row_model in df_model.iterrows():
                current_combination += 1
                model_name = row_model['手机型号']
                size = row_model['尺寸']
                
                # 组合图片名称和手机型号（移除型号中的空格）
                # 示例: iPhone X -> iPhoneX
                model_name_no_space = model_name.replace(" ", "")
                combined_image_name = f"{image_name}{model_name_no_space}"
                
                # 替换标题中的手机型号信息
                # 标题格式应该是: for iPhone X Case 5.8 inch ...
                # 我们需要替换标题中的型号信息以匹配特定型号并添加尺寸
                
                # 从型号中提取基本型号名称（移除"iPhone"前缀）
                base_model = model_name.replace("iPhone", "").strip()
                
                # 构造新标题
                if pd.isna(amazon_title) or str(amazon_title).strip() == "":
                    # 如果原标题为空，构造一个新标题
                    new_title = f"for iPhone {base_model} Case {size} Premium Quality Protective Phone Case with Stylish Design for Daily Use"
                elif "for iPhone" in str(amazon_title):
                    # 对于现有标题，确保包含尺寸信息且格式正确
                    # 首先提取原标题的主要内容（移除"for iPhone"部分）
                    title_content = str(amazon_title).replace("for iPhone", "", 1).strip()
                    
                    # 检查是否已经存在型号信息
                    models_list = [model.replace("iPhone", "").strip() for model in df_model['手机型号'].unique()]
                    starts_with_model = False
                    model_prefix = ""
                    for model in models_list:
                        if title_content.startswith(model):
                            starts_with_model = True
                            model_prefix = model
                            break
                    
                    if starts_with_model:
                        # 如果已经存在型号信息，移除它
                        title_content = title_content[len(model_prefix):].strip()
                    
                    # 移除多余的"Case"关键字
                    title_content = title_content.replace("Case", "", 1).strip()
                    
                    # 构造带有型号和尺寸的新标题，格式: for iPhone X Case 5.8 inch ...
                    new_title = f"for iPhone {base_model} Case {size} {title_content}"
                else:
                    # 如果标题不包含"for iPhone"，添加完整信息
                    new_title = f"for iPhone {base_model} Case {size} {str(amazon_title)}"
                
                # 添加到结果列表
                results.append({
                    '图片编号': image_name,  # 添加图片编号列，值为原始图片名称
                    '型号': model_name,  # 添加型号列，值为手机型号
                    '父类编号': parent_class_id,  # 添加父类编号列
                    '图片名称': combined_image_name,
                    '亚马逊产品标题': new_title,
                    '亚马逊产品标题翻译': amazon_title_translation,  # 保留翻译
                    '短标题': short_title,  # 保留短标题
                    '短标题翻译': short_title_translation  # 保留短标题翻译
                })
                
                # 每处理100个组合显示一次进度
                if current_combination % 100 == 0:
                    print(f"进度: {current_combination}/{total_combinations} 个组合已处理")
        
        # 创建结果DataFrame
        df_result = pd.DataFrame(results)
        
        # 重新排列列顺序
        desired_order = ['图片名称', '父类编号', '亚马逊产品标题', '亚马逊产品标题翻译', '短标题', '短标题翻译', '图片编号', '型号']
        # 确保所有需要的列都在
        available_cols = [col for col in desired_order if col in df_result.columns]
        # 添加任何不在期望顺序中的剩余列
        remaining_cols = [col for col in df_result.columns if col not in available_cols]
        final_order = available_cols + remaining_cols
        df_result = df_result[final_order]
        
        # 显示一些结果示例
        print("处理完成。前5行结果:")
        print(df_result.head())
        
        # 检查文件是否存在，如果存在则重命名
        final_output_excel = output_excel
        counter = 1
        while final_output_excel.exists():
            # 获取文件名和扩展名
            name, ext = final_output_excel.stem, final_output_excel.suffix
            # 在文件名中添加计数器
            final_output_excel = Path(final_output_excel.parent) / f"{name}_{counter}{ext}"
            counter += 1
        
        # 保存结果到Excel文件
        df_result.to_excel(final_output_excel, index=False, engine='openpyxl')
        print(f"结果已保存到: {final_output_excel}")
        print(f"总共生成: {len(df_result)} 行")
        
    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()