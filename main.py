import os
import sys
import shutil
import importlib.util
import traceback
from pathlib import Path
from datetime import datetime

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def is_console_available():
    """Check if console is available (not in windowed mode)"""
    return sys.stdin and sys.stdin.isatty()

def safe_input(prompt):
    """Safe input that works in both console and windowed modes"""
    if is_console_available():
        try:
            return input(prompt)
        except RuntimeError:
            # In windowed mode, just return None
            return None
    return None

# 获取项目根目录的正确方法
def get_project_root():
    """获取项目根目录，兼容开发环境和PyInstaller打包环境"""
    try:
        # PyInstaller创建一个临时文件夹并将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
        # 在打包环境中，可执行文件通常位于dist目录中
        # 我们需要获取可执行文件所在目录的上级目录作为项目根目录
        executable_dir = os.path.dirname(sys.executable)
        project_root = os.path.dirname(executable_dir)
    except Exception:
        # 开发环境中，向上一级到达项目根目录
        # Release/main.py
        script_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(script_dir)
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
    # 记录开始时间
    start_time = datetime.now()
    print(f"程序开始运行: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Define directories
    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    image_gen_dir = script_dir / "Image Title Generation"
    
    print(f"脚本目录: {script_dir}")
    print(f"图片处理脚本目录: {image_gen_dir}")
    
    # 使用正确的项目根目录
    project_root = get_project_root()
    print(f"项目根目录: {project_root}")
    
    # 从配置文件读取结果文件夹路径
    config_file = Path(project_root) / "config.txt"
    result_dir = get_result_folder_from_config(config_file)
    print(f"结果目录: {result_dir}")
    
    # List of scripts to run in order
    scripts_to_run = [
        "Add Model.py",
        "Ultimately.py",
        "Organize.py"
    ]
    
    # 用于存储统计信息
    stats = None
    
    # Check if config.txt exists
    print(f"检查配置文件: {config_file}")
    if not config_file.exists():
        print("错误: 找不到配置文件 config.txt")
        safe_input("按任意键退出...")
        return
    
    # 显示配置文件内容
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            print("配置文件内容:")
            for line in f:
                print(f"  {line.strip()}")
    except Exception as e:
        print(f"读取配置文件时出错: {e}")
    
    # Run each script in order
    total_scripts = len(scripts_to_run)
    for i, script_name in enumerate(scripts_to_run, 1):
        script_path = image_gen_dir / script_name
        print(f"\n[{i}/{total_scripts}] 正在运行: {script_name}")
        print("-" * 50)
        
        if script_path.exists():
            try:
                script_start_time = datetime.now()
                print(f"开始时间: {script_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
                
                # Import and call main function
                spec = importlib.util.spec_from_file_location(
                    script_name.replace(".py", ""),
                    script_path
                )
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                
                # Call main function if it exists
                if hasattr(module, 'main') and callable(getattr(module, 'main')):
                    module.main()
                
                script_end_time = datetime.now()
                script_duration = script_end_time - script_start_time
                print(f"完成时间: {script_end_time.strftime('%Y-%m-%d %H:%M:%S')}")
                print(f"执行耗时: {script_duration}")
                print(f"完成: {script_name}\n")
                
            except Exception as e:
                print(f"错误: 运行 {script_name} 时发生异常:")
                print(f"错误类型: {type(e).__name__}")
                print(f"错误信息: {str(e)}")
                print("详细错误追踪:")
                traceback.print_exc()
                print("\n程序将继续执行下一个脚本...\n")
        else:
            print(f"警告: 找不到脚本 {script_name}，跳过...")
    
    end_time = datetime.now()
    total_duration = end_time - start_time
    print("=" * 50)
    print(f"所有任务已完成!")
    print(f"开始时间: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"结束时间: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"总耗时: {total_duration}")
    
    # 显示统计信息
    # 已移除 Title Generation 的统计信息显示
    
    # 只在控制台模式下等待用户输入
    if is_console_available():
        safe_input("按任意键退出...")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"程序运行时发生未处理的异常:")
        print(f"错误类型: {type(e).__name__}")
        print(f"错误信息: {str(e)}")
        print("详细错误追踪:")
        traceback.print_exc()
        
        # 只在控制台模式下等待用户输入
        if is_console_available():
            safe_input("按任意键退出...")