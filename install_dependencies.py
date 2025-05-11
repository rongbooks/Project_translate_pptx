import subprocess
import importlib
import sys

# 定义项目所需的库列表
REQUIRED_LIBS = {
    "python-pptx": "python-pptx",
    "requests": "requests",
    "pywin32": "pywin32"  # 可选：如果需要处理Windows特定功能（如文件关联）
}

def check_install_lib(lib_name, pip_name):
    """检查并安装库"""
    try:
        importlib.import_module(lib_name)
        print(f"✅ {lib_name} 已安装")
    except ImportError:
        print(f"❌ 未找到 {lib_name}，开始安装...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
            print(f"✅ {lib_name} 安装成功")
        except subprocess.CalledProcessError:
            print(f"❌ 安装 {lib_name} 失败，请手动安装：pip install {pip_name}")

def main():
    print("开始检查并安装依赖库...")
    for lib, pip_name in REQUIRED_LIBS.items():
        check_install_lib(lib, pip_name)
    print("\n所有依赖库检查/安装完成！")
    print("可以开始运行PPT翻译程序了！")

if __name__ == "__main__":
    main()