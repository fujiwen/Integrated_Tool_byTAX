import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import subprocess
from datetime import datetime
import importlib.util
import logging

def get_config_path():
    """获取配置文件路径"""
    return os.path.join(get_base_path(), "config.txt")

def ensure_config_file():
    """确保配置文件存在，如果不存在则创建"""
    config_path = get_config_path()
    if not os.path.exists(config_path):
        default_config = '''B2:海口索菲特大酒店
D2:海南省海口市龙华区滨海大道105号
E2:符小瑜 0898-31289999
B32:abbyfu@hksft.com
hotelname:海口索菲特大酒店
Sheet_tittle:供货明细表'''
        with open(config_path, 'w', encoding='utf-8') as f:
            f.write(default_config)
        return True
    return False

class IntegratedTool:
    def __init__(self, root):
        self.root = root
        
        # 首先检查过期时间
        check_expiration_time()
        
        self.root.title("供应商对账工具集byTAX")
        
        # 设置窗口大小并居中
        self.set_window_geometry(400, 400)
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题标签
        title_label = ttk.Label(self.main_frame, text="请选择要使用的功能：", font=("微软雅黑", 12))
        title_label.pack(pady=20)
        
        # 创建按钮框架
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(pady=20)
        
        # 创建供应商供货明细表工具按钮
        self.recon_btn = ttk.Button(
            button_frame,
            text="供应商供货明细表工具",
            command=self.launch_recon_tool,
            width=25
        )
        self.recon_btn.pack(pady=10)
        
        # 创建供应商对帐确认函按钮
        self.classification_btn = ttk.Button(
            button_frame,
            text="供应商对帐确认函byTAX",
            command=self.launch_classification_tool_byTAX,
            width=25
        )
        self.classification_btn.pack(pady=10)
        
        # 添加开发者信息
        self.dev_label = ttk.Label(self.main_frame, text="Powered By Cayman Fu @ Sofitel HAIKOU 2025 Ver 2.4")
        self.dev_label.pack(side=tk.BOTTOM, pady=10)
    
    def set_window_geometry(self, width, height):
        """设置窗口大小并居中"""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def _import_module(self, module_name):
        """导入模块，支持打包环境"""
        try:
            # 首先尝试从当前目录导入
            module_path = os.path.join(get_base_path(), f"{module_name}.py")
            if os.path.exists(module_path):
                logging.info(f"从路径导入模块：{module_path}")
                spec = importlib.util.spec_from_file_location(module_name, module_path)
                module = importlib.util.module_from_spec(spec)
                sys.modules[module_name] = module  # 将模块添加到 sys.modules
                spec.loader.exec_module(module)
                return module
            else:
                # 如果文件不存在，尝试直接导入
                logging.info(f"尝试直接导入模块：{module_name}")
                return __import__(module_name)
        except Exception as e:
            error_msg = f"导入模块 {module_name} 失败：{str(e)}"
            logging.error(error_msg, exc_info=True)
            raise ImportError(error_msg)
    
    def launch_recon_tool(self):
        """启动供应商供货明细表工具"""
        try:
            logging.info("正在启动供应商供货明细表工具")
            Bldbuy_Recon_UI = self._import_module("Bldbuy_Recon_ByTAX")
            logging.info("成功导入 Bldbuy_Recon_ByTAX 模块")
            
            # 创建新窗口，并设置为主窗口的子窗口
            child_window = tk.Toplevel(self.root)
            child_window.withdraw()  # 先隐藏窗口
            
            # 设置窗口标题和图标
            child_window.title("供应商供货明细表工具")
            try:
                favicon_path = os.path.join(get_base_path(), "favicon.ico")
                if os.path.exists(favicon_path):
                    child_window.iconbitmap(favicon_path)
            except Exception as e:
                logging.warning(f"设置图标失败：{str(e)}")
            
            logging.info("正在初始化供应商供货明细表工具界面")
            app = Bldbuy_Recon_UI.BldBuyApp(child_window)
            
            # 显示窗口并设置焦点
            child_window.deiconify()
            child_window.lift()
            child_window.focus_force()
            
            logging.info("供应商供货明细表工具启动成功")
            child_window.mainloop()
            
        except Exception as e:
            error_msg = f"启动供应商供货明细表工具失败：{str(e)}"
            logging.error(error_msg, exc_info=True)
            messagebox.showerror("错误", error_msg)
    
    def launch_classification_tool_byTAX(self):
        """启动供应商对帐确认函byTAX"""
        try:
            logging.info("正在启动供应商对帐确认函")
            Product_Classification_Tool = self._import_module("Product_Classification_Tool_ByTAX")
            logging.info("成功导入 Product_Classification_Tool_ByTAX 模块")
            
            # 创建新窗口，并设置为主窗口的子窗口
            child_window = tk.Toplevel(self.root)
            child_window.withdraw()  # 先隐藏窗口
            
            # 设置窗口标题和图标
            child_window.title("供应商对帐确认函")
            try:
                favicon_path = os.path.join(get_base_path(), "favicon.ico")
                if os.path.exists(favicon_path):
                    child_window.iconbitmap(favicon_path)
            except Exception as e:
                logging.warning(f"设置图标失败：{str(e)}")
            
            logging.info("正在初始化供应商对帐确认函界面")
            app = Product_Classification_Tool.ProductClassificationApp(child_window)
            
            # 显示窗口并设置焦点
            child_window.deiconify()
            child_window.lift()
            child_window.focus_force()
            
            logging.info("供应商对帐确认函启动成功")
            child_window.mainloop()
            
        except Exception as e:
            error_msg = f"启动供应商对帐确认函失败：{str(e)}"
            logging.error(error_msg, exc_info=True)
            messagebox.showerror("错误", error_msg)
    
def check_expiration_time():
    """检查时间是否到期"""
    current_date = datetime.now()
    expiration_date = datetime(2025, 12, 31)  # 2025年底到期
    
    if current_date > expiration_date:
        messagebox.showerror("错误", "DLL注册失败，请联系Cayman更新")
        sys.exit(1)

def get_base_path():
    """获取程序运行时的基础路径"""
    if getattr(sys, 'frozen', False):
        # 如果是打包后的程序
        return os.path.dirname(sys.executable)
    else:
        # 如果是开发环境
        return os.path.dirname(os.path.abspath(__file__))

if __name__ == "__main__":
    try:
        # 设置日志文件
        base_path = get_base_path()
        log_path = os.path.join(base_path, 'error.log')
        logging.basicConfig(
            filename=log_path,
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
        logging.info(f"程序基础路径：{base_path}")
        
        # 首先检查过期时间
        logging.info("正在检查过期时间")
        check_expiration_time()
        
        # 检查并确保配置文件存在
        logging.info("正在检查配置文件")
        if ensure_config_file():
            messagebox.showinfo("提示", "已创建默认配置文件：config.txt")
        
        logging.info("正在初始化主窗口")
        root = tk.Tk()
        app = IntegratedTool(root)
        logging.info("程序启动成功")
        root.mainloop()
    except Exception as e:
        error_msg = f"程序启动失败：{str(e)}"
        print(error_msg)
        try:
            logging.error(error_msg, exc_info=True)
        except:
            pass
        sys.exit(1)