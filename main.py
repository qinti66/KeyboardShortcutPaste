import tkinter as tk
from tkinter import ttk, messagebox
import pyautogui
import pyperclip
import keyboard
import os
import sys
import json
import time

def get_data_file_path():
    """获取数据文件的路径 (兼容打包后的路径)"""
    if getattr(sys, 'frozen', False):
        program_dir = os.path.dirname(sys.executable)
        return os.path.join(program_dir, "config.json")
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

class QuickPasteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("自动化快捷粘贴工具 Pro")
        self.root.geometry("650x550")

        # 核心数据结构
        self.rows = []           # 存储每一行的控件引用
        self.hotkey_hooks = {}   # 存储 keyboard 的 hook 对象: {hotkey_name: hook}
        self.hotkey_mapping = {} # 存储占用映射: {hotkey_name: row_index}
        self.active_keys = {}    # 存储当前行绑定的键: {row_index: hotkey_name}

        self.init_ui()
        self.load_data()
        
        # 启动“防掉线”守护进程：每15分钟刷新一次钩子
        self.keep_alive()

    def init_ui(self):
        """初始化基础布局"""
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=20, pady=10)

        ttk.Label(top_frame, text="输入框数量 (1-12):").pack(side="left")
        self.input_num_entry = ttk.Entry(top_frame, width=5)
        self.input_num_entry.pack(side="left", padx=5)
        self.input_num_entry.insert(0, "5")

        ttk.Button(top_frame, text="生成/重置布局", command=self.generate_widgets).pack(side="left", padx=5)
        ttk.Button(top_frame, text="全部清空", command=self.clear_all).pack(side="left", padx=5)

        # 滚动区域
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=5)

    def generate_widgets(self):
        """动态生成行"""
        try:
            num = int(self.input_num_entry.get())
        except:
            messagebox.showerror("错误", "请输入有效的数字")
            return

        if not (1 <= num <= 12):
            messagebox.showerror("错误", "数量请保持在 1-12 之间")
            return

        # 清理旧插件
        self.clear_all_hooks()
        for row in self.rows:
            row['frame'].destroy()
        self.rows.clear()

        # 生成新行
        for i in range(num):
            row_frame = ttk.Frame(self.main_frame)
            row_frame.pack(fill="x", pady=2)

            ttk.Label(row_frame, text=f"{i+1}:", width=3).pack(side="left")
            
            # 内容输入
            ent = ttk.Entry(row_frame)
            ent.pack(side="left", fill="x", expand=True, padx=5)
            ent.bind('<KeyRelease>', lambda e, idx=i: self.update_trigger(idx))

            # 快捷键选择
            cb = ttk.Combobox(row_frame, width=8, state="readonly")
            cb['values'] = ['无'] + [f'f{j}' for j in range(1, 13)]
            cb.set('无')
            cb.pack(side="left", padx=5)
            cb.bind('<<ComboboxSelected>>', lambda e, idx=i: self.update_trigger(idx))

            self.rows.append({'frame': row_frame, 'entry': ent, 'combo': cb})

    def update_trigger(self, index):
        """当UI内容或键位改变时触发"""
        new_key = self.rows[index]['combo'].get()
        
        # 1. 冲突检查
        if new_key != '无':
            # 如果这个键被别人占用了
            if new_key in self.hotkey_mapping and self.hotkey_mapping[new_key] != index:
                messagebox.showwarning("冲突", f"快捷键 {new_key} 已被第 {self.hotkey_mapping[new_key]+1} 行占用！")
                # 恢复之前的值
                old_val = self.active_keys.get(index, '无')
                self.rows[index]['combo'].set(old_val)
                return

        # 2. 重新注册
        self.register_single_hotkey(index)
        # 3. 保存
        self.save_data()

    def register_single_hotkey(self, index):
        """核心注册逻辑：先释放，再绑定"""
        # A. 释放该行之前绑定的任何按键
        old_key = self.active_keys.get(index)
        if old_key and old_key in self.hotkey_hooks:
            try:
                keyboard.remove_hotkey(self.hotkey_hooks[old_key])
            except: pass
            del self.hotkey_hooks[old_key]
            if old_key in self.hotkey_mapping:
                del self.hotkey_mapping[old_key]

        # B. 绑定新按键
        new_key = self.rows[index]['combo'].get()
        content = self.rows[index]['entry'].get()

        if new_key != '无' and content.strip():
            # 注册新钩子 (注意：去掉了 suppress=True 以增加稳定性)
            try:
                hook = keyboard.add_hotkey(new_key, lambda: self.smart_paste(content))
                self.hotkey_hooks[new_key] = hook
                self.hotkey_mapping[new_key] = index
                self.active_keys[index] = new_key
            except Exception as e:
                print(f"绑定失败: {e}")
        else:
            self.active_keys[index] = '无'

    def smart_paste(self, content):
        """执行粘贴操作"""
        try:
            pyperclip.copy(content)
            # 增加 50ms 延迟，防止剪贴板未同步导致粘贴旧内容或失效
            time.sleep(0.05)
            pyautogui.hotkey('ctrl', 'v')
        except Exception as e:
            print(f"粘贴执行异常: {e}")

    def clear_all_hooks(self):
        """注销所有全局热键"""
        keyboard.unhook_all()
        self.hotkey_hooks.clear()
        self.hotkey_mapping.clear()
        self.active_keys.clear()

    def clear_all(self):
        """清空界面和数据"""
        if messagebox.askyesno("确认", "确定要清空所有内容吗？"):
            for row in self.rows:
                row['entry'].delete(0, tk.END)
                row['combo'].set('无')
            self.clear_all_hooks()
            self.save_data()

    def save_data(self):
        """保存配置到 JSON 文件"""
        data = []
        for row in self.rows:
            data.append({
                'content': row['entry'].get(),
                'hotkey': row['combo'].get()
            })
        try:
            with open(get_data_file_path(), 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except: pass

    def load_data(self):
        """加载 JSON 配置"""
        path = get_data_file_path()
        if not os.path.exists(path):
            self.generate_widgets()
            return

        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 更新数量框并生成
            self.input_num_entry.delete(0, tk.END)
            self.input_num_entry.insert(0, str(len(data)))
            self.generate_widgets()

            # 填充数据并激活
            for i, item in enumerate(data):
                if i < len(self.rows):
                    self.rows[i]['entry'].insert(0, item['content'])
                    self.rows[i]['combo'].set(item['hotkey'])
                    self.register_single_hotkey(i)
        except:
            self.generate_widgets()

    def keep_alive(self):
        """防失效机制：每15分钟刷新一次 Hook 链"""
        print("正在刷新热键钩子以保持活性...")
        for i in range(len(self.rows)):
            self.register_single_hotkey(i)
        # 900000 毫秒 = 15 分钟
        self.root.after(900000, self.keep_alive)

if __name__ == "__main__":
    root = tk.Tk()
    app = QuickPasteApp(root)
    
    # 退出前保存最后的状态
    def on_closing():
        app.save_data()
        root.destroy()
        
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()