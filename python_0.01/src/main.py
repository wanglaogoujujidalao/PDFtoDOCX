import sys
import os
import threading
# 修复：无控制台模式下stdout重定向（核心）
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stdin is None:
    sys.stdin = open(os.devnull, 'r', encoding='utf-8')

# 正确配置tkinter中文显示（删除错误的rcParams）
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
# 解决tkinter按钮/标签中文乱码的正确方式
tk.Tk().option_add('*Font', 'Microsoft YaHei 10')

from pdf2docx import Converter

# 全局变量存储路径
pdf_path_global = ""
docx_path_global = ""

def select_pdf():
    """选择PDF文件，弹窗选择"""
    global pdf_path_global
    path = filedialog.askopenfilename(
        title="选择要转换的PDF文件",
        filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")],
        initialdir="./"
    )
    if path:
        pdf_path_global = path
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, path)
        # 自动生成docx保存路径
        global docx_path_global
        if path.lower().endswith(".pdf"):
            docx_path_global = path[:-4] + ".docx"
        else:
            docx_path_global = path + ".docx"
        entry_docx.delete(0, tk.END)
        entry_docx.insert(0, docx_path_global)

def select_docx():
    """选择DOCX保存路径，弹窗选择"""
    global docx_path_global
    path = filedialog.asksaveasfilename(
        title="选择DOCX保存位置",
        defaultextension=".docx",
        filetypes=[("DOCX文件", "*.docx"), ("所有文件", "*.*")],
        initialdir="./"
    )
    if path:
        docx_path_global = path
        entry_docx.delete(0, tk.END)
        entry_docx.insert(0, path)

def convert_single_pdf_gui(pdf_path, docx_path, page_choice, img_choice):
    """GUI版单个PDF转换，带参数传入"""
    start_page = 0
    end_page = None
    # 处理页码选择
    if page_choice == "自定义":
        try:
            start_page = int(entry_start.get().strip()) if entry_start.get().strip() else 0
            end_page = int(entry_end.get().strip()) if entry_end.get().strip() else None
        except ValueError:
            return False, "❌ 页码请输入有效数字！"
    # 处理图片选择
    parse_images = False if img_choice == "忽略图片(更快)" else True
    # 执行转换
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=start_page, end=end_page, parse_images=parse_images)
        cv.close()
        return True, f"✅ 转换完成！\n文件保存至：\n{docx_path}"
    except FileNotFoundError:
        return False, "❌ PDF文件不存在，请检查路径！"
    except Exception as e:
        return False, f"❌ 转换失败：\n{str(e)}"

def start_convert():
    """开始转换按钮触发函数，子线程避免卡死"""
    if not pdf_path_global or not docx_path_global:
        messagebox.showerror("错误", "请先选择PDF文件和保存路径！")
        return
    # 获取参数
    page_choice = combox_page.get()
    img_choice = combox_img.get()
    # 按钮置灰
    btn_convert.config(state=tk.DISABLED, text="转换中...")
    # 子线程执行
    t = threading.Thread(target=convert_thread, args=(pdf_path_global, docx_path_global, page_choice, img_choice))
    t.daemon = True
    t.start()

def convert_thread(pdf_path, docx_path, page_choice, img_choice):
    """转换子线程"""
    success, msg = convert_single_pdf_gui(pdf_path, docx_path, page_choice, img_choice)
    if success:
        messagebox.showinfo("成功", msg)
    else:
        messagebox.showerror("失败", msg)
    # 恢复按钮
    btn_convert.config(state=tk.NORMAL, text="开始转换")
    # 清空路径
    entry_pdf.delete(0, tk.END)
    entry_docx.delete(0, tk.END)
    global pdf_path_global, docx_path_global
    pdf_path_global = ""
    docx_path_global = ""

def create_gui():
    """创建GUI界面"""
    root = tk.Tk()
    root.title("PDF转DOCX工具 | 最终版")
    root.geometry("650x320")
    root.resizable(False, False)

    # 全局控件
    global entry_pdf, entry_docx, combox_page, combox_img, entry_start, entry_end, btn_convert

    # ========== PDF选择 ==========
    lab_pdf = tk.Label(root, text="📄 PDF源文件：")
    lab_pdf.place(x=20, y=30)
    entry_pdf = tk.Entry(root, width=45)
    entry_pdf.place(x=100, y=30)
    btn_pdf = tk.Button(root, text="选择文件", width=10, command=select_pdf)
    btn_pdf.place(x=520, y=27)

    # ========== DOCX选择 ==========
    lab_docx = tk.Label(root, text="💾 保存路径：")
    lab_docx.place(x=20, y=80)
    entry_docx = tk.Entry(root, width=45)
    entry_docx.place(x=100, y=80)
    btn_docx = tk.Button(root, text="选择路径", width=10, command=select_docx)
    btn_docx.place(x=520, y=77)

    # ========== 页码选择 ==========
    lab_page = tk.Label(root, text="📖 页码选择：")
    lab_page.place(x=20, y=130)
    combox_page = ttk.Combobox(root, width=15, state="readonly")
    combox_page["values"] = ("全部页面", "自定义")
    combox_page.current(0)
    combox_page.place(x=100, y=130)
    # 自定义页码输入框
    lab_start = tk.Label(root, text="起始页(0开始)：")
    lab_start.place(x=250, y=132)
    entry_start = tk.Entry(root, width=8)
    entry_start.place(x=350, y=130)
    lab_end = tk.Label(root, text="结束页：")
    lab_end.place(x=420, y=132)
    entry_end = tk.Entry(root, width=8)
    entry_end.place(x=470, y=130)

    # ========== 图片选择 ==========
    lab_img = tk.Label(root, text="🖼️ 图片处理：")
    lab_img.place(x=20, y=180)
    combox_img = ttk.Combobox(root, width=15, state="readonly")
    combox_img["values"] = ("保留图片", "忽略图片(更快)")
    combox_img.current(0)
    combox_img.place(x=100, y=180)

    # ========== 转换按钮 ==========
    btn_convert = tk.Button(root, text="开始转换", width=20, height=2, 
                            bg="#4CAF50", fg="white", command=start_convert)
    btn_convert.place(x=180, y=230)

    # 窗口居中
    root.update_idletasks()
    x = (root.winfo_screenwidth() - root.winfo_width()) // 2
    y = (root.winfo_screenheight() - root.winfo_height()) // 2
    root.geometry(f"+{x}+{y}")

    root.mainloop()

if __name__ == "__main__":
    create_gui()
# partly developed with AI