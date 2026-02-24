from pdf2docx import Converter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

# 全局变量存储路径
pdf_path_global = ""
docx_path_global = ""

def select_pdf():
    """选择PDF文件，弹窗选择"""
    global pdf_path_global
    path = filedialog.askopenfilename(
        title="选择要转换的PDF文件",
        filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")],
        initialdir="./"  # 初始打开当前目录
    )
    if path:
        pdf_path_global = path
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, path)
        # 自动生成docx保存路径（同目录同名）
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
    """GUI版单个PDF转换，带参数传入，返回结果信息"""
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
    parse_images = False if img_choice == "忽略" else True
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
    """开始转换按钮触发函数，开启子线程避免界面卡死"""
    # 校验路径
    if not pdf_path_global or not docx_path_global:
        messagebox.showerror("错误", "请先选择PDF文件和保存路径！")
        return
    # 获取选择的参数
    page_choice = combox_page.get()
    img_choice = combox_img.get()
    # 按钮置灰，防止重复点击
    btn_convert.config(state=tk.DISABLED, text="转换中...")
    # 开启子线程执行转换（GUI不卡死）
    t = threading.Thread(target=convert_thread, args=(pdf_path_global, docx_path_global, page_choice, img_choice))
    t.daemon = True
    t.start()

def convert_thread(pdf_path, docx_path, page_choice, img_choice):
    """转换子线程，执行后更新界面"""
    success, msg = convert_single_pdf_gui(pdf_path, docx_path, page_choice, img_choice)
    # 弹窗提示结果
    if success:
        messagebox.showinfo("成功", msg)
    else:
        messagebox.showerror("失败", msg)
    # 恢复按钮状态
    btn_convert.config(state=tk.NORMAL, text="开始转换")
    # 清空输入框，准备下一次转换（保留GUI，支持循环转换）
    entry_pdf.delete(0, tk.END)
    entry_docx.delete(0, tk.END)
    global pdf_path_global, docx_path_global
    pdf_path_global = ""
    docx_path_global = ""

def create_gui():
    """创建GUI界面"""
    root = tk.Tk()
    root.title("PDF转DOCX工具 | 一键转换")
    root.geometry("650x320")  # 窗口大小
    root.resizable(False, False)  # 禁止缩放窗口
    root.iconbitmap()  # 可自定义图标，注释则用默认

    # 全局化控件，方便子线程调用
    global entry_pdf, entry_docx, combox_page, combox_img, entry_start, entry_end, btn_convert

    # ========== 第一行：PDF文件选择 ==========
    lab_pdf = tk.Label(root, text="📄 PDF源文件：", font=("微软雅黑", 10))
    lab_pdf.place(x=20, y=30)
    entry_pdf = tk.Entry(root, width=45, font=("微软雅黑", 10))
    entry_pdf.place(x=100, y=30)
    btn_pdf = tk.Button(root, text="选择文件", font=("微软雅黑", 9), width=10, command=select_pdf)
    btn_pdf.place(x=520, y=27)

    # ========== 第二行：DOCX保存选择 ==========
    lab_docx = tk.Label(root, text="💾 保存路径：", font=("微软雅黑", 10))
    lab_docx.place(x=20, y=80)
    entry_docx = tk.Entry(root, width=45, font=("微软雅黑", 10))
    entry_docx.place(x=100, y=80)
    btn_docx = tk.Button(root, text="选择路径", font=("微软雅黑", 9), width=10, command=select_docx)
    btn_docx.place(x=520, y=77)

    # ========== 第三行：页码选择 ==========
    lab_page = tk.Label(root, text="📖 页码选择：", font=("微软雅黑", 10))
    lab_page.place(x=20, y=130)
    combox_page = ttk.Combobox(root, width=15, font=("微软雅黑", 10), state="readonly")
    combox_page["values"] = ("全部页面", "自定义")
    combox_page.current(0)  # 默认选全部页面
    combox_page.place(x=100, y=130)
    # 自定义页码输入框
    lab_start = tk.Label(root, text="起始页(0开始)：", font=("微软雅黑", 9))
    lab_start.place(x=250, y=132)
    entry_start = tk.Entry(root, width=8, font=("微软雅黑", 10))
    entry_start.place(x=350, y=130)
    lab_end = tk.Label(root, text="结束页(留空为最后)：", font=("微软雅黑", 9))
    lab_end.place(x=420, y=132)
    entry_end = tk.Entry(root, width=8, font=("微软雅黑", 10))
    entry_end.place(x=530, y=130)

    # ========== 第四行：图片选择 ==========
    lab_img = tk.Label(root, text="🖼️ 图片处理：", font=("微软雅黑", 10))
    lab_img.place(x=20, y=180)
    combox_img = ttk.Combobox(root, width=15, font=("微软雅黑", 10), state="readonly")
    combox_img["values"] = ("保留图片", "忽略图片(更快)")
    combox_img.current(0)  # 默认保留图片
    combox_img.place(x=100, y=180)

    # ========== 第五行：转换按钮 ==========
    btn_convert = tk.Button(root, text="开始转换", font=("微软雅黑", 12, "bold"), 
                            width=20, height=2, bg="#4CAF50", fg="white", command=start_convert)
    btn_convert.place(x=180, y=230)

    # 窗口居中显示（可选，优化体验）
    root.update_idletasks()
    x = (root.winfo_screenwidth() - root.winfo_width()) // 2
    y = (root.winfo_screenheight() - root.winfo_height()) // 2
    root.geometry(f"+{x}+{y}")

    root.mainloop()

if __name__ == "__main__":
    # 解决tkinter中文乱码（Windows）
    tk.rcParams = {'font.sans-serif': ['Microsoft YaHei'], 'axes.unicode_minus': False}
    create_gui()