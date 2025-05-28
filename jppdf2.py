import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import threading
import pypdf

class NDLPDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("jp国会国立图书馆PDF合并工具")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # 设置窗口主题颜色
        self.root.configure(bg="#f3b6e1")
        
        # 创建主框架
        main_frame = tk.Frame(root, bg="#e4f7d7", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 输入路径选择
        input_frame = tk.LabelFrame(main_frame, text="选择PDF文件所在文件夹", bg="#f0f0f0", padx=10, pady=10)
        input_frame.pack(fill=tk.X, pady=10)
        
        self.source_path = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.source_path, width=70).pack(side=tk.LEFT, padx=5)
        tk.Button(input_frame, text="浏览...", command=self.browse_input).pack(side=tk.RIGHT)
        
        # 输出路径选择
        output_frame = tk.LabelFrame(main_frame, text="选择合并后PDF保存位置", bg="#f0f0f0", padx=10, pady=10)
        output_frame.pack(fill=tk.X, pady=10)
        
        self.library_path = tk.StringVar()
        tk.Entry(output_frame, textvariable=self.library_path, width=70).pack(side=tk.LEFT, padx=5)
        tk.Button(output_frame, text="浏览...", command=self.browse_output).pack(side=tk.RIGHT)
        
        # 合并模式选择
        mode_frame = tk.LabelFrame(main_frame, text="合并模式", bg="#f0f0f0", padx=10, pady=10)
        mode_frame.pack(fill=tk.X, pady=10)
        
        self.merge_mode = tk.StringVar(value="group_by_id")
        tk.Radiobutton(mode_frame, text="按NDL ID分组合并", variable=self.merge_mode, 
                      value="group_by_id", bg="#f0f0f0", command=self.toggle_regex_frame).pack(side=tk.LEFT, padx=20)
        tk.Radiobutton(mode_frame, text="直接合并所有文件", variable=self.merge_mode,
                      value="merge_all", bg="#f0f0f0", command=self.toggle_regex_frame).pack(side=tk.LEFT, padx=20)
        
        # NDL ID正则表达式
        self.regex_frame = tk.LabelFrame(main_frame, text="NDL ID正则表达式", bg="#f0f0f0", padx=10, pady=10)
        self.regex_frame.pack(fill=tk.X, pady=10)
        
        self.ndl_id_re = tk.StringVar(value=r'digidepo_(\d+)_')
        tk.Entry(self.regex_frame, textvariable=self.ndl_id_re, width=70).pack(fill=tk.X, padx=5)
        
        # 输出文件名设置（用于直接合并模式）
        self.output_name_frame = tk.LabelFrame(main_frame, text="输出文件名", bg="#f0f0f0", padx=10, pady=10)
        self.output_name = tk.StringVar(value="")
        tk.Entry(self.output_name_frame, textvariable=self.output_name, width=70).pack(fill=tk.X, padx=5)
        
        # 排序选项
        sort_frame = tk.LabelFrame(main_frame, text="PDF文件排序方式", bg="#f0f0f0", padx=10, pady=10)
        sort_frame.pack(fill=tk.X, pady=10)
        
        self.sort_method = tk.StringVar(value="bracket_number")
        tk.Radiobutton(sort_frame, text="按括号中的数字排序", variable=self.sort_method, 
                      value="bracket_number", bg="#f0f0f0").pack(side=tk.LEFT, padx=20)
        tk.Radiobutton(sort_frame, text="别选", variable=self.sort_method,
                      value="name", bg="#f0f0f0").pack(side=tk.LEFT, padx=20)
        tk.Radiobutton(sort_frame, text="别选", variable=self.sort_method,
                      value="number", bg="#f0f0f0").pack(side=tk.LEFT, padx=20)
        
        # 执行按钮
        button_frame = tk.Frame(main_frame, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, pady=10)
        
        self.merge_button = tk.Button(button_frame, text="开始合并", 
                                     command=self.start_merge, width=20, height=2,
                                     bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        self.merge_button.pack()
        
        # 进度显示区域
        progress_frame = tk.LabelFrame(main_frame, text="处理进度", bg="#f0f0f0", padx=10, pady=10)
        progress_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        # 日志文本框
        self.log_text = scrolledtext.ScrolledText(progress_frame, height=15, width=80)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 状态变量
        self.is_merging = False
        
        # 初始化UI状态
        self.toggle_regex_frame()
        
    def toggle_regex_frame(self):
        if self.merge_mode.get() == "group_by_id":
            self.regex_frame.pack(fill=tk.X, pady=10, after=self.regex_frame.master.children['!labelframe3'])
            try:
                self.output_name_frame.pack_forget()
            except:
                pass
        else:  # merge_all mode
            try:
                self.regex_frame.pack_forget()
                self.output_name_frame.pack(fill=tk.X, pady=10, after=self.regex_frame.master.children['!labelframe3'])
            except:
                pass
        
    def browse_input(self):
        folder_path = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
        if folder_path:
            self.source_path.set(folder_path)
            
    def browse_output(self):
        if self.merge_mode.get() == "group_by_id":
            folder_path = filedialog.askdirectory(title="选择合并后PDF文件的保存位置")
            if folder_path:
                self.library_path.set(folder_path)
        else:
            file_path = filedialog.asksaveasfilename(
                title="保存合并后的PDF文件",
                defaultextension=".pdf",
                filetypes=[("PDF文件", "*.pdf")]
            )
            if file_path:
                self.library_path.set(os.path.dirname(file_path))
                self.output_name.set(os.path.basename(file_path))
    
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def start_merge(self):
        # 检查输入和输出路径
        source_path = self.source_path.get()
        library_path = self.library_path.get()
        
        if not source_path or not os.path.isdir(source_path):
            messagebox.showerror("错误", "请选择有效的输入文件夹!")
            return
            
        if not library_path:
            if self.merge_mode.get() == "group_by_id":
                if not os.path.isdir(library_path):
                    messagebox.showerror("错误", "请选择有效的输出文件夹!")
                    return
            else:
                messagebox.showerror("错误", "请指定输出文件路径!")
                return
                
        # 检查文件夹是否为空
        if not os.listdir(source_path):
            messagebox.showerror("错误", "输入文件夹中没有文件!")
            return
            
        # 防止重复执行
        if self.is_merging:
            return
            
        # 禁用按钮
        self.merge_button.config(state=tk.DISABLED)
        self.is_merging = True
        
        # 清空日志
        self.log_text.delete(1.0, tk.END)
        
        # 创建线程执行合并操作
        merge_thread = threading.Thread(
            target=self.process_pdf_files if self.merge_mode.get() == "group_by_id" else self.merge_all_pdfs,
            args=(source_path, library_path)
        )
        merge_thread.daemon = True
        merge_thread.start()
    
    def merge_all_pdfs(self, source_path, library_path):
        """直接合并所有PDF文件，不进行分组，但使用元数据命名"""
        try:
            self.log("开始合并操作...")
            # 获取PDF文件
            pdf_files = self.get_pdf_files(source_path)
            if not pdf_files:
                self.log("文件夹中没有PDF文件!")
                messagebox.showinfo("信息", "文件夹中没有PDF文件!")
                self.merge_button.config(state=tk.NORMAL)
                self.is_merging = False
                return
                
            # 排序文件
            self.log("对文件进行排序...")
            self.sort_pdf_files(pdf_files)
                
            # 设置进度条
            self.progress_bar["maximum"] = len(pdf_files)
            self.progress_bar["value"] = 0
            
            # 尝试从第一个PDF获取元数据
            try:
                pdf_metadata = pypdf.PdfReader(pdf_files[0]).metadata
                pdf_metadata = {k: pdf_metadata[k] for k in pdf_metadata.keys() if k in pdf_metadata}
                
                # 生成文件名
                output_filename = self.output_name.get()  # 默认使用用户输入的名称
                
                # 如果有Keywords，使用元数据生成文件名
                if '/Keywords' in pdf_metadata:
                    self.log(f"从第一个PDF文件获取到Keywords: {pdf_metadata['/Keywords']}")
                    
                    # 分割Keywords
                    split_keywords = self.keywords_splitter(pdf_metadata['/Keywords'])
                    
                    if len(split_keywords) >= 3:
                        keywords_title_author = split_keywords[0]
                        keywords_publisher = split_keywords[1]
                        keywords_year = split_keywords[2]
                        
                        # 处理年份
                        dot_index = keywords_year.find('.')
                        if dot_index > 0:
                            keywords_year = keywords_year[:dot_index]
                        
                        # 清理不允许的字符
                        keywords_title_author = re.sub(r'[\\/:*?"<>|]', '', keywords_title_author)
                        
                        # 创建文件名
                        output_filename = f"{keywords_title_author}_{keywords_publisher}_{keywords_year}.pdf"
                        self.log(f"使用元数据生成的文件名: {output_filename}")
                    else:
                        self.log("无法正确分割Keywords，使用默认文件名")
                else:
                    self.log("PDF没有Keywords元数据，使用默认文件名")
            except Exception as e:
                self.log(f"获取元数据时出错: {str(e)}，使用默认文件名")
            
            # 合并PDF
            self.log("开始合并PDF文件...")
            merger = pypdf.PdfMerger()
            
            # 添加文件并更新进度条
            for i, pdf_file in enumerate(pdf_files):
                try:
                    merger.append(pdf_file)
                    self.log(f"已添加: {os.path.basename(pdf_file)}")
                    # 更新进度条
                    self.progress_bar["value"] = i + 1
                    self.root.update_idletasks()
                except Exception as e:
                    self.log(f"无法添加 {os.path.basename(pdf_file)}: {str(e)}")
            
            # 设置元数据
            try:
                merger.add_metadata(pdf_metadata)
            except:
                self.log("无法添加元数据到合并后的PDF")
            
            # 确保输出目录存在
            os.makedirs(library_path, exist_ok=True)
            
            # 确保文件名以.pdf结尾
            if not output_filename.lower().endswith('.pdf'):
                output_filename += '.pdf'
            
            output_file = os.path.join(library_path, output_filename)
            self.log(f"正在保存合并后的PDF到: {output_file}")
            
            # 写入合并后的PDF
            merger.write(output_file)
            merger.close()
            
            self.log("\n合并操作完成!")
            messagebox.showinfo("完成", "PDF文件合并完成!")
            
            # 询问是否打开文件
            if messagebox.askyesno("打开文件", "是否打开合并后的PDF文件?"):
                self.open_file(output_file)
                
        except Exception as e:
            self.log(f"发生错误: {str(e)}")
            messagebox.showerror("错误", f"合并过程中发生错误:\n{str(e)}")
        finally:
            # 恢复按钮状态
            self.merge_button.config(state=tk.NORMAL)
            self.is_merging = False
    
    def process_pdf_files(self, source_path, library_path):
        try:
            self.log("开始合并操作...")
            # 获取PDF文件
            pdf_files = self.get_pdf_files(source_path)
            if not pdf_files:
                self.log("文件夹中没有PDF文件!")
                messagebox.showinfo("信息", "文件夹中没有PDF文件!")
                self.merge_button.config(state=tk.NORMAL)
                self.is_merging = False
                return
            
            # 按NDL ID分组文件
            pdf_files_grouped = self.group_pdf_files(pdf_files)
            if not pdf_files_grouped:
                self.log("未找到包含NDL ID的PDF文件!")
                
                # 询问是否要直接合并所有文件
                if messagebox.askyesno("未找到NDL ID", "未找到包含NDL ID的PDF文件，是否直接合并所有文件?"):
                    self.merge_all_pdfs(source_path, library_path)
                    return
                else:
                    self.merge_button.config(state=tk.NORMAL)
                    self.is_merging = False
                    return
                
            # 设置进度条最大值
            self.progress_bar["maximum"] = len(pdf_files_grouped)
            self.progress_bar["value"] = 0
            
            # 合并PDF文件
            self.merge_pdf_files(pdf_files_grouped, library_path)
            
            # 完成进度条
            self.progress_bar["value"] = len(pdf_files_grouped)
            
            self.log("\n合并操作完成!")
            messagebox.showinfo("完成", "PDF文件合并完成!")
                
        except Exception as e:
            self.log(f"发生错误: {str(e)}")
            messagebox.showerror("错误", f"合并过程中发生错误:\n{str(e)}")
        finally:
            # 恢复按钮状态
            self.merge_button.config(state=tk.NORMAL)
            self.is_merging = False
            
    def get_pdf_files(self, source_path):
        pdf_files = []
        for filename in os.listdir(source_path):
            if filename.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(source_path, filename))
        self.log(f"找到PDF文件数量: {len(pdf_files)}")
        return pdf_files
        
    def group_pdf_files(self, pdf_files):
        ndl_id_pattern = self.ndl_id_re.get()
        pdf_files_grouped = {}
        for pdf_file in pdf_files:
            ndl_id_match = re.search(ndl_id_pattern, pdf_file)
            if ndl_id_match:
                extracted_ndl_id = ndl_id_match.group(1)
                if extracted_ndl_id in pdf_files_grouped:
                    pdf_files_grouped[extracted_ndl_id].append(pdf_file)
                else:
                    pdf_files_grouped[extracted_ndl_id] = [pdf_file]
        self.log(f"PDF分组数量: {len(pdf_files_grouped)}")
        return pdf_files_grouped

    def sort_pdf_files(self, pdf_files):
        """根据选择的方法对PDF文件列表进行排序"""
        if self.sort_method.get() == "bracket_number":
            # 按括号中的数字排序
            def sort_key(pdf_file):
                match = re.search(r'\((\d+)\)', pdf_file)
                if match:
                    return int(match.group(1))
                else:
                    return -1
            pdf_files.sort(key=sort_key)
            self.log("使用括号内数字排序...")
        elif self.sort_method.get() == "name":
            # 按文件名排序
            pdf_files.sort()
            self.log("使用文件名排序...")
        else:  # 数字排序
            # 从文件名中提取数字进行排序
            def extract_number(filename):
                numbers = ''.join(c for c in os.path.basename(filename) if c.isdigit())
                return int(numbers) if numbers else float('inf')
            
            pdf_files.sort(key=extract_number)
            self.log("使用文件名中的数字排序...")
        
    def merge_pdf_files(self, pdf_files_grouped, library_path):
        group_count = 0
        for ndl_id, pdf_files in pdf_files_grouped.items():
            group_count += 1
            self.log(f"\n处理PDF组 {group_count}: {ndl_id}")
            
            # 更新进度条
            self.progress_bar["value"] = group_count - 1
            self.root.update_idletasks()
            
            # 对当前组的文件进行排序
            self.sort_pdf_files(pdf_files)
            
            try:
                # 获取第一个PDF的元数据
                pdf_metadata = pypdf.PdfReader(pdf_files[0]).metadata
                pdf_metadata = {k: pdf_metadata[k] for k in pdf_metadata.keys()}
                
                # 处理Keywords
                if '/Keywords' in pdf_metadata:
                    self.log(f"Keywords: {pdf_metadata['/Keywords']}")
                    
                    # 分割Keywords
                    split_keywords = self.keywords_splitter(pdf_metadata['/Keywords'])
                    
                    if len(split_keywords) >= 3:
                        keywords_title_author = split_keywords[0]
                        keywords_publisher = split_keywords[1]
                        keywords_year = split_keywords[2]
                        
                        # 处理年份
                        dot_index = keywords_year.find('.')
                        if dot_index > 0:
                            keywords_year = keywords_year[:dot_index]
                        
                        # 创建文件夹名称
                        folder_name = f"{keywords_publisher}_{keywords_year}_{ndl_id}_{keywords_title_author}"
                        
                        # 清理不允许的字符
                        keywords_title_author = re.sub(r'[\\/:*?"<>|]', '', keywords_title_author)
                        folder_name = re.sub(r'[\\/:*?"<>|]', '', folder_name)
                        
                        # 创建文件名
                        merged_file_name = f"{keywords_title_author}_{ndl_id}.pdf"
                    else:
                        # 如果无法正确分割Keywords，使用默认命名
                        self.log("无法正确分割Keywords，使用默认命名")
                        folder_name = f"group_{ndl_id}"
                        merged_file_name = f"merged_{ndl_id}.pdf"
                else:
                    # 如果没有Keywords，使用默认命名
                    self.log("PDF没有Keywords元数据，使用默认命名")
                    folder_name = f"group_{ndl_id}"
                    merged_file_name = f"merged_{ndl_id}.pdf"
                
                self.log(f"文件夹名称: {folder_name}")
                self.log(f"合并后文件名: {merged_file_name}")
                
                # 合并PDF
                merger = pypdf.PdfMerger()
                for pdf_file in pdf_files:
                    try:
                        merger.append(pdf_file)
                        self.log(f"已添加: {os.path.basename(pdf_file)}")
                    except Exception as e:
                        self.log(f"无法添加 {os.path.basename(pdf_file)}: {str(e)}")
                
                # 添加元数据
                merger.add_metadata(pdf_metadata)
                
                # 创建输出路径
                if '/Keywords' in pdf_metadata and len(split_keywords) >= 3:
                    library_output_path = os.path.join(library_path, keywords_publisher, folder_name)
                else:
                    library_output_path = os.path.join(library_path, folder_name)
                
                self.log(f"输出路径: {library_output_path}")
                
                # 创建目录
                os.makedirs(library_output_path, exist_ok=True)
                
                # 写入合并后的PDF
                output_file = os.path.join(library_output_path, merged_file_name)
                merger.write(output_file)
                merger.close()
                
                self.log(f"已处理PDF组 {group_count}")
            
            except Exception as e:
                self.log(f"处理组 {ndl_id} 时出错: {str(e)}")
    
    def keywords_splitter(self, keywords):
        """按照原始代码中的逻辑分割Keywords"""
        keywords_slicer = [m.start() for m in re.finditer(r'\S,\S', keywords)]
        parts = []
        start = 0
        for slicer in keywords_slicer:
            part = keywords[start:slicer + 1]
            parts.append(part)
            start = slicer + 2
        parts.append(keywords[start:])
        return parts
        
    def open_file(self, file_path):
        """打开指定的文件"""
        try:
            if sys.platform == 'win32':
                os.startfile(file_path)
            elif sys.platform == 'darwin':  # macOS
                os.system(f'open "{file_path}"')
            else:  # Linux
                os.system(f'xdg-open "{file_path}"')
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = NDLPDFMergerApp(root)
    root.mainloop()