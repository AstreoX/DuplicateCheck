import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from tqdm import tqdm
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import warnings
import threading
from tkinter import ttk
import shutil

# 忽略警告
warnings.filterwarnings('ignore')

class LoadingWindow:
    def __init__(self, parent, title="处理中"):
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("500x200")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()
        
        # 设置窗口在父窗口中央
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        window_width = 500
        window_height = 200
        
        x = parent_x + (parent_width - window_width) // 2
        y = parent_y + (parent_height - window_height) // 2
        
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 添加主提示文本
        self.label_var = tk.StringVar(value="正在处理，请稍候...")
        label = tk.Label(self.window, textvariable=self.label_var, font=("Arial", 12, "bold"))
        label.pack(pady=(15, 5))
        
        # 添加详细状态指示
        self.status_frame = tk.Frame(self.window)
        self.status_frame.pack(fill=tk.X, padx=20, pady=5)
        
        # 网络状态
        self.network_frame = tk.Frame(self.status_frame)
        self.network_frame.pack(fill=tk.X, pady=2)
        
        self.network_label = tk.Label(self.network_frame, text="网络状态:", width=15, anchor="w")
        self.network_label.pack(side=tk.LEFT)
        
        self.network_status_var = tk.StringVar(value="--")
        self.network_status = tk.Label(self.network_frame, textvariable=self.network_status_var, anchor="w", fg="blue")
        self.network_status.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 内存状态
        self.memory_frame = tk.Frame(self.status_frame)
        self.memory_frame.pack(fill=tk.X, pady=2)
        
        self.memory_label = tk.Label(self.memory_frame, text="内存状态:", width=15, anchor="w")
        self.memory_label.pack(side=tk.LEFT)
        
        self.memory_status_var = tk.StringVar(value="--")
        self.memory_status = tk.Label(self.memory_frame, textvariable=self.memory_status_var, anchor="w", fg="blue")
        self.memory_status.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 处理阶段
        self.phase_frame = tk.Frame(self.status_frame)
        self.phase_frame.pack(fill=tk.X, pady=2)
        
        self.phase_label = tk.Label(self.phase_frame, text="处理阶段:", width=15, anchor="w")
        self.phase_label.pack(side=tk.LEFT)
        
        self.phase_status_var = tk.StringVar(value="--")
        self.phase_status = tk.Label(self.phase_frame, textvariable=self.phase_status_var, anchor="w", fg="blue")
        self.phase_status.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 添加进度条框架
        self.progress_frame = tk.Frame(self.window)
        self.progress_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # 进度条比例标签
        self.progress_label_var = tk.StringVar(value="0%")
        self.progress_label = tk.Label(self.progress_frame, textvariable=self.progress_label_var, width=6)
        self.progress_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # 添加进度条
        self.progress = ttk.Progressbar(self.progress_frame, mode="determinate", length=400)
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    def update_text(self, text):
        """更新主状态文本"""
        self.label_var.set(text)
        self.window.update_idletasks()
    
    def update_network_status(self, status):
        """更新网络状态"""
        self.network_status_var.set(status)
        self.window.update_idletasks()
    
    def update_memory_status(self, status):
        """更新内存状态"""
        self.memory_status_var.set(status)
        self.window.update_idletasks()
    
    def update_phase(self, phase):
        """更新处理阶段"""
        self.phase_status_var.set(phase)
        self.window.update_idletasks()
    
    def update_progress(self, percent):
        """更新进度条"""
        self.progress["value"] = percent
        self.progress_label_var.set(f"{percent}%")
        self.window.update_idletasks()
    
    def close(self):
        self.progress.stop() if hasattr(self.progress, 'stop') and self.progress['mode'] == 'indeterminate' else None
        self.window.destroy()

class DuplicateQuestionChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("题库查重工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        self.model = None
        self.df = None
        self.embeddings = None
        self.similarity_matrix = None
        self.duplicate_pairs = []
        self.loading_window = None
        
        # 设置默认模型缓存路径
        self.default_model_dir = os.path.join(os.path.expanduser("~"), ".cache", "sentence-transformers")
        self.model_cache_dir = self.default_model_dir
        
        self.setup_ui()
    
    def setup_ui(self):
        # 创建主框架
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 文件选择区域
        file_frame = tk.LabelFrame(main_frame, text="文件选择")
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.file_path_var = tk.StringVar()
        self.file_path_var.set("d:\\Projects\\AICCwithEM\\面向对象程序设计-题库 (1).xls")
        
        file_entry = tk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        file_entry.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        
        browse_button = tk.Button(file_frame, text="浏览", command=self.browse_file)
        browse_button.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 模型缓存路径区域
        model_cache_frame = tk.LabelFrame(main_frame, text="模型缓存设置")
        model_cache_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 模型缓存路径输入
        cache_path_label = tk.Label(model_cache_frame, text="模型缓存路径:")
        cache_path_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.cache_path_var = tk.StringVar(value=self.model_cache_dir)
        cache_path_entry = tk.Entry(model_cache_frame, textvariable=self.cache_path_var, width=50)
        cache_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        
        browse_cache_button = tk.Button(model_cache_frame, text="浏览", command=self.browse_cache_dir)
        browse_cache_button.grid(row=0, column=2, padx=5, pady=5)
        
        # 缓存状态显示
        self.cache_status_var = tk.StringVar(value="")
        cache_status_label = tk.Label(model_cache_frame, textvariable=self.cache_status_var, fg="blue")
        cache_status_label.grid(row=1, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)
        
        # 查询缓存状态按钮
        check_cache_button = tk.Button(model_cache_frame, text="检查缓存状态", command=self.check_model_cache)
        check_cache_button.grid(row=2, column=0, columnspan=3, padx=5, pady=5)
        
        # 参数设置区域
        param_frame = tk.LabelFrame(main_frame, text="参数设置")
        param_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 相似度阈值
        threshold_label = tk.Label(param_frame, text="相似度阈值:")
        threshold_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.threshold_var = tk.DoubleVar(value=0.85)
        threshold_scale = tk.Scale(param_frame, variable=self.threshold_var, 
                                  from_=0.5, to=1.0, resolution=0.01, orient=tk.HORIZONTAL, length=200)
        threshold_scale.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 模型选择
        model_label = tk.Label(param_frame, text="嵌入模型:")
        model_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.model_var = tk.StringVar(value="paraphrase-multilingual-MiniLM-L12-v2")
        models = ["paraphrase-multilingual-MiniLM-L12-v2", "distiluse-base-multilingual-cased-v1"]
        model_menu = tk.OptionMenu(param_frame, self.model_var, *models)
        model_menu.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 按钮区域
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        analyze_button = tk.Button(button_frame, text="开始分析", command=self.start_analysis_thread, bg="#4CAF50", fg="white", height=2)
        analyze_button.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        
        # 结果显示区域
        result_frame = tk.LabelFrame(main_frame, text="查重结果")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建文本框和滚动条
        self.result_text = tk.Text(result_frame, wrap=tk.WORD, height=15)
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = tk.Scrollbar(result_frame, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 进度条框架
        self.progress_frame = tk.Frame(main_frame)
        self.progress_frame.pack(fill=tk.X, padx=5, pady=5)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = tk.Scale(self.progress_frame, variable=self.progress_var, 
                                   from_=0, to=100, orient=tk.HORIZONTAL, 
                                   showvalue=True, state=tk.DISABLED)
        self.progress_bar.pack(fill=tk.X)

        # 启动时检查缓存状态
        self.root.after(500, self.check_model_cache)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            initialdir=os.path.dirname(self.file_path_var.get()),
            title="选择Excel题库文件",
            filetypes=(("Excel文件", "*.xls *.xlsx"), ("所有文件", "*.*"))
        )
        if file_path:
            self.file_path_var.set(file_path)
    
    def browse_cache_dir(self):
        """选择模型缓存目录"""
        dir_path = filedialog.askdirectory(
            initialdir=self.cache_path_var.get(),
            title="选择模型缓存目录"
        )
        if dir_path:
            self.cache_path_var.set(dir_path)
            self.model_cache_dir = dir_path
            self.check_model_cache()
    
    def model_exists_in_cache(self, model_name, cache_dir):
        """检查模型是否存在于缓存目录中"""
        # 正常缓存目录结构
        model_dir_name = f"models--sentence-transformers--{model_name}"
        model_dir = os.path.join(cache_dir, model_dir_name)
        
        # 直接检查目录存在性，不再检查config.json
        if os.path.exists(model_dir) and os.path.isdir(model_dir):
            # 检查是否有任何文件，确保不是空目录
            if len(os.listdir(model_dir)) > 0:
                return True, model_dir
        
        # 尝试在任何子目录中查找模型目录
        for root, dirs, _ in os.walk(cache_dir):
            for dir_name in dirs:
                if dir_name == model_dir_name:
                    model_path = os.path.join(root, dir_name)
                    if os.path.exists(model_path) and len(os.listdir(model_path)) > 0:
                        return True, model_path
        
        return False, None
    
    def check_model_cache(self):
        """检查指定目录中是否有缓存的模型"""
        cache_dir = self.cache_path_var.get()
        model_name = self.model_var.get()
        
        # 确保目录存在
        if not os.path.exists(cache_dir):
            try:
                os.makedirs(cache_dir, exist_ok=True)
                self.cache_status_var.set(f"已创建新的缓存目录: {cache_dir}")
                return False
            except Exception as e:
                self.cache_status_var.set(f"创建缓存目录失败: {str(e)}")
                return False
        
        # 检查模型是否已缓存
        is_cached, model_cache_path = self.model_exists_in_cache(model_name, cache_dir)
        if is_cached:
            self.cache_status_var.set(f"模型 '{model_name}' 已在本地缓存\n路径: {model_cache_path}")
            return True
        
        self.cache_status_var.set(f"模型 '{model_name}' 未在本地缓存，将需要下载")
        return False
    
    def update_status(self, text):
        """更新状态栏文本，确保在主线程中执行"""
        self.root.after(0, lambda: self.status_var.set(text))
    
    def update_loading_text(self, text):
        """更新加载窗口文本，确保在主线程中执行"""
        if self.loading_window:
            self.root.after(0, lambda: self.loading_window.update_text(text))
    
    def load_data(self):
        try:
            # 尝试读取Excel文件
            file_path = self.file_path_var.get()
            self.update_status(f"正在读取文件: {os.path.basename(file_path)}")
            self.update_loading_text(f"正在读取文件: {os.path.basename(file_path)}")
            
            # 更新加载状态
            self.root.after(0, lambda: self.loading_window.update_phase("读取数据文件"))
            self.root.after(0, lambda: self.loading_window.update_network_status("不需要网络连接"))
            self.root.after(0, lambda: self.loading_window.update_memory_status("正在读取文件..."))
            self.root.after(0, lambda: self.loading_window.update_progress(10))
            
            # 尝试不同的编码方式读取Excel
            try:
                self.df = pd.read_excel(file_path)
            except Exception as e:
                # 如果默认读取失败，尝试指定引擎
                self.df = pd.read_excel(file_path, engine='xlrd')
            
            self.root.after(0, lambda: self.loading_window.update_progress(40))
            
            # 检查数据是否成功加载
            if self.df is None or self.df.empty:
                self.root.after(0, lambda: messagebox.showerror("错误", "无法读取文件或文件为空"))
                return False
            
            # 检查数据结构
            self.update_status("分析数据结构...")
            self.update_loading_text("分析数据结构...")
            self.root.after(0, lambda: self.loading_window.update_memory_status("分析数据结构..."))
            self.root.after(0, lambda: self.loading_window.update_progress(60))
            
            # 假设题目内容在第三列（索引为2）
            # 如果列名不存在或数据结构不符合预期，尝试基于位置获取
            if len(self.df.columns) >= 3:
                self.question_col = 2  # 假设题目在第三列
            else:
                self.root.after(0, lambda: messagebox.showerror("错误", "文件结构不符合预期，无法识别题目列"))
                return False
            
            # 移除空行
            self.df = self.df.dropna(subset=[self.df.columns[self.question_col]])
            
            # 获取题目ID和内容
            self.question_ids = self.df.iloc[:, 0].astype(str).tolist()  # 第一列作为ID
            self.questions = self.df.iloc[:, self.question_col].astype(str).tolist()  # 题目内容
            
            # 保存原始行号作为题目编号（Excel表格中的行号）
            # 注意：DataFrame的index是从0开始的，而Excel表格的行号是从1开始的，所以需要加1
            self.original_row_numbers = [self.df.index[i] + 1 for i in range(len(self.questions))]
            self.question_numbers = [f"行{num}" for num in self.original_row_numbers]
            
            data_info = f"成功加载 {len(self.questions)} 道题目"
            self.update_status(data_info)
            self.update_loading_text(data_info)
            self.root.after(0, lambda: self.loading_window.update_memory_status(data_info))
            self.root.after(0, lambda: self.loading_window.update_progress(100))
            return True
        
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"加载数据时出错: {str(e)}"))
            return False
    
    def load_model(self):
        try:
            model_name = self.model_var.get()
            cache_dir = self.cache_path_var.get()
            
            self.update_status(f"加载模型: {model_name}")
            self.update_loading_text(f"加载模型: {model_name}")
            
            # 更新加载状态信息
            self.root.after(0, lambda: self.loading_window.update_phase("加载模型"))
            
            # 检查模型是否已在本地缓存
            model_cached, model_cache_path = self.model_exists_in_cache(model_name, cache_dir)
            
            if model_cached:
                self.root.after(0, lambda: self.loading_window.update_network_status("使用本地缓存模型，无需下载"))
                self.root.after(0, lambda: self.loading_window.update_memory_status(f"从缓存读取模型...\n{model_cache_path}"))
                
                # 从缓存加载模型
                os.environ['SENTENCE_TRANSFORMERS_HOME'] = cache_dir
                self.model = SentenceTransformer(model_name)
                
                self.root.after(0, lambda: self.loading_window.update_memory_status("已从本地缓存成功加载模型"))
            else:
                self.root.after(0, lambda: self.loading_window.update_network_status("正在从服务器下载模型..."))
                
                # 设置缓存目录并下载/加载模型
                os.environ['SENTENCE_TRANSFORMERS_HOME'] = cache_dir
                self.model = SentenceTransformer(model_name)
                
                self.root.after(0, lambda: self.loading_window.update_network_status("模型下载完成"))
                self.root.after(0, lambda: self.loading_window.update_memory_status("模型已加载到内存"))
                
                # 检查模型是否已保存到缓存目录
                is_cached_now, model_path_now = self.model_exists_in_cache(model_name, cache_dir)
                if is_cached_now:
                    self.root.after(0, lambda: self.loading_window.update_memory_status(f"模型已缓存至:\n{model_path_now}"))
            
            self.update_loading_text(f"模型 {model_name} 加载完成")
            return True
        
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"加载模型时出错: {str(e)}"))
            return False
    
    def generate_embeddings(self):
        try:
            if not self.questions:
                return False
            
            self.update_status("生成文本嵌入向量...")
            self.update_loading_text("生成文本嵌入向量...")
            self.root.after(0, lambda: self.loading_window.update_phase("生成文本嵌入向量"))
            self.root.after(0, lambda: self.loading_window.update_network_status("不需要网络连接"))
            self.root.after(0, lambda: self.loading_window.update_memory_status("正在处理数据..."))
            
            # 批量处理以提高效率
            batch_size = 32
            embeddings = []
            total = len(self.questions)
            
            for i in range(0, len(self.questions), batch_size):
                batch = self.questions[i:i+batch_size]
                
                # 更新内存状态
                batch_info = f"处理批次 {i//batch_size + 1}/{(total-1)//batch_size + 1}"
                self.root.after(0, lambda info=batch_info: self.loading_window.update_memory_status(info))
                
                batch_embeddings = self.model.encode(batch)
                embeddings.extend(batch_embeddings)
                
                # 更新进度条
                progress = min(100, int((i + len(batch)) / total * 100))
                self.root.after(0, lambda p=progress: self.progress_var.set(p))
                self.root.after(0, lambda p=progress: self.loading_window.update_progress(p))
                self.update_loading_text(f"生成嵌入向量... {progress}%")
            
            self.embeddings = np.array(embeddings)
            self.root.after(0, lambda: self.loading_window.update_memory_status("嵌入向量生成完成"))
            self.update_status("嵌入向量生成完成")
            self.update_loading_text("嵌入向量生成完成")
            return True
        
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"生成嵌入向量时出错: {str(e)}"))
            return False
    
    def find_duplicates(self):
        try:
            if self.embeddings is None:
                return False
            
            self.update_status("计算相似度矩阵...")
            self.update_loading_text("计算相似度矩阵...")
            self.root.after(0, lambda: self.loading_window.update_phase("计算相似度和查找重复"))
            self.root.after(0, lambda: self.loading_window.update_network_status("不需要网络连接"))
            self.root.after(0, lambda: self.loading_window.update_memory_status("计算相似度矩阵..."))
            
            # 计算余弦相似度矩阵
            self.similarity_matrix = cosine_similarity(self.embeddings)
            
            # 设置对角线为0，避免自身匹配
            np.fill_diagonal(self.similarity_matrix, 0)
            
            # 查找相似度高于阈值的题目对
            threshold = self.threshold_var.get()
            self.duplicate_pairs = []
            
            total = len(self.questions)
            self.update_loading_text("查找重复题目...")
            self.root.after(0, lambda: self.loading_window.update_memory_status("查找重复题目..."))
            
            for i in range(total):
                # 只比较不同题目之间的相似度，避免自身比较
                for j in range(i+1, total):  # 从i+1开始，确保i≠j
                    if self.similarity_matrix[i, j] >= threshold:
                        self.duplicate_pairs.append((i, j, self.similarity_matrix[i, j]))
                
                # 更新进度条
                progress = min(100, int((i + 1) / total * 100))
                self.root.after(0, lambda p=progress: self.progress_var.set(p))
                self.root.after(0, lambda p=progress: self.loading_window.update_progress(p))
                
                if i % 10 == 0:  # 每10次更新一次提示，避免频繁更新UI
                    memory_info = f"已处理 {i+1}/{total} 题目，找到 {len(self.duplicate_pairs)} 对重复"
                    self.root.after(0, lambda info=memory_info: self.loading_window.update_memory_status(info))
                    self.update_loading_text(f"查找重复题目... {progress}%")
            
            # 按相似度降序排序
            self.duplicate_pairs.sort(key=lambda x: x[2], reverse=True)
            
            result_info = f"找到 {len(self.duplicate_pairs)} 对重复题目"
            self.update_status(result_info)
            self.update_loading_text(result_info)
            self.root.after(0, lambda: self.loading_window.update_memory_status(result_info))
            return True
        
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"查找重复题目时出错: {str(e)}"))
            return False
    
    def display_results(self):
        try:
            self.root.after(0, lambda: self.result_text.delete(1.0, tk.END))
            
            if not self.duplicate_pairs:
                self.root.after(0, lambda: self.result_text.insert(tk.END, "未找到重复题目\n"))
                return
            
            result_text = f"找到 {len(self.duplicate_pairs)} 对重复题目:\n\n"
            
            for i, (idx1, idx2, similarity) in enumerate(self.duplicate_pairs):
                q_id1 = self.question_ids[idx1]
                q_id2 = self.question_ids[idx2]
                q_num1 = self.question_numbers[idx1]
                q_num2 = self.question_numbers[idx2]
                q1 = self.questions[idx1]
                q2 = self.questions[idx2]
                
                result_text += f"重复对 #{i+1} (相似度: {similarity:.4f}):\n"
                result_text += f"题目 {q_num1} (ID: {q_id1}): {q1[:100]}...\n"
                result_text += f"题目 {q_num2} (ID: {q_id2}): {q2[:100]}...\n\n"
            
            final_text = result_text
            self.root.after(0, lambda: self.result_text.insert(tk.END, final_text))
            
            # 生成相似度分布图
            self.root.after(0, self.plot_similarity_distribution)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"显示结果时出错: {str(e)}"))
    
    def plot_similarity_distribution(self):
        try:
            # 创建新窗口显示图表
            plot_window = tk.Toplevel(self.root)
            plot_window.title("Similarity Distribution")
            plot_window.geometry("600x400")
            
            # 提取所有相似度值
            similarities = [sim for _, _, sim in self.duplicate_pairs]
            
            # 创建图表
            fig, ax = plt.subplots(figsize=(8, 5))
            ax.hist(similarities, bins=20, alpha=0.7, color='blue')
            ax.set_xlabel('Similarity')
            ax.set_ylabel('Frequency')
            ax.set_title('Duplicate Questions Similarity Distribution')
            ax.grid(True, linestyle='--', alpha=0.7)
            
            # 在Tkinter窗口中显示图表
            canvas = FigureCanvasTkAgg(fig, master=plot_window)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error generating chart: {str(e)}")
    
    def export_results(self):
        # 可以添加导出结果到CSV或Excel的功能
        pass
    
    def analysis_process(self):
        """在单独线程中运行的分析流程"""
        try:
            # 清空之前的结果
            self.root.after(0, lambda: self.result_text.delete(1.0, tk.END))
            self.duplicate_pairs = []
            
            # 执行分析流程
            if not self.load_data():
                return
            
            if not self.load_model():
                return
            
            if not self.generate_embeddings():
                return
            
            if not self.find_duplicates():
                return
            
            # 显示结果
            self.display_results()
            
            # 分析完成
            self.update_status("分析完成")
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"分析过程出错: {str(e)}"))
        finally:
            # 关闭加载窗口
            if self.loading_window:
                self.root.after(0, self.loading_window.close)
                self.loading_window = None
            
            # 禁用进度条
            self.root.after(0, lambda: self.progress_bar.config(state=tk.DISABLED))
    
    def start_analysis_thread(self):
        """启动分析线程并显示加载窗口"""
        # 设置进度条
        self.progress_var.set(0)
        self.progress_bar.config(state=tk.NORMAL)
        
        # 创建加载窗口
        self.loading_window = LoadingWindow(self.root, "处理中")
        self.loading_window.update_progress(0)
        
        # 创建并启动工作线程
        analysis_thread = threading.Thread(target=self.analysis_process)
        analysis_thread.daemon = True  # 设置为守护线程，随主线程退出而退出
        analysis_thread.start()

# 主程序入口
if __name__ == "__main__":
    root = tk.Tk()
    app = DuplicateQuestionChecker(root)
    root.mainloop()