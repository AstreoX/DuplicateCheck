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

# 忽略警告
warnings.filterwarnings('ignore')

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
        
        analyze_button = tk.Button(button_frame, text="开始分析", command=self.start_analysis, bg="#4CAF50", fg="white", height=2)
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
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            initialdir=os.path.dirname(self.file_path_var.get()),
            title="选择Excel题库文件",
            filetypes=(("Excel文件", "*.xls *.xlsx"), ("所有文件", "*.*"))
        )
        if file_path:
            self.file_path_var.set(file_path)
    
    def load_data(self):
        try:
            # 尝试读取Excel文件
            file_path = self.file_path_var.get()
            self.status_var.set(f"正在读取文件: {os.path.basename(file_path)}")
            self.root.update()
            
            # 尝试不同的编码方式读取Excel
            try:
                self.df = pd.read_excel(file_path)
            except Exception as e:
                # 如果默认读取失败，尝试指定引擎
                self.df = pd.read_excel(file_path, engine='xlrd')
            
            # 检查数据是否成功加载
            if self.df is None or self.df.empty:
                messagebox.showerror("错误", "无法读取文件或文件为空")
                return False
            
            # 检查数据结构
            self.status_var.set("分析数据结构...")
            self.root.update()
            
            # 假设题目内容在第三列（索引为2）
            # 如果列名不存在或数据结构不符合预期，尝试基于位置获取
            if len(self.df.columns) >= 3:
                self.question_col = 2  # 假设题目在第三列
            else:
                messagebox.showerror("错误", "文件结构不符合预期，无法识别题目列")
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
            
            self.status_var.set(f"成功加载 {len(self.questions)} 道题目")
            return True
        
        except Exception as e:
            messagebox.showerror("错误", f"加载数据时出错: {str(e)}")
            return False
    
    def load_model(self):
        try:
            model_name = self.model_var.get()
            self.status_var.set(f"加载模型: {model_name}")
            self.root.update()
            
            # 加载预训练模型
            self.model = SentenceTransformer(model_name)
            return True
        
        except Exception as e:
            messagebox.showerror("错误", f"加载模型时出错: {str(e)}")
            return False
    
    def generate_embeddings(self):
        try:
            if not self.questions:
                return False
            
            self.status_var.set("生成文本嵌入向量...")
            self.root.update()
            
            # 使用tqdm创建进度条
            total = len(self.questions)
            
            # 批量处理以提高效率
            batch_size = 32
            embeddings = []
            
            for i in tqdm(range(0, len(self.questions), batch_size), desc="生成嵌入向量"):
                batch = self.questions[i:i+batch_size]
                batch_embeddings = self.model.encode(batch)
                embeddings.extend(batch_embeddings)
                
                # 更新进度条
                progress = min(100, int((i + len(batch)) / total * 100))
                self.progress_var.set(progress)
                self.root.update_idletasks()
            
            self.embeddings = np.array(embeddings)
            self.status_var.set("嵌入向量生成完成")
            return True
        
        except Exception as e:
            messagebox.showerror("错误", f"生成嵌入向量时出错: {str(e)}")
            return False
    
    def find_duplicates(self):
        try:
            if self.embeddings is None:
                return False
            
            self.status_var.set("计算相似度矩阵...")
            self.root.update()
            
            # 计算余弦相似度矩阵
            self.similarity_matrix = cosine_similarity(self.embeddings)
            
            # 设置对角线为0，避免自身匹配
            np.fill_diagonal(self.similarity_matrix, 0)
            
            # 查找相似度高于阈值的题目对
            threshold = self.threshold_var.get()
            self.duplicate_pairs = []
            
            total = len(self.questions)
            for i in tqdm(range(total), desc="查找重复题目"):
                # 只比较不同题目之间的相似度，避免自身比较
                for j in range(i+1, total):  # 从i+1开始，确保i≠j
                    if self.similarity_matrix[i, j] >= threshold:
                        self.duplicate_pairs.append((i, j, self.similarity_matrix[i, j]))
                
                # 更新进度条
                progress = min(100, int((i + 1) / total * 100))
                self.progress_var.set(progress)
                self.root.update_idletasks()
            
            # 按相似度降序排序
            self.duplicate_pairs.sort(key=lambda x: x[2], reverse=True)
            
            self.status_var.set(f"找到 {len(self.duplicate_pairs)} 对重复题目")
            return True
        
        except Exception as e:
            messagebox.showerror("错误", f"查找重复题目时出错: {str(e)}")
            return False
    
    def display_results(self):
        try:
            self.result_text.delete(1.0, tk.END)
            
            if not self.duplicate_pairs:
                self.result_text.insert(tk.END, "未找到重复题目\n")
                return
            
            self.result_text.insert(tk.END, f"找到 {len(self.duplicate_pairs)} 对重复题目:\n\n")
            
            for i, (idx1, idx2, similarity) in enumerate(self.duplicate_pairs):
                q_id1 = self.question_ids[idx1]
                q_id2 = self.question_ids[idx2]
                q_num1 = self.question_numbers[idx1]
                q_num2 = self.question_numbers[idx2]
                q1 = self.questions[idx1]
                q2 = self.questions[idx2]
                
                self.result_text.insert(tk.END, f"重复对 #{i+1} (相似度: {similarity:.4f}):\n")
                self.result_text.insert(tk.END, f"题目 {q_num1} (ID: {q_id1}): {q1[:100]}...\n")
                self.result_text.insert(tk.END, f"题目 {q_num2} (ID: {q_id2}): {q2[:100]}...\n\n")
            
            # 生成相似度分布图
            self.plot_similarity_distribution()
            
        except Exception as e:
            messagebox.showerror("错误", f"显示结果时出错: {str(e)}")
    
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
    
    def start_analysis(self):
        # 重置进度条
        self.progress_var.set(0)
        self.progress_bar.config(state=tk.NORMAL)
        
        # 清空之前的结果
        self.result_text.delete(1.0, tk.END)
        self.duplicate_pairs = []
        
        # 执行分析流程
        if not self.load_data():
            self.progress_bar.config(state=tk.DISABLED)
            return
        
        if not self.load_model():
            self.progress_bar.config(state=tk.DISABLED)
            return
        
        if not self.generate_embeddings():
            self.progress_bar.config(state=tk.DISABLED)
            return
        
        if not self.find_duplicates():
            self.progress_bar.config(state=tk.DISABLED)
            return
        
        # 显示结果
        self.display_results()
        
        # 分析完成
        self.progress_bar.config(state=tk.DISABLED)
        self.status_var.set("分析完成")

# 主程序入口
if __name__ == "__main__":
    root = tk.Tk()
    app = DuplicateQuestionChecker(root)
    root.mainloop()