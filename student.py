import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from PIL import Image, ImageTk
import re
import datetime
import zipfile
import tempfile
import shutil
import random
import webbrowser
import base64
from io import BytesIO
import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except:
        pass

def resource_path(relative_path):
    """返回资源文件的绝对路径：
    打包成exe后从内置解压目录读取；开发时从脚本所在目录读取"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

class StudentRecordSystem:
    MAX_PHOTOS = 20  # 每个学生最多照片数量
    
    def __init__(self, root):
        self.root = root
        self.root.title("学生学籍管理系统")
        self.root.geometry("1600x1000")
        
        # 高DPI渲染设置
        try:
            self.root.tk.call('tk', 'scaling', 1.5)
        except:
            pass
        
        # 高质量图片渲染配置
        import tkinter as tkinter_module
        if hasattr(tkinter_module, 'PhotoImage'):
            pass
        
        # 设置数据文件路径 - 基于脚本/exe所在目录，避免受当前工作目录影响
        self.base_dir = resource_path("_internal")
        self.pic_folder = os.path.join(self.base_dir, "pic")
        self.data_file1 = os.path.join(self.base_dir, "studentinfo", "students_data.xlsx")
        
        # 检查文件夹是否存在，如果不存在则创建
        self.setup_directories()
        
        # 当前选中的学生信息
        self.current_student = None
        
        # 多页照片相关
        self.current_photo_index = 0
        self.current_photo_paths = []
        
        # 初始化student_photos为空列表，避免AttributeError
        self.student_photos = []
        
        # 数据框
        self.df1 = pd.DataFrame()
        
        # 加载数据
        self.load_data()
        
        # 配置全局样式
        self.setup_styles()
        
        # 创建GUI组件
        self.create_widgets()
    
    def setup_styles(self):
        """配置全局界面样式，统一美观的主题"""
        self.colors = {
            'bg': '#eef1f6',
            'card_bg': '#ffffff',
            'header_bg': '#1e3a5f',
            'header_fg': '#ffffff',
            'primary': '#2f6fb3',
            'primary_dark': '#1e4f86',
            'hover': '#3d82cc',
            'danger': '#c0392b',
            'danger_dark': '#96281b',
            'danger_hover': '#d64541',
            'text': '#2b3a4a',
            'muted': '#7b8ca0',
            'border': '#c9d3e0',
            'selection': '#c8dcf5',
            'row_odd': '#f3f6fb',
            'row_even': '#ffffff',
        }
        
        c = self.colors
        self.root.configure(bg=c['bg'])
        
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass
        
        # 框架
        self.style.configure('TFrame', background=c['bg'])
        self.style.configure('Header.TFrame', background=c['header_bg'])
        self.style.configure('Card.TFrame', background=c['card_bg'])
        
        # 标签
        self.style.configure('TLabel', background=c['bg'], foreground=c['text'], font=('Microsoft YaHei', 11))
        self.style.configure('Card.TLabel', background=c['card_bg'], foreground=c['text'], font=('Microsoft YaHei', 11))
        self.style.configure('Muted.TLabel', background=c['card_bg'], foreground=c['muted'], font=('Microsoft YaHei', 10))
        self.style.configure('Title.TLabel', background=c['header_bg'], foreground=c['header_fg'], font=('Microsoft YaHei', 17, 'bold'))
        self.style.configure('Header.TLabel', background=c['header_bg'], foreground='#cfe0f2', font=('Microsoft YaHei', 11))
        
        # 按钮
        self.style.configure('TButton', font=('Microsoft YaHei', 11), padding=(14, 7))
        self.style.configure('Toolbar.TButton', font=('Microsoft YaHei', 11), padding=(14, 7),
                             background=c['card_bg'], foreground=c['text'], bordercolor=c['border'])
        self.style.map('Toolbar.TButton',
                       background=[('pressed', c['selection']), ('active', '#e7eef7')],
                       foreground=[('active', c['primary'])])
        self.style.configure('Primary.TButton', font=('Microsoft YaHei', 11), padding=(18, 7),
                             background=c['primary'], foreground=c['header_fg'], bordercolor=c['primary'])
        self.style.map('Primary.TButton',
                       background=[('pressed', c['primary_dark']), ('active', c['hover'])])
        self.style.configure('Danger.TButton', font=('Microsoft YaHei', 11), padding=(14, 7),
                             background=c['danger'], foreground=c['header_fg'], bordercolor=c['danger'])
        self.style.map('Danger.TButton',
                       background=[('pressed', c['danger_dark']), ('active', c['danger_hover'])])
        
        # 单选按钮
        self.style.configure('Toolbar.TRadiobutton', background=c['card_bg'], foreground=c['text'], font=('Microsoft YaHei', 11))
        self.style.map('Toolbar.TRadiobutton',
                       background=[('active', c['card_bg'])],
                       foreground=[('selected', c['primary'])])
        
        # 输入框
        self.style.configure('TEntry', fieldbackground=c['card_bg'], foreground=c['text'],
                             bordercolor=c['border'], lightcolor=c['border'], darkcolor=c['border'])
        self.style.map('TEntry', fieldbackground=[('focus', '#ffffff')])
        
        # 分组框
        self.style.configure('Card.TLabelframe', background=c['card_bg'], bordercolor=c['border'])
        self.style.configure('Card.TLabelframe.Label', background=c['card_bg'], foreground=c['primary'], font=('Microsoft YaHei', 12, 'bold'))
        
        # 树形视图
        self.style.configure('Treeview', font=('Microsoft YaHei', 11), rowheight=30,
                             background=c['card_bg'], fieldbackground=c['card_bg'], bordercolor=c['border'])
        self.style.map('Treeview',
                       background=[('selected', c['primary'])],
                       foreground=[('selected', c['header_fg'])])
        self.style.configure('Treeview.Heading', font=('Microsoft YaHei', 12, 'bold'),
                             background=c['primary'], foreground=c['header_fg'], relief=tk.FLAT)
        self.style.map('Treeview.Heading', background=[('active', c['hover'])])
        
        # 滚动条
        self.style.configure('Vertical.TScrollbar', background=c['border'], troughcolor=c['bg'],
                             bordercolor=c['bg'], arrowcolor=c['muted'])
        self.style.configure('Horizontal.TScrollbar', background=c['border'], troughcolor=c['bg'],
                             bordercolor=c['bg'], arrowcolor=c['muted'])
    
    def setup_directories(self):
        """创建必要的目录结构"""
        try:
            # 创建_internal主目录
            if not os.path.exists(self.base_dir):
                os.makedirs(self.base_dir)
                print(f"创建目录: {self.base_dir}")
            
            # 创建pic目录
            if not os.path.exists(self.pic_folder):
                os.makedirs(self.pic_folder)
                print(f"创建目录: {self.pic_folder}")
            
            # 创建studentinfo目录
            studentinfo_dir = os.path.dirname(self.data_file1)
            if not os.path.exists(studentinfo_dir):
                os.makedirs(studentinfo_dir)
                print(f"创建目录: {studentinfo_dir}")
                
        except Exception as e:
            print(f"创建目录时出错: {str(e)}")
    
    def load_data(self):
        """加载学生数据"""
        try:
            print("开始加载数据...")
            
            # 1. 加载图片信息（作为照片匹配用）
            print("正在加载图片信息...")
            self.student_photos = self.extract_photo_info()
            print(f"成功加载 {len(self.student_photos)} 个学生的图片信息")
            
            # 2. 加载Excel数据（作为主数据库）
            print("正在加载Excel数据...")
            self.df1 = pd.DataFrame()
            
            # 加载第一个Excel文件
            if os.path.exists(self.data_file1):
                try:
                    temp_df = pd.read_excel(self.data_file1)
                    str_cols = {}
                    for col in temp_df.columns:
                        col_lower = str(col).lower()
                        if any(kw in col_lower for kw in ['身份证', '学籍号', '学号', '手机', '电话']):
                            str_cols[col] = str
                    
                    if str_cols:
                        self.df1 = pd.read_excel(self.data_file1, dtype=str_cols)
                    else:
                        self.df1 = pd.read_excel(self.data_file1)
                    print(f"成功加载数据文件1: {self.data_file1}")
                    print(f"数据表1形状: {self.df1.shape}")
                except Exception as e:
                    print(f"加载数据文件1时出错: {str(e)}")
                    try:
                        self.df1 = pd.read_excel(self.data_file1, engine='openpyxl')
                        print("使用openpyxl引擎成功加载数据文件1")
                    except:
                        print("无法加载数据文件1")
            else:
                print(f"数据文件不存在: {self.data_file1}")
            
            # 3. 从Excel表格构建学生列表
            self.build_student_list()
            
            total_students = len(self.df1)
            print(f"数据加载完成，共 {total_students} 个学生")
            
        except Exception as e:
            print(f"加载数据时发生错误: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("数据加载错误", f"无法加载数据文件: {str(e)}")
    
    def build_student_list(self):
        """从Excel表格构建学生列表（以表格为主数据库）"""
        self.table_students = []
        seen_students = set()
        used_photos = set()
        
        # ============================================================
        # 第一步：从Excel表格中提取学生信息
        # ============================================================
        if not self.df1.empty:
            print("正在从Excel表格提取学生信息...")
            for idx, row in self.df1.iterrows():
                try:
                    student = self._extract_student_from_row(row, "表1")
                    if student:
                        name = student.get('姓名', '')
                        student_id = student.get('学籍号', '')
                        id_card = student.get('身份证号', '')
                        
                        if not name:
                            continue
                        
                        # 仅用身份证号/学籍号去重，无ID的行各自独立（避免同名合并）
                        unique_key = f"idcard:{id_card}" if id_card else (f"studentid:{student_id}" if student_id else f"row:{row.name}")
                        
                        if unique_key in seen_students:
                            continue
                        
                        seen_students.add(unique_key)
                        
                        photo_path, extra_photos = self.find_photo_for_student(student, used_photos)
                        if photo_path:
                            used_photos.add(photo_path)
                        student['照片路径'] = photo_path
                        student['extra_photos'] = extra_photos
                        student['has_photo'] = photo_path is not None
                        student['数据来源'] = 'Excel表格'
                        self.table_students.append(student)
                except Exception as e:
                    print(f"处理第 {idx} 行时出错: {str(e)}")
                    continue
        
        print(f"从Excel表格构建学生列表完成，共 {len(self.table_students)} 个学生")
        
        # ============================================================
        # 第二步：从照片文件夹中补充Excel中没有的学生信息
        # 用于处理Excel表格中信息为空但照片文件名中有信息的学生
        # ============================================================
        print("正在从照片文件夹补充学生信息...")
        
        # 将已处理的学生信息转换为用于匹配的字典
        excel_students_info = {}
        for student in self.table_students:
            id_card = student.get('身份证号', '')
            student_id = student.get('学籍号', '')
            national_id = student.get('全国学籍号', '')
            if not national_id and id_card:
                national_id = "G" + str(id_card)
            
            if id_card and id_card != '未知身份证号':
                excel_students_info[id_card] = student
            if student_id and student_id != '未知学籍号':
                excel_students_info[student_id] = student
                excel_students_info[re.sub(r'^[GLP]', '', str(student_id))] = student
            if national_id and national_id != '未知全国学籍号':
                excel_students_info[national_id] = student
                excel_students_info[re.sub(r'^[GLP]', '', str(national_id))] = student
        
        # 遍历所有照片，查找未被Excel表格覆盖的学生
        for photo_info in self.student_photos:
            photo_path = photo_info.get('filename', '')
            
            # 跳过已使用的照片
            if photo_path in used_photos:
                continue
            
            photo_name = photo_info.get('姓名', '')
            photo_id = photo_info.get('学籍号', '')
            photo_idcard = photo_info.get('身份证号', '')
            photo_national = photo_info.get('全国学籍号', '')
            
            # 检查该学生是否已经在Excel表格中
            matched_excel_student = None
            
            # 通过身份证号检查
            if photo_idcard and photo_idcard != '未知身份证号':
                if photo_idcard in excel_students_info:
                    matched_excel_student = excel_students_info[photo_idcard]
            
            # 通过全国学籍号检查
            if not matched_excel_student and photo_national and photo_national != '未知全国学籍号':
                for cand in (photo_national, re.sub(r'^[GLP]', '', photo_national)):
                    if cand in excel_students_info:
                        matched_excel_student = excel_students_info[cand]
                        break
            
            # 通过学籍号检查
            if not matched_excel_student and photo_id and photo_id != '未知学籍号':
                for cand in (photo_id, re.sub(r'^[GLP]', '', photo_id)):
                    if cand in excel_students_info:
                        matched_excel_student = excel_students_info[cand]
                        break
            
            if matched_excel_student:
                # 该照片属于已有学生，将其添加到额外照片列表中
                if photo_path not in matched_excel_student.get('extra_photos', []):
                    matched_excel_student.setdefault('extra_photos', []).append(photo_path)
                used_photos.add(photo_path)
                continue
            
            # 如果该学生不在Excel表格中且照片含有学籍号/身份证号/全国学籍号，则从照片中提取信息添加到列表
            if ((photo_id and photo_id != '未知学籍号') or
                (photo_idcard and photo_idcard != '未知身份证号') or
                (photo_national and photo_national != '未知全国学籍号')):
                # 生成唯一键
                if photo_idcard:
                    unique_key = f"idcard:{photo_idcard}"
                elif photo_id:
                    unique_key = f"studentid:{photo_id}"
                else:
                    unique_key = f"national:{photo_national}"
                
                if unique_key not in seen_students:
                    # 创建新的学生信息
                    new_student = {
                        'table': '照片补充',
                        'row_index': -1,
                        '姓名': photo_name,
                        '学籍号': photo_id if photo_id else '',
                        '身份证号': photo_idcard if photo_idcard else '',
                        '全国学籍号': photo_national if photo_national else '',
                        '照片路径': photo_path,
                        'extra_photos': photo_info.get('extra_photos', []),
                        'has_photo': True,
                        '数据来源': '照片文件名'
                    }
                    seen_students.add(unique_key)
                    used_photos.add(photo_path)
                    self.table_students.append(new_student)
                    print(f"从照片补充学生: {photo_name}, 学籍号: {photo_id}, 身份证号: {photo_idcard}, 全国学籍号: {photo_national}")
        
        print(f"最终学生列表共 {len(self.table_students)} 个学生")
    
    def _extract_student_from_row(self, row, table_name):
        """从表格行提取学生信息"""
        student = {'table': table_name, 'row_index': row.name}
        
        name = None
        student_id = None
        id_card = None
        
        for col in row.index:
            col_name = str(col).strip()
            col_lower = col_name.lower()
            value = str(row[col]).strip() if pd.notna(row[col]) else ''
            if not value or value == 'nan' or value == 'None':
                continue
            
            # 格式化值
            formatted_value = self.format_cell_value(col, row[col])
            if formatted_value:
                value = formatted_value
            
            # 查找姓名
            if not name and any(kw in col_lower for kw in ['姓名', '名字', 'name']):
                name = value
            
            # 查找学籍号
            if not student_id and any(kw in col_lower for kw in ['学籍号', '学号', 'student_id']):
                student_id = value
            
            # 查找身份证号
            if not id_card and any(kw in col_lower for kw in ['身份证', 'idcard']):
                id_card = value
            
            # 存储所有字段（不添加表名前缀）
            student[col_name] = value
        
        if name:
            student['姓名'] = name
            student['学籍号'] = student_id if student_id else ''
            student['身份证号'] = id_card if id_card else ''
            return student
        
        return None
    
    def _is_name_duplicated(self, name):
        """判断姓名在信息库中是否重复（重复则不能用姓名匹配）"""
        if not name or name == '未知姓名':
            return True
        count = 0
        if self.df1 is not None and not self.df1.empty:
            for col in self.df1.columns:
                if any(kw in str(col).lower() for kw in ['姓名', '名字', 'name']):
                    try:
                        count += int((self.df1[col] == name).sum())
                    except Exception:
                        pass
        return count > 1
    
    def build_photo_index(self):
        """为照片建立编号索引，加速匹配（身份证号/学籍号/全国学籍号/姓名）"""
        self.photo_index = {}
        for photo in self.student_photos:
            ids = []
            if photo.get('身份证号', '未知身份证号') != '未知身份证号':
                ids.append(('idc', photo['身份证号']))
            if photo.get('学籍号', '未知学籍号') != '未知学籍号':
                ids.append(('sid', photo['学籍号']))
            if photo.get('全国学籍号', '未知全国学籍号') != '未知全国学籍号':
                ids.append(('nat', photo['全国学籍号']))
            if photo.get('姓名', '未知姓名') != '未知姓名':
                ids.append(('name', photo['姓名']))
            
            for kind, val in ids:
                val_u = str(val).upper()
                self.photo_index.setdefault((kind, val_u), []).append(photo)
                nv = re.sub(r'^[GLP]', '', val_u)
                if nv != val_u:
                    self.photo_index.setdefault((kind, nv), []).append(photo)
    
    def _find_unused_photo(self, kind, val, used_photos):
        """从索引中查找指定编号未被使用过的照片"""
        for photo in self.photo_index.get((kind, str(val).upper()), []):
            if photo.get('filename', '') not in used_photos:
                return photo
        return None
    
    def find_photo_for_student(self, student, used_photos=None):
        """根据学生信息查找对应的照片，返回 (主照片路径, 额外照片列表)
        支持 身份证号/全国学籍号/学籍号/唯一姓名 的任意组合命名匹配（索引加速）"""
        if used_photos is None:
            used_photos = set()
        
        if not hasattr(self, 'photo_index') or not self.photo_index:
            self.build_photo_index()
        
        stu_id = student.get('学籍号', '')
        stu_idcard = student.get('身份证号', '')
        stu_national = student.get('全国学籍号', '')
        if not stu_national and stu_idcard:
            stu_national = "G" + str(stu_idcard)
        
        # 按优先级构建查询候选（含全国学籍号与身份证号/学籍号互换的情况）
        candidates = []
        
        def add_cand(kind, val):
            if not val or str(val).strip() in ('', 'nan', 'None', '未知身份证号', '未知学籍号', '未知全国学籍号'):
                return
            candidates.append((kind, str(val).upper()))
        
        add_cand('idc', stu_idcard)      # 照片身份证号 == 学生身份证号
        add_cand('nat', stu_national)    # 照片全国学籍号 == 学生全国学籍号
        add_cand('nat', stu_idcard)      # 照片全国学籍号 == 学生身份证号
        add_cand('nat', stu_id)          # 照片全国学籍号 == 学生学籍号
        add_cand('sid', stu_id)          # 照片学籍号 == 学生学籍号
        add_cand('sid', stu_national)    # 照片学籍号 == 学生全国学籍号
        
        for kind, val in candidates:
            photo = self._find_unused_photo(kind, val, used_photos)
            if photo is not None:
                return photo.get('filename', ''), photo.get('extra_photos', [])
        
        # 姓名匹配（仅在信息库中该姓名唯一时）
        stu_name = student.get('姓名', '')
        if stu_name and stu_name != '未知姓名' and not self._is_name_duplicated(stu_name):
            photo = self._find_unused_photo('name', stu_name, used_photos)
            if photo is not None:
                return photo.get('filename', ''), photo.get('extra_photos', [])
        
        return None, []
    
    def extract_photo_info(self):
        """从图片文件名中提取学生信息，支持多页照片"""
        photo_info = []
        try:
            if os.path.exists(self.pic_folder):
                print(f"正在扫描文件夹: {self.pic_folder}")
                
                # 第一步：解析所有图片文件，提取学生信息
                all_parsed = []
                for filename in os.listdir(self.pic_folder):
                    if filename.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif')):
                        info = self.parse_filename(filename)
                        if info:
                            info['filename'] = os.path.join(self.pic_folder, filename)
                            info['original_filename'] = filename
                            all_parsed.append(info)
                            print(f"解析文件: {filename} -> {info.get('姓名')}, 学籍号={info.get('学籍号')}, 身份证号={info.get('身份证号')}")
                
                # 第二步：按学生身份（学籍号或身份证号）分组，将同一学生的多个照片合并
                used = set()
                # 按文件名排序，确保不含 _数字 的文件优先作为主照片
                all_parsed.sort(key=lambda x: (0 if re.search(r'_\d+(\.\w+)?$', x.get('original_filename', '')) else -1, x.get('original_filename', '')))
                for info in all_parsed:
                    filepath = info['filename']
                    if filepath in used:
                        continue
                    
                    student_id = info.get('学籍号', '')
                    id_card = info.get('身份证号', '')
                    national_id = info.get('全国学籍号', '')
                    
                    # 确定该学生的唯一标识（按学籍号/身份证号/全国学籍号/姓名）
                    if student_id and student_id != '未知学籍号':
                        student_key = f"sid:{student_id}"
                    elif id_card and id_card != '未知身份证号':
                        student_key = f"idc:{id_card}"
                    elif national_id and national_id != '未知全国学籍号':
                        student_key = f"nat:{national_id}"
                    else:
                        # 仅姓名的照片按姓名分组（匹配时仍会校验姓名是否唯一）
                        student_key = f"name:{info.get('姓名', '')}"
                    
                    info['extra_photos'] = []
                    used.add(filepath)
                    
                    # 查找同一学生的其他照片（通过学籍号或身份证号匹配）
                    for other in all_parsed:
                        other_path = other['filename']
                        if other_path in used:
                            continue
                        other_sid = other.get('学籍号', '')
                        other_idc = other.get('身份证号', '')
                        if (student_id and student_id != '未知学籍号' and student_id == other_sid) or \
                           (id_card and id_card != '未知身份证号' and id_card == other_idc):
                            info['extra_photos'].append(other_path)
                            used.add(other_path)
                    
                    # 按文件名排序 extra_photos，保证顺序一致
                    info['extra_photos'].sort()
                    
                    photo_info.append(info)
                    print(f"分组: {info.get('姓名')} 主照片={info['original_filename']}, extra_photos数量={len(info['extra_photos'])}")
            else:
                print(f"警告: 图片文件夹不存在: {self.pic_folder}")
        except Exception as e:
            print(f"提取照片信息时出错: {str(e)}")
            import traceback
            traceback.print_exc()
            photo_info = []
        # 重建照片索引，供后续快速匹配
        self.student_photos = photo_info
        self.build_photo_index()
        return photo_info
    
    def parse_filename(self, filename):
        """解析图片文件名，提取姓名/学籍号/身份证号/全国学籍号
        支持 姓名/学籍号/身份证号/全国学籍号 的任意组合与任意顺序"""
        # 移除文件扩展名
        name_without_ext = os.path.splitext(filename)[0]
        # 移除文件名末尾的 _数字 后缀（多页照片后缀）
        name_without_ext = re.sub(r'_\d+$', '', name_without_ext)
        print(f"解析文件名(去除多页后缀后): {name_without_ext}")
        
        name = ""
        student_id = ""
        id_card = ""
        national_id = ""
        
        # 1. 全国学籍号（优先识别带标签的，如 全国学籍号：G110101...）
        m = re.search(r'全国学籍号\s*[:：]?\s*([A-Za-z]?\d{17,20})', name_without_ext)
        if m:
            national_id = m.group(1).upper()
        if not national_id:
            m = re.search(r'(?<![A-Za-z0-9])([GLP]\d{18})(?![A-Za-z0-9])', name_without_ext, re.IGNORECASE)
            if m:
                national_id = m.group(1).upper()
        
        # 2. 身份证号（18位末位可X，或15位；优先识别带标签的）
        m = re.search(r'身份证号\s*[:：]?\s*(\d{17}[\dXx]|\d{15})', name_without_ext)
        if m:
            id_card = m.group(1).upper()
        if not id_card:
            m = re.search(r'(?<!\d)(\d{17}[\dXx])(?!\d)', name_without_ext)
            if m:
                id_card = m.group(1).upper()
        if not id_card:
            m = re.search(r'(?<!\d)(\d{15})(?!\d)', name_without_ext)
            if m:
                id_card = m.group(1).upper()
        
        # 3. 学籍号（10位以上；优先识别带标签的，可能带G/L前缀）
        m = re.search(r'学籍号\s*[:：]?\s*([A-Za-z]?\d{10,})', name_without_ext)
        if m:
            student_id = m.group(1).upper()
        if not student_id:
            numbers = re.findall(r'(?<!\d)(\d{10,})(?!\d)', name_without_ext)
            candidates = [n for n in numbers if n != id_card]
            if national_id:
                nat_digits = re.sub(r'^[A-Za-z]', '', national_id)
                candidates = [n for n in candidates if n != nat_digits]
            if candidates:
                student_id = max(candidates, key=len)
        
        # 4. 姓名：剔除所有编号和字段标签后提取中文（2-4字）
        remaining = name_without_ext
        remaining = re.sub(r'(?<!\d)(\d{17}[\dXx]|\d{15})(?!\d)', ' ', remaining)
        remaining = re.sub(r'(?<![A-Za-z0-9])[GLP]\d{18}(?![A-Za-z0-9])', ' ', remaining, flags=re.IGNORECASE)
        remaining = re.sub(r'(?<!\d)\d{10,}(?!\d)', ' ', remaining)
        remaining = re.sub(r'全国学籍号|省学籍辅号|现学籍号|身份证号码|身份证号|学籍号|学号|姓名|号码|辅号|[:：_\-—.()（）]', ' ', remaining)
        chinese_matches = re.findall(r'[\u4e00-\u9fff]{2,4}', remaining)
        if chinese_matches:
            name = chinese_matches[0]
        
        print(f"解析结果 - 姓名: {name}, 学籍号: {student_id}, 身份证号: {id_card}, 全国学籍号: {national_id}")
        
        return {
            '姓名': name if name else "未知姓名",
            '学籍号': student_id if student_id else "未知学籍号",
            '身份证号': id_card if id_card else "未知身份证号",
            '全国学籍号': national_id if national_id else "未知全国学籍号"
        }
    
    def create_widgets(self):
        """创建GUI界面（主列表页 + 学生详情页 双页面跳转）"""
        c = self.colors
        
        # ========== 页面容器 ==========
        self.page_container = ttk.Frame(self.root)
        self.page_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.page_container.columnconfigure(0, weight=1)
        self.page_container.rowconfigure(0, weight=1)
        
        # ================================================
        # 页面1：主列表页
        # ================================================
        self.page_main = ttk.Frame(self.page_container)
        self.page_main.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        header = ttk.Frame(self.page_main, style='Header.TFrame')
        header.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        title_label = ttk.Label(header, text="学生学籍管理系统", style='Title.TLabel')
        title_label.pack(side=tk.LEFT, padx=20, pady=12)
        
        info_bar = ttk.Frame(header, style='Header.TFrame')
        info_bar.pack(side=tk.RIGHT, padx=20)
        
        self.user_label = ttk.Label(info_bar, text="用户: Admin", style='Header.TLabel')
        self.user_label.pack(side=tk.LEFT, padx=15)
        
        self.status_label = ttk.Label(info_bar, text="● 就绪", style='Header.TLabel')
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        main_frame = ttk.Frame(self.page_main, padding="15")
        main_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # ---------- 搜索区域 ----------
        search_frame = ttk.LabelFrame(main_frame, text=" 搜索学生 ", padding="12", style='Card.TLabelframe')
        search_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 12))
        
        ttk.Label(search_frame, text="关键词:", style='Card.TLabel').grid(row=0, column=0, padx=(5, 8))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=42,
                                      font=('Microsoft YaHei', 11))
        self.search_entry.grid(row=0, column=1, padx=(0, 15), pady=3)
        
        self.search_type = tk.StringVar(value="姓名")
        ttk.Radiobutton(search_frame, text="姓名", variable=self.search_type, value="姓名", style='Toolbar.TRadiobutton').grid(row=0, column=2, padx=3)
        ttk.Radiobutton(search_frame, text="学籍号", variable=self.search_type, value="学籍号", style='Toolbar.TRadiobutton').grid(row=0, column=3, padx=3)
        ttk.Radiobutton(search_frame, text="身份证号", variable=self.search_type, value="身份证号", style='Toolbar.TRadiobutton').grid(row=0, column=4, padx=3)
        
        ttk.Button(search_frame, text="搜  索", command=self.search_student, style='Primary.TButton').grid(row=0, column=5, padx=(15, 5))
        ttk.Button(search_frame, text="显示所有", command=self.show_all_students, style='Toolbar.TButton').grid(row=0, column=6, padx=5)
        ttk.Button(search_frame, text="清  空", command=self.clear_search, style='Toolbar.TButton').grid(row=0, column=7, padx=5)
        
        # 绑定回车键
        self.search_entry.bind('<Return>', lambda e: self.search_student())
        
        # ---------- 学生列表区域 ----------
        list_frame = ttk.LabelFrame(main_frame, text=" 学生列表 ", padding="10", style='Card.TLabelframe')
        list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        columns = ('序号', '姓名', '学籍号', '身份证号')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=12, style='Treeview')
        
        # 排序状态
        self.sort_column = None
        self.sort_reverse = False
        
        column_widths = {'序号': 80, '姓名': 160, '学籍号': 220, '身份证号': 220}
        for col in columns:
            self.tree.heading(col, text=col, anchor=tk.CENTER,
                              command=lambda c=col: self.sort_treeview(c))
            self.tree.column(col, width=column_widths.get(col, 150), anchor=tk.CENTER)
        
        self.tree.tag_configure('odd', background=c['row_odd'])
        self.tree.tag_configure('even', background=c['row_even'])
        
        tree_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        # ---------- 底部操作按钮（仅列表级操作） ----------
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=(12, 0))
        
        actions = [
            ("重载", self.reload_data),
            ("导入表格", self.import_xlsx),
            ("打开文件夹", self.open_data_folder),
            ("随机展示", self.show_random_student),
            ("退出", self.root.quit),
        ]
        for text, cmd in actions:
            btn_style = 'Danger.TButton' if text == "退出" else 'Toolbar.TButton'
            ttk.Button(button_frame, text=text, command=cmd, style=btn_style).pack(side=tk.LEFT, padx=5)
        
        # 配置主页面网格权重
        self.page_main.columnconfigure(0, weight=1)
        self.page_main.rowconfigure(1, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # ================================================
        # 页面2：学生详情页
        # ================================================
        self.page_detail = ttk.Frame(self.page_container)
        self.page_detail.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        detail_header = ttk.Frame(self.page_detail, style='Header.TFrame')
        detail_header.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # 左上角返回按钮
        self.back_btn = ttk.Button(detail_header, text="◀ 返回", command=self.show_main_page, style='Toolbar.TButton')
        self.back_btn.pack(side=tk.LEFT, padx=10, pady=8)
        
        self.detail_title = ttk.Label(detail_header, text="学生详情", style='Title.TLabel')
        self.detail_title.pack(side=tk.LEFT, padx=8)
        
        detail_main = ttk.Frame(self.page_detail, padding="15")
        detail_main.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        content_frame = ttk.Frame(detail_main)
        content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 左侧信息卡片
        info_frame = ttk.LabelFrame(content_frame, text=" 学生详细信息 ", padding="10", style='Card.TLabelframe')
        info_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 12))
        
        text_frame = ttk.Frame(info_frame, style='Card.TFrame')
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        v_scrollbar = ttk.Scrollbar(text_frame)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.info_text = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=v_scrollbar.set,
                                width=85, height=32, font=('Microsoft YaHei', 12),
                                bg=c['card_bg'], fg=c['text'], relief=tk.FLAT,
                                padx=12, pady=12, selectbackground=c['selection'],
                                insertbackground=c['primary'], highlightthickness=0)
        self.info_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        v_scrollbar.config(command=self.info_text.yview)
        
        # 右侧照片卡片
        self.image_frame = ttk.LabelFrame(content_frame, text=" 学生照片 ", padding="10", style='Card.TLabelframe')
        self.image_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(12, 0))
        
        self.photo_canvas = tk.Canvas(self.image_frame, width=460, height=460, bg='#f8fafc',
                                     highlightthickness=2, highlightbackground=c['primary'])
        self.photo_canvas.grid(row=0, column=0, padx=5, pady=5)
        
        # 绑定双击事件查看大图
        self.photo_canvas.bind("<Double-Button-1>", self.show_large_image)
        
        self.photo_label = ttk.Label(self.image_frame, text="双击图片查看大图", style='Muted.TLabel')
        self.photo_label.grid(row=1, column=0, pady=(5, 0))
        
        # 多页照片导航按钮
        nav_frame = ttk.Frame(self.image_frame, style='Card.TFrame')
        nav_frame.grid(row=2, column=0, pady=(5, 0))
        self.photo_prev_btn = ttk.Button(nav_frame, text="◀ 上一张", command=self.photo_prev, state="disabled", style='Toolbar.TButton')
        self.photo_prev_btn.pack(side=tk.LEFT, padx=6)
        self.photo_page_label = ttk.Label(nav_frame, text="", style='Muted.TLabel')
        self.photo_page_label.pack(side=tk.LEFT, padx=6)
        self.photo_next_btn = ttk.Button(nav_frame, text="下一张 ▶", command=self.photo_next, state="disabled", style='Toolbar.TButton')
        self.photo_next_btn.pack(side=tk.LEFT, padx=6)
        
        # 照片管理按钮
        manage_frame = ttk.Frame(self.image_frame, style='Card.TFrame')
        manage_frame.grid(row=3, column=0, pady=(8, 0))
        ttk.Button(manage_frame, text="增加照片", command=self.add_photo, style='Toolbar.TButton').pack(side=tk.LEFT, padx=6)
        ttk.Button(manage_frame, text="删除照片", command=self.delete_photo, style='Toolbar.TButton').pack(side=tk.LEFT, padx=6)
        self.photo_count_label = ttk.Label(manage_frame, text="", style='Muted.TLabel')
        self.photo_count_label.pack(side=tk.LEFT, padx=10)
        
        # 详情页底部操作按钮（单学生相关操作）
        detail_actions_frame = ttk.Frame(detail_main)
        detail_actions_frame.grid(row=1, column=0, pady=(12, 0))
        
        detail_actions = [
            ("修正信息", self.manual_correct_info),
            ("打印", self.print_student_info),
            ("导出当前信息", self.export_current_zip),
            ("导出所有信息", self.export_all_zip),
        ]
        for text, cmd in detail_actions:
            ttk.Button(detail_actions_frame, text=text, command=cmd, style='Toolbar.TButton').pack(side=tk.LEFT, padx=5)
        
        # 配置详情页网格权重
        self.page_detail.columnconfigure(0, weight=1)
        self.page_detail.rowconfigure(1, weight=1)
        detail_main.columnconfigure(0, weight=1)
        detail_main.rowconfigure(0, weight=1)
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)
        info_frame.columnconfigure(0, weight=1)
        info_frame.rowconfigure(0, weight=1)
        
        # 默认显示主列表页
        self.page_main.tkraise()
        
        # 初始显示所有学生
        self.show_all_students()
    
    def show_main_page(self):
        """返回主列表页"""
        self.page_main.tkraise()
    
    def print_student_info(self):
        """打印学生信息"""
        if not self.current_student:
            messagebox.showwarning("警告", "请先选择一个学生")
            return
        
        # 创建打印选项对话框
        print_dialog = tk.Toplevel(self.root)
        print_dialog.title("打印选项")
        print_dialog.geometry("700x560")
        print_dialog.transient(self.root)
        
        # 高DPI渲染
        try:
            print_dialog.tk.call('tk', 'scaling', 1.5)
        except:
            pass
        
        # 居中显示
        screen_width = print_dialog.winfo_screenwidth()
        screen_height = print_dialog.winfo_screenheight()
        window_width = 700
        window_height = 560
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        print_dialog.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        
        # 获取学生信息
        name = self.current_student.get('姓名', '未知姓名')
        
        # 创建界面
        ttk.Label(print_dialog, text=f"打印学生: {name}", 
                 font=('Microsoft YaHei', 14, 'bold')).pack(pady=(20, 10))
        
        # 打印选项框架
        options_frame = ttk.Frame(print_dialog, padding="20")
        options_frame.pack(fill=tk.BOTH, expand=True)
        
        # 打印选项变量
        print_text = tk.BooleanVar(value=True)
        print_photo = tk.BooleanVar(value=True)  # 默认打印照片
        print_detailed_info = tk.BooleanVar(value=True)
        
        # 创建选项复选框
        ttk.Checkbutton(options_frame, text="打印基本信息", variable=print_text).pack(anchor=tk.W, pady=5)
        ttk.Checkbutton(options_frame, text="打印学籍照片（将位于最上方）", variable=print_photo).pack(anchor=tk.W, pady=5)
        ttk.Checkbutton(options_frame, text="打印详细信息", variable=print_detailed_info).pack(anchor=tk.W, pady=5)
        
        # 打印说明
        note_label = ttk.Label(options_frame, 
                              text="注：打印将生成HTML文件，在浏览器中打开后可进行打印操作。",
                              foreground="blue", font=('Microsoft YaHei', 9))
        note_label.pack(anchor=tk.W, pady=(10, 0))
        
        # 照片尺寸选项
        photo_size_frame = ttk.Frame(options_frame)
        photo_size_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(photo_size_frame, text="照片尺寸:").pack(side=tk.LEFT)
        photo_size_var = tk.StringVar(value="中等")
        size_combo = ttk.Combobox(photo_size_frame, textvariable=photo_size_var, 
                                  values=["小", "中等", "大"], state="readonly", width=10)
        size_combo.pack(side=tk.LEFT, padx=(10, 0))
        
        # 打印函数
        def execute_print():
            """执行打印操作"""
            # 获取用户选择的选项
            print_text_val = print_text.get()
            print_photo_val = print_photo.get()
            print_detailed_info_val = print_detailed_info.get()
            photo_size = photo_size_var.get()
            
            print_dialog.destroy()
            
            try:
                # 显示打印进度
                self.status_label.config(text="正在准备打印...", foreground="blue")
                self.root.update()
                
                # 调用打印函数
                self.actual_print(print_text_val, print_photo_val, print_detailed_info_val, photo_size)
                
            except Exception as e:
                messagebox.showerror("打印失败", f"打印过程中出错: {str(e)}")
                self.status_label.config(text="打印失败", foreground="red")
        
        # 按钮框架
        button_frame = ttk.Frame(print_dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="立即打印", command=execute_print).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=print_dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def create_print_html(self, temp_dir, print_text, print_photo, print_detailed_info, photo_size):
        """创建打印HTML文件"""
        name = self.current_student.get('姓名', '未知姓名')
        photo_path = self.current_student.get('照片路径')
        
        safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
        html_filename = f"{safe_name}_学籍信息_打印版.html"
        html_path = os.path.join(temp_dir, html_filename)
        
        photo_sizes = {"小": "200", "中等": "300", "大": "400"}
        photo_width = photo_sizes.get(photo_size, "300")
        
        photo_html = ""
        if print_photo:
            all_photos = []
            if photo_path and os.path.exists(photo_path):
                all_photos.append(photo_path)
            extra = self.current_student.get('extra_photos', [])
            all_photos.extend([ep for ep in extra if os.path.exists(ep)])
            
            try:
                for pi, pp in enumerate(all_photos):
                    with open(pp, 'rb') as f:
                        photo_bytes = f.read()
                        photo_base64 = base64.b64encode(photo_bytes).decode('utf-8')
                    
                    _, ext = os.path.splitext(pp)
                    mime_type = "image/jpeg"
                    if ext.lower() == '.png':
                        mime_type = "image/png"
                    elif ext.lower() == '.gif':
                        mime_type = "image/gif"
                    elif ext.lower() == '.bmp':
                        mime_type = "image/bmp"
                    
                    photo_data = f'data:{mime_type};base64,{photo_base64}'
                    caption = f"{name} - 在校学籍照片" if pi == 0 else f"{name} - 照片{pi + 1}"
                    photo_html += f'''
                    <div class="photo-container">
                        <h2>{"学籍照片" if pi == 0 else f"照片 {pi + 1}"}</h2>
                        <img src="{photo_data}" alt="{name}照" class="student-photo">
                        <p class="photo-caption">{caption}</p>
                    </div>
                    '''
            except Exception as e:
                print(f"转换照片为base64时出错: {e}")
                photo_html = ""
        
        # 获取统一格式的详细信息
        content = self.get_display_content(self.current_student)
        # 过滤掉标题和分隔线，只保留信息行
        info_lines = [line for line in content if not line.startswith('=') and '详细信息' not in line]
        info_rows = ''.join([f'<tr><th>{line.split("：")[0]}</th><td>{"：".join(line.split("：")[1:]) if "：" in line else ""}</td></tr>' for line in info_lines if '：' in line])
        
        html_content = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{name} - 学籍信息打印版</title>
    <style>
        body {{
            font-family: 'Microsoft YaHei', 'SimHei', sans-serif;
            margin: 20px;
            line-height: 1.6;
            color: #333;
        }}
        .header {{
            text-align: center;
            border-bottom: 3px double #333;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }}
        .header h1 {{
            margin: 0;
            color: #0066cc;
        }}
        .photo-container {{
            text-align: center;
            margin: 20px 0;
            page-break-inside: avoid;
        }}
        .student-photo {{
            max-width: {photo_width}px;
            border: 2px solid #333;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        }}
        .photo-caption {{
            font-style: italic;
            color: #666;
            margin-top: 5px;
        }}
        .section {{
            margin-bottom: 20px;
            page-break-inside: avoid;
        }}
        .section h2 {{
            color: #0066cc;
            border-bottom: 1px solid #ccc;
            padding-bottom: 5px;
        }}
        .info-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
        }}
        .info-table th, .info-table td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }}
        .info-table th {{
            background-color: #f2f2f2;
            font-weight: bold;
            width: 150px;
        }}
        .info-table tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        @media print {{
            body {{
                font-size: 12pt;
            }}
            .no-print {{
                display: none;
            }}
            .student-photo {{
                max-width: 300px;
            }}
        }}
        .print-btn {{
            background-color: #0066cc;
            color: white;
            border: none;
            padding: 10px 20px;
            font-size: 16px;
            cursor: pointer;
            margin: 20px;
            border-radius: 5px;
        }}
        .print-btn:hover {{
            background-color: #0052a3;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>学生学籍信息档案</h1>
    </div>
    
    {photo_html if photo_html else ''}
    
    {f'<div class="section"><h2>学生学籍详细信息</h2><table class="info-table">{info_rows}</table></div>' if info_rows else ''}
    
    <div class="footer">
        <p>--- 学生学籍信息档案结束 ---</p>
    </div>
    
    <div class="no-print" style="text-align: center;">
        <button class="print-btn" onclick="window.print()">打印此页面</button>
        <button class="print-btn" onclick="window.close()">关闭窗口</button>
    </div>
    
    <script>
        window.onload = function() {{
            console.log("页面加载完成，准备打印...");
        }};
    </script>
</body>
</html>'''
        
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return html_path
    
    def actual_print(self, print_text, print_photo, print_detailed_info, photo_size):
        """实际的打印操作"""
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp()
            
            # 创建HTML打印文件
            html_path = self.create_print_html(temp_dir, print_text, print_photo, print_detailed_info, photo_size)
            
            # 使用浏览器打开HTML文件
            webbrowser.open('file://' + os.path.abspath(html_path))
            
            # 显示成功信息
            self.status_label.config(text="打印文件已生成，请在浏览器中打印", foreground="green")
            
            messagebox.showinfo("打印准备就绪", 
                f"已生成打印文件，请在浏览器中:\\n\\n"
                f"1. 使用浏览器的打印功能 (Ctrl+P)\\n"
                f"2. 设置打印机和打印选项\\n"
                f"3. 点击打印按钮\\n\\n"
                f"文件已保存为: {html_path}")
            
            # 提示3秒后清理临时文件
            self.root.after(3000, lambda: self.cleanup_temp_files(temp_dir))
                
        except Exception as e:
            messagebox.showerror("打印错误", f"打印过程中出错: {str(e)}")
            self.status_label.config(text="打印失败", foreground="red")
            import traceback
            traceback.print_exc()
    
    def cleanup_temp_files(self, temp_dir):
        """清理临时文件"""
        try:
            shutil.rmtree(temp_dir)
            print(f"已清理临时目录: {temp_dir}")
        except Exception as e:
            print(f"清理临时文件时出错: {e}")
    
    def show_random_student(self):
        """随机展示一个学生的信息"""
        if not hasattr(self, 'table_students') or len(self.table_students) == 0:
            messagebox.showwarning("提示", "没有可用的学生数据")
            return
        
        random_idx = random.randint(0, len(self.table_students) - 1)
        random_student = self.table_students[random_idx]
        
        # 在 sorted_students 中找到对应的索引，确保跳转到主界面时索引正确
        sorted_idx = -1
        for i, s in enumerate(self.sorted_students):
            if s is random_student:
                sorted_idx = i
                break
        if sorted_idx == -1:
            sorted_idx = random_idx
        
        # 创建新窗口显示随机学生信息
        random_window = tk.Toplevel(self.root)
        random_window.title("随机")
        random_window.geometry("1500x1000")
        
        # 高DPI渲染
        try:
            random_window.tk.call('tk', 'scaling', 1.5)
        except:
            pass
        
        # 设置窗口居中
        screen_width = random_window.winfo_screenwidth()
        screen_height = random_window.winfo_screenheight()
        window_width = 1500
        window_height = 1000
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        random_window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        
        # 存储当前随机学生信息（使用局部变量避免污染主界面状态）
        self.random_window = random_window
        self.random_student = random_student
        self.random_idx = sorted_idx
        
        # 随机窗口多页照片状态
        random_photo_paths = []
        rp = random_student.get('照片路径')
        if rp and os.path.exists(rp):
            random_photo_paths.append(rp)
        for ep in random_student.get('extra_photos', []):
            if os.path.exists(ep):
                random_photo_paths.append(ep)
        random_photo_index = [0]  # 用列表实现闭包中的可变变量
        
        # 窗口内容框架
        main_frame = ttk.Frame(random_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题区域
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(title_frame, text="", 
                 font=('Microsoft YaHei', 16, 'bold'), 
                 foreground="blue").pack(side=tk.LEFT)
        
        # 随机信息标签
        random_info = f"随机选择: {random_student.get('姓名', '未知姓名')}"
        ttk.Label(title_frame, text=random_info, 
                 font=('Microsoft YaHei', 10),
                 foreground="green").pack(side=tk.RIGHT)
        
        # 内容区域
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧图片区域
        left_frame = ttk.LabelFrame(content_frame, text="学生照片", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 创建Canvas显示照片
        self.random_photo_canvas = tk.Canvas(left_frame, bg='white', 
                                            highlightthickness=1, highlightbackground="gray")
        self.random_photo_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 绑定双击事件查看大图
        self.random_photo_canvas.bind("<Double-Button-1>", 
                                     lambda e: self.show_large_image_for_student(random_student))
        
        # 照片说明标签
        self.random_photo_label = ttk.Label(left_frame, text="学生照片展示 (双击查看大图)", foreground="gray")
        self.random_photo_label.pack(pady=(5, 0))
        
        # 多页照片导航 - 随机窗口
        random_nav_frame = ttk.Frame(left_frame)
        random_nav_frame.pack(pady=(5, 0))
        random_prev_btn = ttk.Button(random_nav_frame, text="◀ 上一张", state="disabled")
        random_prev_btn.pack(side=tk.LEFT, padx=10)
        random_page_label = ttk.Label(random_nav_frame, text="", font=('Microsoft YaHei', 10))
        random_page_label.pack(side=tk.LEFT, padx=10)
        random_next_btn = ttk.Button(random_nav_frame, text="下一张 ▶", state="disabled")
        random_next_btn.pack(side=tk.LEFT, padx=10)
        
        def update_random_nav():
            t = len(random_photo_paths)
            if t <= 1:
                random_prev_btn.config(state="disabled")
                random_next_btn.config(state="disabled")
                random_page_label.config(text="")
            else:
                random_prev_btn.config(state="normal" if random_photo_index[0] > 0 else "disabled")
                random_next_btn.config(state="normal" if random_photo_index[0] < t - 1 else "disabled")
                random_page_label.config(text=f"{random_photo_index[0] + 1} / {t}")
        
        def random_photo_prev():
            if random_photo_index[0] > 0:
                random_photo_index[0] -= 1
                self.display_random_photo(random_student, random_photo_index[0], random_photo_paths, random_page_label)
                update_random_nav()
        
        def random_photo_next():
            if random_photo_index[0] < len(random_photo_paths) - 1:
                random_photo_index[0] += 1
                self.display_random_photo(random_student, random_photo_index[0], random_photo_paths, random_page_label)
                update_random_nav()
        
        random_prev_btn.config(command=random_photo_prev)
        random_next_btn.config(command=random_photo_next)
        
        # 右侧信息区域
        right_frame = ttk.LabelFrame(content_frame, text="学生详细信息", padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 创建带滚动条的文本区域
        text_frame = ttk.Frame(right_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        # 垂直滚动条
        v_scrollbar = ttk.Scrollbar(text_frame)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 信息显示文本框
        self.random_info_text = tk.Text(text_frame, wrap=tk.WORD, 
                                       yscrollcommand=v_scrollbar.set,
                                       font=('Microsoft YaHei', 11))
        self.random_info_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        v_scrollbar.config(command=self.random_info_text.yview)
        
        # 底部按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 定义刷新随机学生的函数
        def refresh_random():
            ridx = random.randint(0, len(self.table_students) - 1)
            self.random_student = self.table_students[ridx]
            # 在 sorted_students 中找到对应索引
            self.random_idx = next((i for i, s in enumerate(self.sorted_students) if s is self.random_student), ridx)
            # 重置多页照片
            random_photo_paths.clear()
            rp = self.random_student.get('照片路径')
            if rp and os.path.exists(rp):
                random_photo_paths.append(rp)
            for ep in self.random_student.get('extra_photos', []):
                if os.path.exists(ep):
                    random_photo_paths.append(ep)
            random_photo_index[0] = 0
            
            # 更新随机信息标签
            random_info = f"随机选择: {self.random_student.get('姓名', '未知姓名')}"
            title_frame.children['!label2'].config(text=random_info)
            
            # 显示照片
            self.display_random_photo(self.random_student, random_photo_index[0], random_photo_paths, random_page_label)
            update_random_nav()
            
            # 显示详细信息
            self.display_random_details(self.random_student)
        
        # 打印函数 - 使用独立方式不污染主界面
        def print_random_student():
            if not self.random_student:
                messagebox.showwarning("警告", "没有学生信息可打印")
                return
            
            saved = self.current_student
            self.current_student = self.random_student
            try:
                self.print_student_info()
            finally:
                self.current_student = saved
        
        # 按钮
        ttk.Button(button_frame, text="重新随机", command=refresh_random).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="在主窗口查看", 
                  command=lambda: self.show_in_main_window(self.random_idx)).pack(side=tk.LEFT, padx=5)
        
        # 随机窗口的打印按钮
        ttk.Button(button_frame, text="打印", command=print_random_student).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="关闭", command=random_window.destroy).pack(side=tk.LEFT, padx=5)
        
        # 初始化显示随机学生信息
        self.display_random_photo(self.random_student, random_photo_index[0], random_photo_paths, random_page_label)
        update_random_nav()
        self.display_random_details(self.random_student)
    
    def display_random_photo(self, student, photo_index=0, photo_paths=None, page_label=None):
        """在随机窗口中显示学生照片"""
        self.random_photo_canvas.delete("all")
        
        student_name = student.get('姓名', '未知姓名')
        
        if photo_paths is None:
            photo_paths = []
            rp = student.get('照片路径')
            if rp and os.path.exists(rp):
                photo_paths.append(rp)
            for ep in student.get('extra_photos', []):
                if os.path.exists(ep):
                    photo_paths.append(ep)
        
        if photo_index < len(photo_paths):
            photo_path = photo_paths[photo_index]
            try:
                img = Image.open(photo_path)
                original_width, original_height = img.size
                
                canvas_width = 400
                canvas_height = 500
                
                width_ratio = canvas_width / original_width
                height_ratio = canvas_height / original_height
                scale_ratio = min(width_ratio, height_ratio)
                
                new_width = int(original_width * scale_ratio)
                new_height = int(original_height * scale_ratio)
                
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                photo = ImageTk.PhotoImage(img)
                
                x_position = (canvas_width - new_width) // 2
                y_position = (canvas_height - new_height) // 2
                
                self.random_photo_canvas.create_image(x_position, y_position, anchor=tk.NW, image=photo)
                self.random_photo_canvas.image = photo
                
                total = len(photo_paths)
                page_info = f" ({photo_index + 1}/{total})" if total > 1 else ""
                self.random_photo_label.config(text=f"{student_name}{page_info} - {new_width}x{new_height}px (双击查看大图)", 
                                              foreground="green")
                
            except Exception as e:
                error_msg = f"无法加载照片\n{str(e)}"
                self.random_photo_canvas.create_text(200, 250, text=error_msg, 
                                                    fill="red", font=('Arial', 12))
                self.random_photo_label.config(text="照片加载失败 (双击查看大图)", foreground="red")
        else:
            self.random_photo_canvas.delete("all")
            self.random_photo_canvas.create_text(200, 250, text="无照片", 
                                                fill="gray", font=('Microsoft YaHei', 24))
            self.random_photo_label.config(text=f"{student_name} - 暂无照片", foreground="orange")
    
    def display_random_details(self, student):
        """在随机窗口中显示学生详细信息"""
        self.random_info_text.delete(1.0, tk.END)
        
        content = self.get_display_content(student)
        
        for line in content:
            if line.startswith("="):
                self.random_info_text.insert(tk.END, line + "\n", "bold")
            else:
                self.random_info_text.insert(tk.END, line + "\n")
        
        self.random_info_text.tag_config("bold", font=("Microsoft YaHei", 13, "bold"))
    
    def show_large_image_for_student(self, student):
        """为指定学生显示大图"""
        photo_paths = []
        rp = student.get('照片路径')
        if rp and os.path.exists(rp):
            photo_paths.append(rp)
        for ep in student.get('extra_photos', []):
            if os.path.exists(ep):
                photo_paths.append(ep)
        
        if not photo_paths:
            messagebox.showwarning("提示", "照片文件不存在")
            return
        
        # 默认显示第一张
        photo_path = photo_paths[0]
        
        try:
            # 创建新窗口显示大图
            large_window = tk.Toplevel(self.root)
            large_window.title("查看大图 - " + student.get('姓名', '未知姓名'))
            
            # 高DPI渲染
            try:
                large_window.tk.call('tk', 'scaling', 1.5)
            except:
                pass
            
            # 打开图片
            img = Image.open(photo_path)
            original_width, original_height = img.size
            
            # 设置窗口最大尺寸为屏幕的70%
            screen_width = large_window.winfo_screenwidth()
            screen_height = large_window.winfo_screenheight()
            max_width = int(screen_width * 0.7)
            max_height = int(screen_height * 0.7)
            
            # 计算缩放比例（保持宽高比）
            width_ratio = max_width / original_width
            height_ratio = max_height / original_height
            scale_ratio = min(width_ratio, height_ratio, 1.5)  # 限制最大缩放为150%
            
            # 计算新尺寸
            new_width = int(original_width * scale_ratio)
            new_height = int(original_height * scale_ratio)
            
            # 调整图片大小
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # 转换为Tkinter可用的格式
            photo = ImageTk.PhotoImage(img)
            
            # 设置窗口大小
            window_width = new_width + 20
            window_height = new_height + 40
            x_position = (screen_width - window_width) // 2
            y_position = (screen_height - window_height) // 2
            
            large_window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
            
            # 创建Canvas显示图片
            canvas = tk.Canvas(large_window, width=new_width, height=new_height, bg='white')
            canvas.pack(padx=10, pady=10)
            canvas.create_image(0, 0, anchor=tk.NW, image=photo)
            canvas.image = photo  # 保持引用
            
            # 添加信息标签
            info_label = ttk.Label(large_window, 
                                  text=f"原始尺寸: {original_width}x{original_height} | 显示尺寸: {new_width}x{new_height}",
                                  font=('Microsoft YaHei', 9))
            info_label.pack(pady=(0, 10))
            
            # 添加关闭按钮
            close_button = ttk.Button(large_window, text="关闭", command=large_window.destroy)
            close_button.pack(pady=(0, 10))
            
        except Exception as e:
            messagebox.showerror("错误", f"无法打开大图: {str(e)}")
    
    def show_large_image(self, event=None):
        """双击查看大图"""
        if not self.current_student:
            return
        
        if not self.current_photo_paths:
            messagebox.showwarning("提示", "照片文件不存在")
            return
        
        photo_path = self.current_photo_paths[self.current_photo_index]
        if not os.path.exists(photo_path):
            messagebox.showwarning("提示", "照片文件不存在")
            return
        
        try:
            # 创建新窗口显示大图
            large_window = tk.Toplevel(self.root)
            name = self.current_student.get('姓名', '未知姓名')
            total = len(self.current_photo_paths)
            page_info = f" ({self.current_photo_index + 1}/{total})" if total > 1 else ""
            large_window.title(f"查看大图 - {name}{page_info}")
            
            # 高DPI渲染
            try:
                large_window.tk.call('tk', 'scaling', 1.5)
            except:
                pass
            
            # 打开图片
            img = Image.open(photo_path)
            original_width, original_height = img.size
            
            # 设置窗口最大尺寸为屏幕的70%
            screen_width = large_window.winfo_screenwidth()
            screen_height = large_window.winfo_screenheight()
            max_width = int(screen_width * 0.7)
            max_height = int(screen_height * 0.7)
            
            # 计算缩放比例（保持宽高比）
            width_ratio = max_width / original_width
            height_ratio = max_height / original_height
            scale_ratio = min(width_ratio, height_ratio, 1.5)  # 限制最大缩放为150%
            
            # 计算新尺寸
            new_width = int(original_width * scale_ratio)
            new_height = int(original_height * scale_ratio)
            
            # 调整图片大小
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # 转换为Tkinter可用的格式
            photo = ImageTk.PhotoImage(img)
            
            # 设置窗口大小
            window_width = new_width + 20
            window_height = new_height + 40
            x_position = (screen_width - window_width) // 2
            y_position = (screen_height - window_height) // 2
            
            large_window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
            
            # 创建Canvas显示图片
            canvas = tk.Canvas(large_window, width=new_width, height=new_height, bg='white')
            canvas.pack(padx=10, pady=10)
            canvas.create_image(0, 0, anchor=tk.NW, image=photo)
            canvas.image = photo  # 保持引用
            
            # 添加信息标签
            info_label = ttk.Label(large_window, 
                                  text=f"原始尺寸: {original_width}x{original_height} | 显示尺寸: {new_width}x{new_height}",
                                  font=('Microsoft YaHei', 9))
            info_label.pack(pady=(0, 10))
            
            # 添加关闭按钮
            close_button = ttk.Button(large_window, text="关闭", command=large_window.destroy)
            close_button.pack(pady=(0, 10))
            
        except Exception as e:
            messagebox.showerror("错误", f"无法打开大图: {str(e)}")
    
    def show_in_main_window(self, student_idx):
        """在主窗口中显示指定学生"""
        if student_idx < 0 or student_idx >= len(self.sorted_students):
            return
        
        if hasattr(self, 'random_window'):
            self.random_window.destroy()
        
        self.tree.selection_set(str(student_idx))
        self.tree.see(str(student_idx))
        
        student = self.sorted_students[student_idx]
        self.display_student_info_by_student(student)
    
    def open_data_folder(self):
        """打开数据文件夹"""
        try:
            if os.path.exists(self.base_dir):
                os.startfile(self.base_dir)
            else:
                messagebox.showinfo("提示", f"文件夹不存在: {self.base_dir}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {str(e)}")
    
    def clear_search(self):
        """清空搜索"""
        self.search_var.set("")
        self.show_all_students()
    
    def search_student(self):
        """搜索学生"""
        keyword = self.search_var.get().strip()
        search_type = self.search_type.get()
        
        if not keyword:
            self.show_all_students()
            return
        
        # 更新状态
        self.status_label.config(text=f"搜索中: {keyword}", foreground="blue")
        
        # 清空树形视图
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 搜索匹配的学生
        matched_students = []
        for idx, student in enumerate(self.sorted_students):
            search_field = student.get(search_type, '')
            if search_field and keyword.lower() in search_field.lower():
                matched_students.append((idx, student))
        
        # 显示搜索结果
        if matched_students:
            for idx, student in matched_students:
                has_photo = "有" if student.get('has_photo') else "无"
                row_tag = 'even' if idx % 2 == 0 else 'odd'
                self.tree.insert('', tk.END, iid=idx, tags=(row_tag,), values=(
                    idx + 1,
                    student.get('姓名', ''),
                    student.get('学籍号', ''),
                    student.get('身份证号', ''),
                    has_photo
                ))
            
            # 默认选择第一个结果并立即显示
            if self.tree.get_children():
                self.tree.selection_set(self.tree.get_children()[0])
                self.on_tree_select(None)
            
            self.status_label.config(text=f"找到{len(matched_students)}个结果", foreground="green")
        else:
            self.status_label.config(text="未找到匹配结果", foreground="red")
            messagebox.showinfo("提示", f"未找到匹配'{keyword}'的学生")
    
    def natural_sort_key(self, value):
        """Windows资源管理器风格的自然排序键（数字按数值比较）"""
        if value is None:
            return []
        s = str(value).strip()
        parts = re.split(r'(\d+)', s)
        return [int(part) if part.isdigit() else part.lower() for part in parts]
    
    def apply_current_sort(self):
        """按当前排序状态对 table_students 排序（未设置时默认按姓名升序）"""
        if self.sort_column and self.sort_column != '序号':
            field = self.sort_column
            self.sorted_students = sorted(
                self.table_students,
                key=lambda x: self.natural_sort_key(x.get(field, '')),
                reverse=self.sort_reverse
            )
        else:
            self.sorted_students = sorted(
                self.table_students,
                key=lambda x: self.natural_sort_key(x.get('姓名', ''))
            )
    
    def sort_treeview(self, col):
        """点击表头排序（Windows资源管理器风格：▲升序 / ▼降序）"""
        if col == '序号':
            # 序号列恢复默认（原始）顺序
            self.sort_column = None
            self.sort_reverse = False
            self.sorted_students = list(self.table_students)
        else:
            if self.sort_column == col:
                self.sort_reverse = not self.sort_reverse
            else:
                self.sort_column = col
                self.sort_reverse = False
            self.apply_current_sort()
        
        self.populate_tree(self.sorted_students)
        self.update_sort_indicators()
        self.clear_display()
    
    def populate_tree(self, students):
        """将学生列表填充到表格中"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for idx, student in enumerate(students):
            has_photo = "有" if student.get('has_photo') else "无"
            row_tag = 'even' if idx % 2 == 0 else 'odd'
            self.tree.insert('', tk.END, iid=idx, tags=(row_tag,), values=(
                idx + 1,
                student.get('姓名', ''),
                student.get('学籍号', ''),
                student.get('身份证号', ''),
                has_photo
            ))
    
    def update_sort_indicators(self):
        """更新表头排序指示器（Windows资源管理器风格箭头）"""
        base_texts = {'序号': '序号', '姓名': '姓名', '学籍号': '学籍号', '身份证号': '身份证号'}
        arrow = ' ▼' if self.sort_reverse else ' ▲'
        for col in base_texts:
            if self.sort_column == col:
                self.tree.heading(col, text=base_texts[col] + arrow)
            else:
                self.tree.heading(col, text=base_texts[col])
    
    def show_all_students(self):
        """显示所有学生（按当前排序状态）"""
        if not hasattr(self, 'table_students'):
            self.table_students = []
        
        self.apply_current_sort()
        self.populate_tree(self.sorted_students)
        self.update_sort_indicators()
        self.clear_display()
    
    def clear_display(self):
        """清空显示区域"""
        self.photo_canvas.delete("all")
        self.photo_canvas.create_text(int(self.photo_canvas['width']) // 2,
                                      int(self.photo_canvas['height']) // 2,
                                      text="请选择学生查看照片", fill="gray", font=('Arial', 14))
        self.photo_label.config(text="点击学生查看照片 (双击图片查看大图)", foreground="gray")
        self.info_text.delete(1.0, tk.END)
        self.current_student = None
    
    def on_tree_select(self, event):
        """当选择树形视图中的项目时"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item_iid = selection[0]
        try:
            student_idx = int(item_iid)
        except (ValueError, TypeError):
            return
        
        if student_idx < 0 or student_idx >= len(self.sorted_students):
            return
        
        student = self.sorted_students[student_idx]
        name = student.get('姓名', '')
        self.status_label.config(text=f"加载中: {name}", foreground="blue")
        self.display_student_info_by_student(student)
    
    def display_student_info(self, student_idx):
        """通过索引显示学生信息"""
        if student_idx < 0 or student_idx >= len(self.table_students):
            return
        
        self.current_student = self.table_students[student_idx]
        self._display_current_student_info()
    
    def display_student_info_by_student(self, student):
        """通过学生对象显示学生信息"""
        self.current_student = student
        self._display_current_student_info()
    
    def _display_current_student_info(self):
        """显示当前选中学生的信息（并切换到详情页）"""
        # 切换到详情页
        if hasattr(self, 'page_detail'):
            student_name = self.current_student.get('姓名', '')
            self.detail_title.config(text=f"学生详情 - {student_name}")
            self.page_detail.tkraise()
        
        name = self.current_student.get('姓名', '')
        student_id = self.current_student.get('学籍号', '')
        id_card = self.current_student.get('身份证号', '')
        photo_path = self.current_student.get('照片路径')
        
        print(f"\n显示学生信息:")
        print(f"姓名: {name}")
        print(f"学籍号: {student_id}")
        print(f"身份证号: {id_card}")
        print(f"照片路径: {photo_path}")
        
        # 构建多页照片列表
        self.current_photo_paths = []
        if photo_path and os.path.exists(photo_path):
            self.current_photo_paths.append(photo_path)
        extra = self.current_student.get('extra_photos', [])
        for ep in extra:
            if os.path.exists(ep):
                self.current_photo_paths.append(ep)
        self.current_photo_index = 0
        
        if self.current_photo_paths:
            self.display_adaptive_photo(self.current_photo_paths[0], name)
        else:
            self.display_no_photo(name)
        self.display_student_details(name, student_id, id_card)
        self.update_photo_nav_buttons()
        
        self.status_label.config(text=f"已显示: {name}", foreground="green")
    
    def display_adaptive_photo(self, photo_path, student_name):
        """自适应显示照片，保持完整显示"""
        self.photo_canvas.delete("all")
        
        if photo_path and os.path.exists(photo_path):
            try:
                img = Image.open(photo_path)
                original_width, original_height = img.size
                
                canvas_width = int(self.photo_canvas['width'])
                canvas_height = int(self.photo_canvas['height'])
                
                width_ratio = canvas_width / original_width
                height_ratio = canvas_height / original_height
                scale_ratio = min(width_ratio, height_ratio)
                
                new_width = int(original_width * scale_ratio)
                new_height = int(original_height * scale_ratio)
                
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                photo = ImageTk.PhotoImage(img)
                
                x_position = (canvas_width - new_width) // 2
                y_position = (canvas_height - new_height) // 2
                
                self.photo_canvas.create_image(x_position, y_position, anchor=tk.NW, image=photo)
                self.photo_canvas.image = photo
                
                total = len(self.current_photo_paths)
                page_info = f" ({self.current_photo_index + 1}/{total})" if total > 1 else ""
                self.photo_label.config(text=f"{student_name}{page_info} - {new_width}x{new_height}px (双击查看大图)", 
                                       foreground="green")
                
            except Exception as e:
                error_msg = f"无法加载照片\n{str(e)}"
                self.photo_canvas.create_text(int(self.photo_canvas['width']) // 2,
                                              int(self.photo_canvas['height']) // 2,
                                              text=error_msg, fill="red", font=('Arial', 12))
                self.photo_label.config(text="照片加载失败 (双击查看大图)", foreground="red")
        else:
            self.display_no_photo(student_name)
    
    def display_no_photo(self, student_name):
        """显示无照片提示"""
        self.photo_canvas.delete("all")
        self.photo_canvas.create_text(int(self.photo_canvas['width']) // 2,
                                      int(self.photo_canvas['height']) // 2,
                                      text="无照片", fill="gray", font=('Microsoft YaHei', 24))
        self.photo_label.config(text=f"{student_name} - 暂无照片", foreground="orange")
    
    def photo_prev(self):
        if self.current_photo_index > 0 and self.current_photo_paths:
            self.current_photo_index -= 1
            name = self.current_student.get('姓名', '')
            self.display_adaptive_photo(self.current_photo_paths[self.current_photo_index], name)
            self.update_photo_nav_buttons()
    
    def photo_next(self):
        if self.current_photo_index < len(self.current_photo_paths) - 1 and self.current_photo_paths:
            self.current_photo_index += 1
            name = self.current_student.get('姓名', '')
            self.display_adaptive_photo(self.current_photo_paths[self.current_photo_index], name)
            self.update_photo_nav_buttons()
    
    def update_photo_nav_buttons(self):
        total = len(self.current_photo_paths)
        if total <= 1:
            self.photo_prev_btn.config(state="disabled")
            self.photo_next_btn.config(state="disabled")
            self.photo_page_label.config(text="")
        else:
            self.photo_prev_btn.config(state="normal" if self.current_photo_index > 0 else "disabled")
            self.photo_next_btn.config(state="normal" if self.current_photo_index < total - 1 else "disabled")
            self.photo_page_label.config(text=f"{self.current_photo_index + 1} / {total}")
        
        # 更新照片数量提示
        if hasattr(self, 'photo_count_label'):
            self.photo_count_label.config(text=f"共 {total} 张 / 上限 {self.MAX_PHOTOS} 张")
    
    def count_student_photos(self, student=None):
        """统计某个学生当前的实际照片数量"""
        student = student or self.current_student
        if not student:
            return 0
        count = 0
        main = student.get('照片路径')
        if main and os.path.exists(main):
            count += 1
        for ep in student.get('extra_photos', []):
            if os.path.exists(ep):
                count += 1
        return count
    
    def add_photo(self):
        """为当前学生增加照片（最多20张）"""
        if not self.current_student:
            messagebox.showwarning("提示", "请先选择一个学生")
            return
        
        current_count = self.count_student_photos()
        if current_count >= self.MAX_PHOTOS:
            messagebox.showwarning("提示", f"该学生已有 {current_count} 张照片，已达上限（最多 {self.MAX_PHOTOS} 张）")
            return
        
        file_paths = filedialog.askopenfilenames(
            title="选择要增加的照片（可多选）",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif"), ("所有文件", "*.*")]
        )
        if not file_paths:
            return
        
        added = 0
        skipped = 0
        index = current_count
        for fp in file_paths:
            if index >= self.MAX_PHOTOS:
                skipped += 1
                continue
            try:
                if self.copy_photo_to_student(fp, index):
                    index += 1
                    added += 1
                else:
                    skipped += 1
            except Exception as e:
                print(f"增加照片失败: {e}")
                skipped += 1
        
        # 重新加载数据并重新选中当前学生
        self.reload_and_reselect()
        
        total = self.count_student_photos()
        if added > 0:
            messagebox.showinfo("增加成功", f"成功增加 {added} 张照片\n当前共 {total} 张照片（上限 {self.MAX_PHOTOS} 张）")
        if skipped > 0:
            messagebox.showwarning("提示", f"有 {skipped} 张照片未添加（已达上限或格式不支持）")
    
    def copy_photo_to_student(self, file_path, index):
        """将图片复制到pic目录并命名为当前学生的格式"""
        if not os.path.exists(self.pic_folder):
            os.makedirs(self.pic_folder, exist_ok=True)
        
        student = self.current_student
        name = student.get('姓名', '')
        student_id = student.get('学籍号', '')
        id_card = student.get('身份证号', '')
        
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in ('.jpg', '.jpeg', '.png', '.bmp', '.gif'):
            return False
        
        # 生成基础文件名（与照片解析格式一致）
        base = ""
        if name and name != "未知姓名":
            base += f"姓名：{name} "
        if student_id and student_id != "未知学籍号":
            base += f"学籍号：{student_id} "
        if id_card and id_card != "未知身份证号":
            base += f"身份证号：{id_card}"
        base = base.strip()
        if not base:
            base = "未知姓名"
        
        # 生成新文件名：第一张无后缀，后续加 _数字 后缀
        new_filename = f"{base}{ext}" if index == 0 else f"{base}_{index}{ext}"
        dest = os.path.join(self.pic_folder, new_filename)
        
        counter = 1
        while os.path.exists(dest):
            new_filename = f"{base}_{index}_{counter}{ext}"
            dest = os.path.join(self.pic_folder, new_filename)
            counter += 1
        
        shutil.copy2(file_path, dest)
        print(f"已增加照片: {file_path} -> {dest}")
        return True
    
    def delete_photo(self):
        """删除当前显示的照片"""
        if not self.current_student:
            messagebox.showwarning("提示", "请先选择一个学生")
            return
        
        if not self.current_photo_paths:
            messagebox.showinfo("提示", "该学生当前没有照片，无需删除")
            return
        
        photo_path = self.current_photo_paths[self.current_photo_index]
        if not os.path.exists(photo_path):
            messagebox.showwarning("提示", "照片文件不存在")
            return
        
        student_name = self.current_student.get('姓名', '未知姓名')
        confirm = messagebox.askyesno("确认删除",
            f"确定要删除当前照片吗？\n\n"
            f"学生: {student_name}\n"
            f"文件: {os.path.basename(photo_path)}\n\n"
            f"删除后文件将不可恢复！")
        if not confirm:
            return
        
        try:
            os.remove(photo_path)
            print(f"已删除照片: {photo_path}")

            
        except Exception as e:
            messagebox.showerror("删除失败", f"删除照片时出错:\n{str(e)}")
            return
        
        # 重新加载数据并重新选中当前学生
        self.reload_and_reselect()
        
        remaining = self.count_student_photos()
        messagebox.showinfo("删除成功", f"照片已删除\n该学生当前剩余 {remaining} 张照片")
    
    def reload_and_reselect(self):
        """重新加载数据并重新选中当前学生"""
        if not self.current_student:
            self.reload_data()
            return
        
        name = self.current_student.get('姓名', '')
        student_id = self.current_student.get('学籍号', '')
        id_card = self.current_student.get('身份证号', '')
        
        self.reload_data()
        
        # 重新查找该学生并选中显示
        for idx, student in enumerate(self.sorted_students):
            if (student.get('姓名') == name and
                (not student_id or student.get('学籍号') == student_id) and
                (not id_card or student.get('身份证号') == id_card)):
                self.tree.selection_set(str(idx))
                self.tree.see(str(idx))
                self.display_student_info_by_student(student)
                return
        
        self.clear_display()
    
    def find_student_in_dataframe(self, df, name, student_id, id_card):
        """在单个DataFrame中查找学生信息"""
        if df.empty:
            return None
        
        # 策略1: 精确匹配身份证号（最可靠）
        if id_card and id_card != "未知身份证号":
            for col in df.columns:
                # 检查列名是否包含身份证相关关键词
                if any(keyword in str(col).lower() for keyword in ['身份证', '证件', 'idcard', 'id']):
                    matches = df[df[col] == id_card]
                    if not matches.empty:
                        print(f"通过身份证号在列'{col}'中找到匹配: {id_card}")
                        return matches.iloc[0]
        
        # 策略2: 精确匹配学籍号
        if student_id and student_id != "未知学籍号":
            for col in df.columns:
                # 检查列名是否包含学籍号相关关键词
                if any(keyword in str(col).lower() for keyword in ['学籍', '学号', 'student', '学籍号']):
                    matches = df[df[col] == student_id]
                    if not matches.empty:
                        print(f"通过学籍号在列'{col}'中找到匹配: {student_id}")
                        return matches.iloc[0]
        
        # 策略3: 精确匹配姓名（需要确保姓名准确）
        if name and name != "未知姓名":
            for col in df.columns:
                # 检查列名是否包含姓名相关关键词
                if any(keyword in str(col).lower() for keyword in ['姓名', '名字', 'name', 'student']):
                    matches = df[df[col] == name]
                    if not matches.empty:
                        print(f"通过姓名在列'{col}'中找到精确匹配: {name}")
                        return matches.iloc[0]
            
            # 策略4: 模糊匹配姓名（包含关系）
            for col in df.columns:
                if any(keyword in str(col).lower() for keyword in ['姓名', '名字', 'name']):
                    # 查找包含学生姓名的记录
                    for idx, row in df.iterrows():
                        cell_value = row[col]
                        if name in str(cell_value):
                            print(f"通过姓名在列'{col}'中找到模糊匹配: {name} in {cell_value}")
                            return row
        
        # 策略5: 尝试其他列匹配
        if name and name != "未知姓名":
            for col in df.columns:
                # 跳过已经检查过的列类型
                col_lower = str(col).lower()
                if any(keyword in col_lower for keyword in ['姓名', '名字', 'name', '学籍', '学号', '身份证', '证件', 'id']):
                    continue
                
                # 在其他列中查找姓名
                for idx, row in df.iterrows():
                    cell_value = str(row[col])
                    if name in cell_value:
                        print(f"在其他列'{col}'中找到匹配: {name} in {cell_value}")
                        return row
        
        print(f"在数据表中未找到匹配: 姓名={name}, 学籍号={student_id}, 身份证号={id_card}")
        return None
    
    def format_cell_value(self, col_name, value):
        """格式化单元格值，避免科学计数法和.0问题"""
        if value is None or str(value).strip() in ['', 'nan', 'None']:
            return None
        
        col_lower = str(col_name).lower()
        value_str = str(value).strip()
        
        # 检测是否为身份证号、学籍号、手机号等需要保持原格式的字段
        is_id_field = any(kw in col_lower for kw in ['身份证', '学籍号', '学号', '手机', '电话', 'idcard', 'id_no', 'student_id'])
        
        if is_id_field and value_str:
            # 去除末尾的.0
            if value_str.endswith('.0'):
                value_str = value_str[:-2]
            
            # 处理科学计数法 (如 1.23E+17 或 1.23e+10)
            if 'e' in value_str.lower() and '+' in value_str:
                try:
                    # 分离科学计数法的各部分
                    parts = value_str.lower().split('e')
                    mantissa = parts[0]  # 尾数部分
                    exp = int(parts[1])  # 指数部分
                    
                    # 移除小数点
                    mantissa = mantissa.replace('.', '')
                    
                    # 根据指数移动小数点位置
                    if exp >= 0:
                        result = mantissa + '0' * exp
                    else:
                        exp = abs(exp)
                        result = '0' * exp + mantissa
                        result = result[:exp] + '.' + result[exp:]
                    
                    # 去除末尾的.0
                    if result.endswith('.0'):
                        result = result[:-2]
                    value_str = result
                except:
                    pass
            
            return value_str
        
        return value_str

    def get_student_details(self, name, student_id, id_card):
        """获取学生的详细信息"""
        details = {}
        
        print(f"\n开始搜索学生信息:")
        print(f"搜索条件 - 姓名: {name}, 学籍号: {student_id}, 身份证号: {id_card}")
        
        # 只在一个表格中查找信息
        if not self.df1.empty:
            print(f"\n在表中搜索...")
            row = self.find_student_in_dataframe(self.df1, name, student_id, id_card)
            
            if row is not None:
                print(f"在表中找到匹配记录")
                for col in self.df1.columns:
                    cell_value = self.format_cell_value(col, row[col])
                    if cell_value and cell_value != 'nan' and cell_value != 'None':
                        details[col] = cell_value
                        print(f"  添加字段: {col} = {cell_value}")
            else:
                print(f"在表中未找到匹配记录")
        
        print(f"共找到 {len(details)} 个字段")
        return details
    
    def get_field_value(self, student, keywords):
        """从学生信息中查找匹配关键词的字段值"""
        for key in student.keys():
            if key in ['照片路径', 'has_photo', 'table', 'row_index']:
                continue
            key_lower = str(key).lower()
            for kw in keywords:
                if kw in key_lower:
                    value = student.get(key, '')
                    if value and str(value).strip() not in ['', 'nan', 'None']:
                        return str(value).strip()
        return ''
    
    def display_student_details(self, name, student_id, id_card):
        """在文本框中显示学生详细信息（统一格式）"""
        self.info_text.delete(1.0, tk.END)
        
        content = []
        content.append("=" * 60)
        content.append("学生学籍详细信息".center(60))
        content.append("=" * 60)
        content.append("")
        
        if self.current_student:
            # 定义显示顺序
            display_order = [
                ('姓名', ['姓名', '名字']),
                ('性别', ['性别', '男', '女']),
                ('民族', ['民族']),
                ('出生日期', ['出生日期', '出生', '生日', '出生年月']),
                ('身份证号', ['身份证号', '身份证', 'ID']),
                ('全国学籍号', None),
                ('省学籍辅号（现学籍号）', ['省学籍辅号', '现学籍号', '学籍号', '学号']),
                ('家庭住址', ['家庭住址', '家庭居住地', '住址', '地址', '现住址', '家庭地址', '住址所在地']),
                ('父亲姓名', ['父亲姓名', '父亲名字', '父亲']),
                ('父亲联系电话', ['父亲联系电话', '父亲电话', '父亲手机', '父亲手机号']),
                ('父亲身份证号', ['父亲身份证号', '父亲身份证']),
                ('父亲工作', ['父亲工作', '父亲职业', '父亲单位']),
                ('母亲姓名', ['母亲姓名', '母亲名字', '母亲']),
                ('母亲联系电话', ['母亲联系电话', '母亲电话', '母亲手机', '母亲手机号']),
                ('母亲身份证号', ['母亲身份证号', '母亲身份证']),
                ('母亲工作', ['母亲工作', '母亲职业', '母亲单位']),
            ]
            
            # 获取身份证号用于生成全国学籍号
            id_card_value = id_card if id_card else self.get_field_value(self.current_student, ['身份证'])
            
            # 按顺序显示信息
            for label, keywords in display_order:
                if label == '全国学籍号':
                    value = f"G{id_card_value}" if id_card_value else ''
                elif label == '身份证号':
                    value = id_card_value
                elif label == '姓名':
                    value = name
                elif label == '省学籍辅号':
                    value = student_id if student_id else self.get_field_value(self.current_student, ['学籍号', '学号', '省学籍'])
                else:
                    value = self.get_field_value(self.current_student, keywords)
                
                if value and str(value).strip():
                    content.append(f"{label}：{str(value).strip()}")
        
        content.append("")
        content.append("=" * 60)
        
        for line in content:
            if line.startswith("="):
                self.info_text.insert(tk.END, line + "\n", "bold")
            else:
                self.info_text.insert(tk.END, line + "\n")
        
        self.info_text.tag_config("bold", font=("Microsoft YaHei", 13, "bold"))
    
    def get_display_content(self, student):
        """获取统一格式的学生信息内容（用于导出和打印）"""
        content = []
        content.append("=" * 60)
        content.append("学生学籍详细信息".center(60))
        content.append("=" * 60)
        content.append("")
        
        name = student.get('姓名', '')
        student_id = student.get('学籍号', '')
        id_card = student.get('身份证号', '')
        
        display_order = [
            ('姓名', ['姓名', '名字', 'name']),
            ('性别', ['性别', '男', '女']),
            ('民族', ['民族']),
            ('出生日期', ['出生日期', '出生', '生日']),
            ('身份证号', ['身份证号', '身份证', 'idcard']),
            ('全国学籍号', None),
            ('省学籍辅号', ['学籍号', '学号', '省学籍']),
            ('联系电话', ['联系电话', '电话', '手机', '手机号']),
            ('现家庭住址', ['家庭住址', '家庭居住地', '住址', '地址', '现住址', '家庭地址']),
            ('父亲姓名', ['父亲姓名', '父亲名字']),
            ('父亲联系电话', ['父亲电话', '父亲手机']),
            ('父亲身份证号', ['父亲身份证']),
            ('父亲职业', ['父亲职业', '父亲工作']),
            ('母亲姓名', ['母亲姓名', '母亲名字']),
            ('母亲联系电话', ['母亲电话', '母亲手机']),
            ('母亲身份证号', ['母亲身份证']),
            ('母亲职业', ['母亲职业', '母亲工作']),
        ]
        
        id_card_value = id_card if id_card else self.get_field_value(student, ['身份证'])
        national_student_id = f"G{id_card_value}" if id_card_value else ''
        
        for label, keywords in display_order:
            if label == '全国学籍号':
                value = national_student_id
            else:
                value = self.get_field_value(student, keywords)
            
            if value:
                content.append(f"{label}：{value}")
        
        content.append("")
        content.append("=" * 60)
        
        return content
    
    def manual_correct_info(self):
        """手动修正学生信息"""
        if not self.current_student:
            messagebox.showwarning("警告", "请先选择一个学生")
            return
        
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("手动修正学生信息")
        dialog.geometry("700x560")
        dialog.transient(self.root)
        
        # 高DPI渲染
        try:
            dialog.tk.call('tk', 'scaling', 1.5)
        except:
            pass
        
        # 获取当前信息
        current_name = self.current_student.get('姓名', '')
        current_student_id = self.current_student.get('学籍号', '')
        current_id_card = self.current_student.get('身份证号', '')
        has_photo = self.current_student.get('has_photo', False)
        
        # 获取在sorted_students中的索引
        try:
            student_idx = self.sorted_students.index(self.current_student)
        except ValueError:
            student_idx = -1
        
        # 创建界面
        ttk.Label(dialog, text=f"照片状态: {'有照片' if has_photo else '无照片'}", font=('Microsoft YaHei', 10, 'bold')).pack(pady=(20, 5))
        
        # 创建输入框
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="姓    名:").grid(row=0, column=0, sticky=tk.W, pady=5)
        name_var = tk.StringVar(value=current_name)
        name_entry = ttk.Entry(frame, textvariable=name_var, width=30)
        name_entry.grid(row=0, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        
        ttk.Label(frame, text="学 籍 号:").grid(row=1, column=0, sticky=tk.W, pady=5)
        student_id_var = tk.StringVar(value=current_student_id)
        student_id_entry = ttk.Entry(frame, textvariable=student_id_var, width=30)
        student_id_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        
        ttk.Label(frame, text="身份证号:").grid(row=2, column=0, sticky=tk.W, pady=5)
        id_card_var = tk.StringVar(value=current_id_card)
        id_card_entry = ttk.Entry(frame, textvariable=id_card_var, width=30)
        id_card_entry.grid(row=2, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        
        def save_correction():
            new_name = name_var.get().strip()
            new_student_id = student_id_var.get().strip()
            new_id_card = id_card_var.get().strip()
            
            if not new_name:
                messagebox.showwarning("警告", "姓名不能为空")
                return
            
            # 更新学生信息
            self.current_student['姓名'] = new_name
            self.current_student['学籍号'] = new_student_id
            self.current_student['身份证号'] = new_id_card
            
            # 重新查找照片
            photo_path, extra_photos = self.find_photo_for_student(self.current_student)
            self.current_student['照片路径'] = photo_path
            self.current_student['extra_photos'] = extra_photos
            self.current_student['has_photo'] = photo_path is not None
            
            # 重新显示信息
            self._display_current_student_info()
            
            dialog.destroy()
            messagebox.showinfo("成功", "学生信息已更新")
        
        def cancel():
            dialog.destroy()
        
        # 按钮
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="保存", command=save_correction).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=cancel).pack(side=tk.LEFT, padx=10)
    
    def reload_data(self):
        """重新加载数据"""
        self.status_label.config(text="重新加载数据中...", foreground="blue")
        self.load_data()
        self.show_all_students()
        self.status_label.config(text="数据重新加载完成", foreground="green")
    
    def import_xlsx(self):
        """导入Excel表格并合并到现有数据中（按学籍号/身份证号/姓名去重）"""
        file_path = filedialog.askopenfilename(
            title="选择要导入的Excel表格",
            filetypes=[("Excel 表格", "*.xlsx"), ("Excel 97-2003 表格", "*.xls"), ("所有文件", "*.*")]
        )
        if not file_path:
            return
        
        try:
            self.status_label.config(text="正在读取表格...", foreground="blue")
            self.root.update()
            
            # 读取导入的表格（身份证号等列保持为字符串）
            new_df = pd.read_excel(file_path)
            str_cols = {}
            for col in new_df.columns:
                col_lower = str(col).lower()
                if any(kw in col_lower for kw in ['身份证', '学籍号', '学号', '手机', '电话']):
                    str_cols[col] = str
            if str_cols:
                new_df = pd.read_excel(file_path, dtype=str_cols)
            
            if new_df.empty:
                messagebox.showwarning("警告", "导入的表格为空，没有可导入的数据")
                return
            
            print(f"导入表格: {file_path}, 形状: {new_df.shape}")
            
            # 合并去重
            if not hasattr(self, 'df1') or self.df1 is None:
                self.df1 = pd.DataFrame()
            
            before = len(self.df1)
            merged_df, added, updated = self.merge_dataframes(self.df1, new_df)
            self.df1 = merged_df
            after = len(self.df1)
            
            # 重新构建学生列表并刷新界面
            self.build_student_list()
            self.show_all_students()
            
            self.status_label.config(
                text=f"导入成功: 新增{added}人, 更新{updated}人, 现有共{after}人",
                foreground="green")
            
            messagebox.showinfo("导入成功",
                f"表格导入成功\n\n"
                f"导入前学生数: {before}\n"
                f"新增学生: {added}\n"
                f"更新学生: {updated}\n"
                f"导入后学生总数: {after}")
            
        except Exception as e:
            messagebox.showerror("导入失败", f"导入表格时出错:\n{str(e)}")
            self.status_label.config(text="导入失败", foreground="red")
            import traceback
            traceback.print_exc()
    
    def merge_dataframes(self, df1, df2):
        """合并两个DataFrame，按身份证号/学籍号/姓名去重"""
        if df1 is None or df1.empty:
            return df2.copy(), len(df2), 0
        if df2 is None or df2.empty:
            return df1.copy(), 0, 0
        
        # 合并所有列，保证两表字段不丢失
        all_cols = []
        seen_cols = set()
        for col in list(df1.columns) + list(df2.columns):
            if col not in seen_cols:
                seen_cols.add(col)
                all_cols.append(col)
        
        old_records = df1.to_dict('records')
        new_records = df2.to_dict('records')
        
        # 合并结果：先放入旧表所有记录
        merged = []
        index_of = {}
        for idx, rec in enumerate(old_records):
            merged.append(rec)
            index_of[id(rec)] = idx
        
        # 建立旧表键索引
        old_index = {}
        for rec in old_records:
            for k in self.get_student_keys(rec):
                old_index.setdefault(k, rec)
        
        added = 0
        updated = 0
        for rec in new_records:
            # 按优先级查找匹配的旧记录
            target = None
            for k in self.get_student_keys(rec):
                if k in old_index:
                    target = old_index[k]
                    break
            
            if target is None:
                merged.append(rec)
                added += 1
            else:
                i = index_of[id(target)]
                merged[i] = self.merge_two_records(merged[i], rec)
                index_of[id(merged[i])] = i
                updated += 1
        
        result = pd.DataFrame(merged, columns=all_cols)
        print(f"合并去重完成: 新增{added}条, 更新{updated}条, 总计{len(result)}条")
        return result, added, updated
    
    def merge_two_records(self, old_rec, new_rec):
        """合并两条学生记录：新记录的非空字段覆盖旧记录，保留旧记录独有字段"""
        merged = dict(old_rec)
        for col, value in new_rec.items():
            if self.clean_record_value(value):
                merged[col] = value
        return merged
    
    def get_student_keys(self, record):
        """获取学生记录的所有可用标识键列表（身份证号 > 学籍号 > 姓名）"""
        keys = []
        col = self.find_column(record, ['身份证', 'idcard'])
        if col is not None:
            value = self.clean_record_value(record[col])
            if value:
                keys.append(('id', value))
        
        col = self.find_column(record, ['学籍号', '学号'])
        if col is not None:
            value = self.clean_record_value(record[col])
            if value:
                keys.append(('sid', value))
        
        col = self.find_column(record, ['姓名', '名字', 'name'])
        if col is not None:
            value = self.clean_record_value(record[col])
            if value:
                keys.append(('name', value))
        
        return keys
    
    def find_column(self, record, keywords):
        """在记录中查找包含关键词的列名，优先完全匹配，排除家长相关干扰列"""
        exclude = ['父亲', '母亲', '家长', '监护人']
        best_col = None
        best_score = -1
        for col in record.keys():
            col_str = str(col).strip()
            col_lower = col_str.lower()
            if any(ex in col_str for ex in exclude):
                continue
            score = 0
            for kw in keywords:
                if col_lower == kw.lower():
                    score = 3
                elif kw.lower() in col_lower:
                    score = max(score, 2)
            if score > best_score:
                best_score = score
                best_col = col
        return best_col
    
    def clean_record_value(self, value):
        """清理单元格值，去除空值和nan"""
        if value is None:
            return ''
        s = str(value).strip()
        if s in ('', 'nan', 'None'):
            return ''
        return s
    
    def export_current_zip(self):
        """导出当前学生信息为ZIP包"""
        if not self.current_student:
            messagebox.showwarning("警告", "请先选择一个学生")
            return
        
        name = self.current_student.get('姓名', '未知姓名')
        photo_path = self.current_student.get('照片路径')
        extra_photos = self.current_student.get('extra_photos', [])
        
        # 创建临时目录
        temp_dir = tempfile.mkdtemp()
        
        try:
            # 1. 创建文本文件
            safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
            txt_filename = f"{safe_name}_学籍信息.txt"
            txt_path = os.path.join(temp_dir, txt_filename)
            
            with open(txt_path, 'w', encoding='utf-8') as f:
                content = self.get_display_content(self.current_student)
                f.write('\n'.join(content))
            
            # 2. 复制所有照片文件
            photo_copied = False
            copied_photos = []
            
            all_photos = []
            if photo_path and os.path.exists(photo_path):
                all_photos.append(photo_path)
            all_photos.extend([ep for ep in extra_photos if os.path.exists(ep)])
            
            for i, pp in enumerate(all_photos):
                _, ext = os.path.splitext(pp)
                if i == 0:
                    photo_filename = f"{safe_name}{ext}"
                else:
                    photo_filename = f"{safe_name}_{i + 1}{ext}"
                photo_dest = os.path.join(temp_dir, photo_filename)
                try:
                    shutil.copy2(pp, photo_dest)
                    copied_photos.append(photo_filename)
                    photo_copied = True
                except Exception as e:
                    print(f"复制照片失败: {e}")
            
            # 3. 创建ZIP文件
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f"{safe_name}_学籍档案_{timestamp}.zip"
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".zip",
                filetypes=[("ZIP压缩包", "*.zip"), ("所有文件", "*.*")],
                initialfile=zip_filename
            )
            
            if file_path:
                with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    zipf.write(txt_path, arcname=txt_filename)
                    if photo_copied:
                        zipf.write(photo_dest, arcname=photo_filename)
                
                message = f"学生信息已导出为ZIP包:\n{file_path}\n\n包含文件:\n1. {txt_filename}"
                if copied_photos:
                    for i, pf in enumerate(copied_photos):
                        message += f"\n{i + 2}. {pf}"
                
                messagebox.showinfo("导出成功", message)
                
        except Exception as e:
            messagebox.showerror("导出失败", f"导出过程中出错: {str(e)}")
        finally:
            # 清理临时目录
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
    
    def export_all_zip(self):
        """批量导出所有学生信息为ZIP包"""
        if not hasattr(self, 'table_students') or not self.table_students:
            messagebox.showwarning("警告", "没有学生数据可导出")
            return
        
        confirm = messagebox.askyesno("确认", f"即将导出 {len(self.table_students)} 个学生档案，这可能需要一些时间，是否继续？")
        if not confirm:
            return
        
        # 创建临时目录
        temp_dir = tempfile.mkdtemp()
        
        try:
            total_count = 0
            success_count = 0
            
            # 显示进度
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("导出进度")
            progress_dialog.geometry("300x100")
            progress_dialog.transient(self.root)
            
            progress_label = ttk.Label(progress_dialog, text="准备导出...")
            progress_label.pack(pady=10)
            
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(progress_dialog, variable=progress_var, maximum=len(self.table_students))
            progress_bar.pack(pady=10, padx=20, fill=tk.X)
            
            progress_dialog.update()
            
            for idx, student in enumerate(self.table_students):
                try:
                    # 更新进度
                    progress = (idx + 1) / len(self.table_students) * 100
                    progress_var.set(idx + 1)
                    progress_label.config(text=f"正在导出第 {idx + 1}/{len(self.table_students)} 个学生: {student.get('姓名', '未知姓名')}")
                    progress_dialog.update()
                    
                    name = student.get('姓名', '未知姓名')
                    photo_path = student.get('照片路径')
                    extra_photos = student.get('extra_photos', [])
                    
                    safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
                    
                    student_dir = os.path.join(temp_dir, safe_name)
                    os.makedirs(student_dir, exist_ok=True)
                    
                    txt_filename = f"{safe_name}_学籍信息.txt"
                    txt_path = os.path.join(student_dir, txt_filename)
                    
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        content = self.get_display_content(student)
                        f.write('\n'.join(content))
                    
                    # 复制所有照片文件
                    all_photos = []
                    if photo_path and os.path.exists(photo_path):
                        all_photos.append(photo_path)
                    all_photos.extend([ep for ep in extra_photos if os.path.exists(ep)])
                    
                    for pi, pp in enumerate(all_photos):
                        _, ext = os.path.splitext(pp)
                        if pi == 0:
                            photo_filename = f"{safe_name}{ext}"
                        else:
                            photo_filename = f"{safe_name}_{pi + 1}{ext}"
                        photo_dest = os.path.join(student_dir, photo_filename)
                        try:
                            shutil.copy2(pp, photo_dest)
                        except Exception as e:
                            print(f"复制照片失败: {e}")
                    
                    success_count += 1
                    total_count += 1
                    
                except Exception as e:
                    print(f"处理学生 {name} 时出错: {e}")
                    total_count += 1
            
            # 关闭进度对话框
            progress_dialog.destroy()
            
            # 创建ZIP文件
            if success_count > 0:
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                zip_filename = f"所有学生学籍档案_{timestamp}.zip"
                
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".zip",
                    filetypes=[("ZIP压缩包", "*.zip"), ("所有文件", "*.*")],
                    initialfile=zip_filename
                )
                
                if file_path:
                    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                file_path_full = os.path.join(root, file)
                                arcname = os.path.relpath(file_path_full, temp_dir)
                                zipf.write(file_path_full, arcname=arcname)
                    
                    message = f"成功导出 {success_count}/{total_count} 个学生档案\n文件已保存为:\n{file_path}"
                    messagebox.showinfo("批量导出成功", message)
            else:
                messagebox.showwarning("导出失败", "未能导出任何学生档案")
                
        except Exception as e:
            messagebox.showerror("导出失败", f"批量导出过程中出错: {str(e)}")
        finally:
            # 清理临时目录
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

def main():
    """主函数"""
    root = tk.Tk()
    app = StudentRecordSystem(root)
    root.mainloop()

if __name__ == "__main__":
    main()
