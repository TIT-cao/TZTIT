import os
import re
import tkinter as tk
import pandas as pd
from pandas import ExcelWriter
from tkinter import Tk, filedialog
from tkinter.messagebox import showinfo, showerror
import warnings
import Levenshtein

# 忽略警告（可选）
warnings.filterwarnings("ignore", category=UserWarning, module="PIL.PngImagePlugin")

class WindowSelector:
    def __init__(self, options):
        self.result = None
        self.root = tk.Tk()
        self.root.title("TIT")

        # 每个组件大致的高度（可根据实际情况调整）
        prompt_height = 50 # 提示文字高度
        radio_button_height = 40 # 单选按钮高度
        button_height = 80 # 确认按钮高度
        padding = 40  # 上下边距

        # 计算窗体高度
        height = prompt_height + len(options) * radio_button_height + button_height + padding

        # 设置窗体大小
        self.root.geometry(f"300x{height}")

        # 添加提示性文字
        prompt_label = tk.Label(self.root, text="请选择:", font=('Arial', 20))
        prompt_label.pack(anchor=tk.W, pady=10)

        # 定义选项
        self.options = options
        self.var = tk.StringVar()
        if options:
            self.var.set(options[0])
        # 设置字体大小
        font_size = 20
        font_style = ('宋体', font_size)
        # 创建单选按钮
        for option in options:
            tk.Radiobutton(self.root, text=option, variable=self.var, value=option, font=font_style).pack(anchor=tk.W,
                                                                                                          padx=30)

        # 确认按钮
        confirm_button = tk.Button(self.root, text="确认选择", command=self.on_confirm, font=font_style)
        confirm_button.pack(pady=30)

    def on_confirm(self):
        selected = self.var.get()
        # messagebox.showinfo("选择结果", f"你选择了: {selected}")
        self.root.destroy()
        self.result = selected

    def run(self):
        self.root.mainloop()
        return self.result

class FileSelector:
    _instance = None  # 单例实例缓存
    _path = None  # 存储选择的路径
    _zd1 = None
    _list = None

    def __new__(cls):
        """确保全局唯一实例"""
        if not cls._instance:
            cls._instance = super().__new__(cls)
            cls._instance.root = tk.Tk()
            cls._instance.root.withdraw()
            cls._instance._init_once()
        return cls._instance

    def _init_once(self):
        """初始化操作只执行一次"""
        self._path = select_folder()
        self._list, self._zd1 = self.huquzidian()

    @property
    def path(self):
        """获取路径，首次调用时自动选择"""
        if self._path is None:
            self._path = select_folder()
        return self._path

    @property
    def zd1(self):
        return self._zd1

    @property
    def list(self):
        return self._list

    def refresh_path(self):
        """强制重新选择路径"""
        self._path = select_folder()

    def huquzidian(self):
        """"获取字典以及要遍历的列表"""
        df = pd.read_csv(self._path, header=None, names=[str(i) for i in range(24)],
                         encoding='ansi', skip_blank_lines=False)
        empty_index = df[df.iloc[:, 0].notnull()].index.tolist()
        empty_index.append(df.index.max())
        pair = {df.iloc[empty_index[i], 0]: empty_index[i:i + 2] for i in range(2, len(empty_index) - 1)}
        pname = [df.iloc[i, 0] for i in empty_index[2:-1]]

        zd_sheet = pd.read_excel(os.path.join(os.path.dirname(self._path), '批量表_v2.0.xlsm'), sheet_name='字典',
                                 header=None)
        zd = {zd_sheet.iloc[i, 0]: [zd_sheet.iloc[i, y] for y in range(1, zd_sheet.shape[1])] for i in zd_sheet.index}
        zd1 = {i: zd[similar(i, zd.keys(), 0.5)] for i in pname}

        return pair, zd1


def select_folder():
    """弹出选择对话框（可手动调用刷新路径）"""
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title='请选择程序模版文件',
        filetypes=[
            ("CSV 文件", "*.csv")
        ]
    )
    if not file_path:
        showinfo("取消", "用户取消选择")
        return None
    return file_path


def select_file(title, fmt="*.xlsx;*.xlsm;*.xls"):
    """选择文件并返回文件路径"""
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[
            ("format", fmt), ("Excel 文件", "*.xlsx;*.xlsm;*.xls"),
            ("Excel 97-2003", "*.xls"),
            ("CSV 文件", "*.csv")
        ]
    )
    if not file_path:
        showinfo("取消", "用户取消选择")
        return None
    return file_path


def similar(target, lst, threshold):
    """相似度匹配"""
    min_distance = float('inf')
    most_similar_char = None

    # 遍历列表，找出与目标字符最相似的元素
    for char in lst:
        distance = Levenshtein.distance(str(target), str(char))
        if distance < min_distance:
            min_distance = distance
            most_similar_char = char

    # 计算相似度得分，相似度得分范围是 0 到 1，值越大越相似
    similarity_score = 1 - (min_distance / max(len(target), len(most_similar_char)))

    # 如果相似度得分低于阈值，则不进行替换，返回原字符
    if similarity_score < threshold:
        return target
    return most_similar_char


def list_tq():
    """提取列非空单元格索引及值"""
    df = pd.read_csv(select_file('请选择模板文件', "*.csv"), header=None, names=[str(i) for i in range(24)],
                     encoding='ansi', skip_blank_lines=False)
    empty_index = df[df.iloc[:, 0].notnull()].index.tolist()
    empty_index.append(df.index.max())
    pair = [[empty_index[i], empty_index[i + 1] - 1] for i in range(2, len(empty_index) - 1)]
    pname = [df.iloc[i, 0] for i in empty_index[2:-1]]
    return pair, pname


def tag_out():
    """合并工作表提取位号"""
    input_file = select_file('请选择需要整理的位号表', ".xls")
    if not input_file:
        return
        # 要输入到的表格
    output_file = os.path.join(os.path.dirname(input_file), "批量表_v2.0.xlsm")

    try:
        # 根据文件类型读取数据
        if input_file.lower().endswith(('.xlsx', '.xlsm', '.xls')):
            # 读取Excel所需工作表
            all_sheets = pd.read_excel(input_file, sheet_name=['AI', 'AO', 'DI', 'DO'], header=None)
        elif input_file.lower().endswith('.csv'):
            # CSV视为单表
            all_sheets = {"CSV_DATA": pd.read_csv(input_file, header=None, encoding='ansi', skip_blank_lines=False)}
        else:
            raise ValueError("不支持的文件格式")

        # 合并数据
        merged_df = pd.concat(all_sheets.values(), ignore_index=True, keys=list(all_sheets.keys()))
        result = {v[1]: v[2] for k, v in merged_df.iloc[:, 1:3].iterrows() if v[2] != "备用"}
        pattern = r'^.*(?=_)|^.*'
        nested_dict = {
            k:
                {tihuan(merged_df.iloc[i, 1], list(result))
                 if re.match('.*_AO', merged_df.iloc[i, 1]) else re.findall(pattern, merged_df.iloc[i, 1])[0]
                 : [result.get(tihuan(merged_df.iloc[i, 1], list(result)), "/")
                    if re.match('.*_AO', merged_df.iloc[i, 1]) else re.findall(pattern, merged_df.iloc[i, 2])[0],
                    re.findall(pattern, merged_df.iloc[i, 1])[0] if re.match('.*_AO', merged_df.iloc[i, 1]) else ""]
                 for i in range(len(merged_df)) if re.match(j[0], merged_df.iloc[i, 1] + merged_df.iloc[i, 2])
                 } for k, j in selector.zd1.items()
        }
        data = []
        for outer_key, inner_dict in nested_dict.items():
            # 遍历内层字典
            for inner_key, value in inner_dict.items():
                # 将数据添加到列表中
                data.append([outer_key, inner_key, value[0], value[1]])
        merged_df = pd.DataFrame(data, columns=['类型', '位号', '描述', '_AO.'])
        # 保存结果
        # wb = load_workbook(output_file, keep_vba=True)
        with ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            # 读取 Excel 文件中的数据
            df = pd.read_excel(output_file, header=None)
            # 清空从第 5 行开始到最后一行的前四列内容
            df.iloc[6:, :4] = pd.NA
            # 将修改后的数据写回原工作表
            df.to_excel(writer, sheet_name='表', index=False, header=False)
            merged_df.to_excel(writer, sheet_name='表', startrow=5, index=False, header=False)

        showinfo("完成", f"文件已合并并保存至：\n{output_file}")

    except Exception as e:
        showerror("错误", f"操作失败：{str(e)}")


def tihuan(tex, lit):
    pattern = r'^.*(?=调节阀)|^.*(?=_AO)|^.*泵'
    matches = re.findall(pattern, tex)
    # result = [val for sublist in zip(df.iloc[:,1], df.iloc[:,2]) for val in sublist]
    if matches:
        name = matches[0]
        jieguo = similar(name, lit, 0.85)
        return jieguo
    return tex


def program_pages():
    """根据文件生成程序页"""
    # 读取程序模版文件
    try:
        # 程序模版文件
        model = pd.read_csv(selector.path, header=None, names=[str(i) for i in range(24)], encoding='ansi',
                            skip_blank_lines=False)
        # 位号表文件
        tag = pd.read_excel(os.path.join(os.path.dirname(selector.path, ), "批量表_v2.0.xlsm"),
                            sheet_name='表', header=None, engine='openpyxl')
        model1 = model.iloc[:4]
        pages = 0
        for only in tag.iloc[5:, 0].unique():
            model2 = model.iloc[selector.list[only][0]:selector.list[only][1]].reset_index(drop=True)
            tag_1 = tag[tag.iloc[:, 0] == only].reset_index(drop=True)
            k = -1
            j = 1
            while k < len(tag_1) - 1:
                k = loop_creation(model1.copy(), model2.copy(), tag, tag_1, only, k, j)
                j += 1
            pages += j - 1
        showinfo("完成", f"成功生成了{pages}份文件\n保存路径：{os.path.dirname(selector.path)}")
    except Exception as e:
        showerror("错误", f"读取文件失败：{str(e)}")
        return None, None


def loop_creation(model, model_2, tag, tag_1, only, k, j):
    ago = selector.zd1[only][3]
    after = selector.zd1[only][4]
    pattern1 = r'^.*(?=_)|^.*(?=\.)|^.+'
    # 处理表头
    model.iloc[1, 0] = f'{model_2.iloc[0, 0]}_{j:02d}'
    model.iloc[1, 1] = f'{tag.iloc[1, 1]}{selector.zd1[only][2]}_{j:02d}'
    model.iloc[1, 2] = tag.iloc[1, 2]
    model.iloc[1, 3] = (j - 1) % 2
    zd = {tag.iloc[4, i]: i for i in range(len(tag.iloc[4]))}
    js = k
    k_dict={}
    # 处理逻辑名称
    z = 0
    ls1 = ls = x = ""
    for i in range(0, len(model_2)):
        if k >= len(tag_1) - 1:
            break
        if model_2.iloc[i, 3] == model_2.iloc[0, 3]:
            k += 1
            model_2.iloc[z, 23] = tag_1.iloc[k, 1].replace(ago, after, 1)
            z += 1
        k_dict.setdefault(model_2.iloc[i, 3], js)
        k_dict[model_2.iloc[i, 3]] += 1
        x = model_2.iloc[i, 3] if pd.notna(model_2.iloc[i, 3]) else x
        r_js = k_dict[x]
        # 替换逻辑名称
        if not (pd.isna(model_2.iloc[i, 9])):
            ls1 = str(model_2.iloc[i, 9])
        if not (pd.isna(model_2.iloc[i, 18])):
            ls = ls1 + str(model_2.iloc[i, 18])
        # 通用字段替换逻辑
        for col in [4, 5, 6, 10, 13, 19]:  # 根据实际列索引调整
            if re.search(tag.iloc[4, 3], str(model_2.iloc[i, col])):
                m = 3
            else:
                m = 1
            if not (pd.isna(model_2.iloc[i, col])):
                match col:
                    case 5:
                        model_2.iloc[i, col] = re.sub(pattern1, tag_1.iloc[r_js, 2], str(model_2.iloc[i, col]),
                                                      flags=re.IGNORECASE)
                    case 4:
                        model_2.iloc[i, col] = re.sub(pattern1, tag_1.iloc[r_js, 1], str(model_2.iloc[i, col]),
                                                      flags=re.IGNORECASE).replace(ago, after, 1)
                    case 6:
                        model_2.iloc[i, col] = tag.iloc[1, 4]

                    case 19:
                        model_2.iloc[i, col] = re.sub(pattern1, str(tag_1.iloc[r_js, zd.get(ls)]),
                                                      str(model_2.iloc[i, col]),
                                                      flags=re.IGNORECASE)
                    case _:
                        model_2.iloc[i, col] = re.sub(pattern1, str(tag_1.iloc[r_js, m]), str(model_2.iloc[i, col]),
                                                      flags=re.IGNORECASE)
    lsm = 0
    while not (pd.isna(model_2.iloc[lsm, 22])):
        if not (pd.isna(model_2.iloc[lsm, 22])) and (pd.isna(model_2.iloc[lsm, 23])):
            model_2.iloc[lsm, 23] = str(model_2.iloc[lsm, 22])
        lsm += 1

    result = pd.concat([model, model_2], axis=0, ignore_index=True)
    output_filename = f"{result.iloc[4, 0]}_{str(j).zfill(2)}.csv"
    output_path = os.path.join(os.path.dirname(selector.path), output_filename)
    result.to_csv(output_path, index=False, header=False, encoding='ansi')

    return k


if __name__ == "__main__":
    selector = WindowSelector(["生成位号","生成程序页"])
    choice =  selector.run()
    if choice:
        selector = FileSelector()
        match choice:
            case '生成位号':
                tag_out()
            case '生成程序页':
                program_pages()
