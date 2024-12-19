import os
import re
import pandas as pd
import zipfile
import chardet
import tempfile
import streamlit as st

# 定义处理函数（使用你提供的代码，并做适当修改以适应上传文件）

def unzip_file(zip_path, unzip_dir):
    if not os.path.exists(zip_path):
        st.error(f"ZIP 文件不存在: {zip_path}")
        return False
    os.makedirs(unzip_dir, exist_ok=True)
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(unzip_dir)
        st.success(f"文件已解压到: {unzip_dir}")
        return True
    except zipfile.BadZipFile:
        st.error(f"无法解压 ZIP 文件: {zip_path}")
        return False
    except Exception as e:
        st.error(f"解压文件时出错: {e}")
        return False

def read_file_contents(file_path):
    with open(file_path, 'rb') as f:
        raw_data = f.read()
        result = chardet.detect(raw_data)
        encoding = result['encoding']
    with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
        return f.readlines()

def extract_highest_x(file_contents):
    highest_x = 0
    for line in file_contents:
        x_matches = re.findall(r'X([0-9]*\.?[0-9]+)', line)
        for x in x_matches:
            try:
                x_value = float(x)
                if x_value > highest_x:
                    highest_x = x_value
            except ValueError:
                continue
    return highest_x

def extract_additional_columns(file_contents):
    additional_columns = {
        'Number of piercings': '',
        'Piercing time': '',
        'Cutting time': '',
        'Machining total time': '',
        'Cutting distance': '',
        'Rapid distance': '',
        'Highest X value': extract_highest_x(file_contents),
        'Tube Type': ''
    }

    for line in file_contents:
        if '; Number of piercings:' in line:
            additional_columns['Number of piercings'] = line.split(':')[-1].strip()
        elif '; Piercing time:' in line:
            additional_columns['Piercing time'] = line.split(':')[-1].strip()
        elif '; Cutting time:' in line:
            additional_columns['Cutting time'] = line.split(':')[-1].strip()
        elif '; Machining total time:' in line:
            additional_columns['Machining total time'] = line.split(':')[-1].strip()
        elif '; Cutting distance:' in line:
            additional_columns['Cutting distance'] = line.split(':')[-1].strip()
        elif '; Rapid distance:' in line:
            additional_columns['Rapid distance'] = line.split(':')[-1].strip()
        elif '; Tube Type:' in line or 'PART MODEL' in line:
            st.write(f"找到 Tube Type 行: {line}")  # 调试信息
            additional_columns['Tube Type'] = '方管' if 'M_TUBE' in line else '圆管'
    return additional_columns

def extract_quantity_from_filename(file_name):
    match = re.search(r'S-(\d+)', file_name)
    if match:
        return int(match.group(1))
    return 1  # 默认数量为1

def extract_v_number(file_name):
    match = re.search(r'-V(\d+)-', file_name)
    if match:
        return int(match.group(1))
    return None

def split_n20(n20):
    st.write(f"正在解析 N20 字段: {n20}")  # 调试信息
    if pd.isna(n20):
        return "", ""
    # 使用正则表达式提取 L 和 D
    size_match = re.search(r'L=([\d\.]+)\s+D=([\d\.]+)', n20, re.IGNORECASE)
    thickness_match = re.search(r'THICKNESS=([\d\.]+)', n20, re.IGNORECASE)

    # 将 L 和 D 拼接成 "80*80" 格式
    if size_match:
        size = f"{size_match.group(1)}*{size_match.group(2)}"
    else:
        size = ""

    thickness = thickness_match.group(1) if thickness_match else ""
    return size, thickness

def translate_datetime(value):
    if not isinstance(value, str):
        return value
    value = value.replace("DATE", "日期").replace("TIME", "时间")
    return value

def translate_n30(value):
    if not isinstance(value, str):
        return value
    value = value.replace("MATERIAL: stainless steel", "材料: 不锈钢")
    return value

def process_single_nc_file(file_path):
    file_name = os.path.basename(file_path)
    try:
        file_contents = read_file_contents(file_path)
    except Exception as e:
        st.error(f"读取文件 {file_name} 时出错: {e}")
        return None

    # 提取相关的 N 行
    relevant_lines = [line for line in file_contents if any(tag in line for tag in ["N5", "N15", "N20", "N23", "N25", "N30"])]

    # 提取额外列
    additional_columns = extract_additional_columns(file_contents)

    # 提取数量和 V Number
    quantity_from_filename = extract_quantity_from_filename(file_name)
    v_number = extract_v_number(file_name)

    # 提取 N-lines
    data = {
        'File Name': file_name,
        'Quantity from Filename': quantity_from_filename,
        'V Number': v_number
    }

    for line in relevant_lines:
        parts = line.split(' ', 1)
        if len(parts) > 1:
            key = parts[0]
            value = parts[1].strip()
            data[key] = value
        else:
            data[parts[0]] = ''

    # 添加额外列
    data.update(additional_columns)

    # 拆分 N20
    size, thickness = split_n20(data.get('N20', ''))
    data['SIZE'] = size
    data['THICKNESS'] = thickness
    data.pop('N20', None)  # 移除原始 N20 列

    # 创建 DataFrame
    df = pd.DataFrame([data])

    # 重新命名列
    translated_columns = {
        'V Number': 'V 号码',
        'File Name': '文件名',
        'N5': 'N5',
        'Quantity from Filename': '数量',
        'N15': 'N15',
        'SIZE': '尺寸',
        'THICKNESS': '厚度',
        'N23': 'N23',
        'N25': 'N25',
        'N30': 'N30',
        'Number of piercings': '穿孔数量',
        'Piercing time': '穿孔时间',
        'Cutting time': '切割时间',
        'Machining total time': '总加工时间',
        'Cutting distance': '切割距离',
        'Rapid distance': '快速移动距离',
        'Highest X value': '最高 X 值（需要长度）',
        'Tube Type': '管类型'
    }
    df = df.rename(columns=translated_columns)

    # 翻译 N5 列中的 DATE 和 TIME
    if 'N5' in df.columns:
        df['N5'] = df['N5'].apply(translate_datetime)

    # 翻译 N30 列的值
    if 'N30' in df.columns:
        df['N30'] = df['N30'].apply(translate_n30)

    return df

def process_multiple_nc_files(unzip_dir):
    # 获取解压目录中的所有 .NC 文件
    file_paths = [os.path.join(unzip_dir, file) for file in os.listdir(unzip_dir) if file.endswith('.NC')]
    all_dfs = []

    for file_path in file_paths:
        file_name = os.path.basename(file_path)
        st.write(f"正在处理文件: {file_name}")
        df = process_single_nc_file(file_path)
        if df is not None:
            all_dfs.append(df)

    if all_dfs:
        # 合并所有 DataFrame
        final_df = pd.concat(all_dfs, ignore_index=True)

        # 重新排列列的顺序
        columns_order_final = [
            '文件名', 'N5', '数量', '最高 X 值（需要长度）', '尺寸', '厚度', 'N30', '管类型',
            '穿孔数量', '穿孔时间', '切割时间', '总加工时间', '切割距离', '快速移动距离', 'N15', 'N23', 'N25'
        ]

        # 确保所有列都存在
        columns_order_final = [col for col in columns_order_final if col in final_df.columns]
        final_df = final_df[columns_order_final]

        return final_df
    else:
        st.warning("未找到任何有效的 .NC 文件或未能提取数据。")
        return None

def main():
    st.title("ZIP 转 Excel 工具")
    st.write("上传一个包含 .NC 文件的 ZIP 文件，系统将自动生成 Excel 文件。")

    uploaded_zip = st.file_uploader("选择 ZIP 文件", type="zip")

    if uploaded_zip is not None:
        with tempfile.TemporaryDirectory() as tmpdirname:
            zip_path = os.path.join(tmpdirname, uploaded_zip.name)
            with open(zip_path, 'wb') as f:
                f.write(uploaded_zip.getbuffer())
            unzip_dir = os.path.join(tmpdirname, "unzipped_files")
            if unzip_file(zip_path, unzip_dir):
                final_df = process_multiple_nc_files(unzip_dir)
                if final_df is not None:
                    # 提供下载链接
                    excel_path = os.path.join(tmpdirname, "output.xlsx")
                    final_df.to_excel(excel_path, index=False)

                    with open(excel_path, "rb") as f:
                        st.download_button(
                            label="下载生成的 Excel",
                            data=f,
                            file_name="output.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

if __name__ == "__main__":
    main()
