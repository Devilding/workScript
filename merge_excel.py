import pandas as pd
import os
# 导入材料类型枚举
from material_types import MATERIAL_TYPES as material_types

def merge_excel_files(file_paths, output_path):
    # 创建一个空的DataFrame用于存储合并后的数据
    merged_df = pd.DataFrame()
    
    # 遍历所有Excel文件
    for file_path in file_paths:
        # 读取Excel文件中的所有工作表
        xls = pd.ExcelFile(file_path)
        
        # 遍历每个工作表
        for sheet_name in xls.sheet_names:
            # 读取工作表数据，跳过第一行（表头）
            df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
            
            # 如果是第一个文件，保留表头
            if merged_df.empty:
                merged_df = df
            else:
                # 追加数据，忽略表头，并确保列对齐
                merged_df = pd.concat([merged_df, df], ignore_index=True, sort=False)
    
    # 将"数量"列转换为数字类型
    merged_df['数量'] = pd.to_numeric(merged_df['数量'], errors='coerce')
    
    # 创建新的Excel文件
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 将合并后的数据写入"总"工作表
        merged_df.to_excel(writer, sheet_name='总', index=False)
        
        # 筛选非标底座
        non_standard = merged_df[merged_df['规格'].str.contains('非标底座', na=False)]
        if not non_standard.empty:
            non_standard.to_excel(writer, sheet_name='非标总', index=False)
        
        # 创建"刷漆Y总"和"镀锌G总"工作表
        # 筛选Y和G的记录，并去除非标底座
        y_paint = merged_df[(merged_df['表面处理'] == 'Y') & 
                          (~merged_df['规格'].str.contains('非标底座', na=False))]
        g_paint = merged_df[(merged_df['表面处理'] == 'G') & 
                          (~merged_df['规格'].str.contains('非标底座', na=False))]
        
        

        # 处理刷漆Y总数据
        if not y_paint.empty:
            # 创建副本以避免修改原始数据
            y_paint_modified = y_paint.copy()
            # 筛选并修改符合条件的行
            condition = (y_paint_modified['是否带腹板'] == '是') & \
                       (y_paint_modified['规格'].str.startswith(('LDK', 'L4', 'L5')))
            y_paint_modified.loc[condition, '规格'] = y_paint_modified.loc[condition, '规格'] + ' P'
            y_paint_modified.to_excel(writer, sheet_name='刷漆Y总', index=False)

            # 对刷漆Y总进行分类筛选
            for material, prefixes in material_types.items():
                # 筛选符合前缀的行
                filtered = y_paint_modified[y_paint_modified['规格'].str.startswith(tuple(prefixes))]
                if not filtered.empty:
                    # 只保留规格和数量两列
                    filtered = filtered[['规格', '数量']]
                    # 按规格分组并求和数量
                    filtered = filtered.groupby('规格', as_index=False)['数量'].sum()
                    # 按规格排序
                    filtered = filtered.sort_values('规格')
                    filtered.to_excel(writer, sheet_name=f'{material}Y', index=False)
        
        # 处理镀锌G总数据
        if not g_paint.empty:
            # 创建副本以避免修改原始数据
            g_paint_modified = g_paint.copy()
            # 筛选并修改符合条件的行
            condition = (g_paint_modified['是否带腹板'] == '是') & \
                       (g_paint_modified['规格'].str.startswith(('LDK', 'L4', 'L5')))
            g_paint_modified.loc[condition, '规格'] = g_paint_modified.loc[condition, '规格'] + ' P'
            g_paint_modified.to_excel(writer, sheet_name='镀锌G总', index=False)

            # 对镀锌G总进行分类筛选
            for material, prefixes in material_types.items():
                # 筛选符合前缀的行
                filtered = g_paint_modified[g_paint_modified['规格'].str.startswith(tuple(prefixes))]
                if not filtered.empty:
                    # 只保留规格和数量两列
                    filtered = filtered[['规格', '数量']]
                    # 按规格分组并求和数量
                    filtered = filtered.groupby('规格', as_index=False)['数量'].sum()
                    # 按规格排序
                    filtered = filtered.sort_values('规格')
                    filtered.to_excel(writer, sheet_name=f'{material}G', index=False)

if __name__ == "__main__":
    # 获取当前目录下的所有Excel文件（使用完整路径）
    excel_files = [os.path.join(os.getcwd(), f) 
                  for f in os.listdir() 
                  if f.endswith('.xlsx') or f.endswith('.xls')]
    
    if not excel_files:
        print("当前目录下未找到Excel文件")
    else:
        # 设置输出文件路径
        output_file = os.path.join(os.getcwd(), 'text.xlsx')
        
        # 合并Excel文件
        try:
            merge_excel_files(excel_files, output_file)
            print(f"合并完成，结果已保存到 {output_file}")
        except Exception as e:
            print(f"合并过程中发生错误: {str(e)}")
