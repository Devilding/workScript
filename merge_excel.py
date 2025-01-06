import pandas as pd
import os
from material_types import MATERIAL_TYPES as material_types

def process_flat_bars(filtered):
    """对扁钢材料进行长度计算"""
    if '规格' not in filtered.columns:
        return filtered
        
    if '长度' not in filtered.columns:
        filtered.insert(1, '长度', 0)
    
    split_cols = filtered['规格'].str.split(r'[-XP]', expand=True)
    a = split_cols[0].fillna('0')
    b = pd.to_numeric(split_cols[1], errors='coerce').fillna(0)
    c = pd.to_numeric(split_cols[2], errors='coerce').fillna(0)
    d = pd.to_numeric(split_cols[3], errors='coerce').fillna(0)
    e = pd.to_numeric(split_cols[4], errors='coerce').fillna(0)
    
    filtered.loc[a == 'FB', '长度'] = c + 2*d - 10
    filtered.loc[a == 'FBF', '长度'] = d + 2*e - 10
    filtered.loc[a == 'FBZ', '长度'] = c + 2*d + 250 - 15

    filtered = filtered.groupby('规格', as_index=False).agg({
                        '长度': 'first',
                        '数量': 'sum'
                    })
                    
    filtered = filtered.sort_values('规格')
    return filtered

def process_pedestal(filtered):
    if not filtered.empty:  
        # 按规格分组并汇总数量
        filtered = filtered.groupby('规格', as_index=False).agg({
            '数量': 'sum'
        })
        
        # 根据规格排序
        filtered = filtered.sort_values('规格')
    else:
        return filtered
            
    return filtered

def process_angle_steel(filtered):
    """处理角钢材料"""
    if filtered.empty:
        return filtered
        
    # 按规格分组并汇总数量
    filtered = filtered.groupby('规格', as_index=False).agg({
        '数量': 'sum'
    })
    
    # 区分带P和不带P的规格
    filtered['带P'] = filtered['规格'].str.contains('P')
    
    # 创建新的DataFrame用于存储结果
    result_df = pd.DataFrame()
    
    # 处理带P的规格
    p_df = filtered[filtered['带P']]
    if not p_df.empty:
        p_df = p_df[['规格', '数量']]
        p_df.columns = ['规格_P', '数量_P']
        result_df = pd.concat([result_df, p_df], axis=1)
    
    # 处理不带P的规格
    non_p_df = filtered[~filtered['带P']]
    if not non_p_df.empty:
        non_p_df = non_p_df[['规格', '数量']]
        non_p_df.columns = ['规格', '数量']
        result_df = pd.concat([result_df, non_p_df], axis=1)
    
    # 根据规格排序
    result_df = result_df.sort_values(by=['规格', '规格_P'])
    
    return result_df

def process_steel_pipe(filtered):
    """处理钢管材料"""
    if filtered.empty:
        return filtered
        
    # 按规格分组并汇总数量
    filtered = filtered.groupby('规格', as_index=False).agg({
        '数量': 'sum'
    })
    
    # 添加长度列
    filtered.insert(1, '长度', 0)
    
    # 处理规格字符串
    for index, row in filtered.iterrows():
        spec = row['规格']
        # 舍弃'Ф'前所有字符串
        spec = spec[spec.find('Ф'):]
        # 用'-'分隔
        parts = spec.split('-')
        if len(parts) > 1:
            filtered.at[index, '规格'] = parts[0]
            filtered.at[index, '长度'] = parts[1]
    
    return filtered

def merge_excel_files(file_paths, output_path):
    merged_df = pd.DataFrame()
    
    for file_path in file_paths:
        xls = pd.ExcelFile(file_path)
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
            
            if merged_df.empty:
                merged_df = df
            else:
                merged_df = pd.concat([merged_df, df], ignore_index=True, sort=False)
    
    merged_df['数量'] = pd.to_numeric(merged_df['数量'], errors='coerce')
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        merged_df.to_excel(writer, sheet_name='总', index=False)
        
        non_standard = merged_df[merged_df['规格'].str.contains('非标底座', na=False)]
        if not non_standard.empty:
            non_standard.to_excel(writer, sheet_name='非标总', index=False)
        
        y_paint = merged_df[(merged_df['表面处理'] == 'Y') & 
                          (~merged_df['规格'].str.contains('非标底座', na=False))]
        g_paint = merged_df[(merged_df['表面处理'] == 'G') & 
                          (~merged_df['规格'].str.contains('非标底座', na=False))]
        
        if not y_paint.empty:
            y_paint_modified = y_paint.copy()
            condition = (y_paint_modified['是否带腹板'] == '是') & \
                       (y_paint_modified['规格'].str.startswith(('LDK', 'L4', 'L5')))
            y_paint_modified.loc[condition, '规格'] = y_paint_modified.loc[condition, '规格'] + ' P'
            y_paint_modified.to_excel(writer, sheet_name='刷漆Y总', index=False)

            for material, prefixes in material_types.items():
                filtered = y_paint_modified[y_paint_modified['规格'].str.startswith(tuple(prefixes))]
                if not filtered.empty:
                    filtered = filtered[['规格', '数量']]
                    if material == '扁钢':
                        filtered = process_flat_bars(filtered)
                        #filtered.to_excel(writer, sheet_name=f'{material}Y', index=False)
                    if material == '基座':
                        filtered = process_pedestal(filtered)
                        #filtered.to_excel(writer, sheet_name=f'{material}Y', index=False)
                    if material == '角钢':
                        filtered = process_angle_steel(filtered)
                        #filtered.to_excel(writer, sheet_name=f'{material}Y', index=False)
                    if material == '钢管':
                        filtered = process_steel_pipe(filtered)
                    filtered.to_excel(writer, sheet_name=f'{material}Y', index=False)

        if not g_paint.empty:
            g_paint_modified = g_paint.copy()
            condition = (g_paint_modified['是否带腹板'] == '是') & \
                       (g_paint_modified['规格'].str.startswith(('LDK', 'L4', 'L5')))
            g_paint_modified.loc[condition, '规格'] = g_paint_modified.loc[condition, '规格'] + ' P'
            g_paint_modified.to_excel(writer, sheet_name='镀锌G总', index=False)

            for material, prefixes in material_types.items():
                filtered = g_paint_modified[g_paint_modified['规格'].str.startswith(tuple(prefixes))]
                if not filtered.empty:
                    filtered = filtered[['规格', '数量']]
                    if material == '扁钢':
                        filtered = process_flat_bars(filtered)
                        #filtered.to_excel(writer, sheet_name=f'{material}G', index=False)
                    if material == '基座':
                        filtered = process_pedestal(filtered)
                        #filtered.to_excel(writer, sheet_name=f'{material}G', index=False)
                    if material == '角钢':
                        filtered = process_angle_steel(filtered)
                        #filtered.to_excel(writer, sheet_name=f'{material}G', index=False)
                    if material == '钢管':
                        filtered = process_steel_pipe(filtered)
                    filtered.to_excel(writer, sheet_name=f'{material}Y', index=False)

if __name__ == "__main__":
    excel_files = [os.path.join(os.getcwd(), f) 
                  for f in os.listdir() 
                  if f.endswith('.xlsx') or f.endswith('.xls')]
    
    if not excel_files:
        print("当前目录下未找到Excel文件")
    else:
        output_file = os.path.join(os.getcwd(), 'text.xlsx')
        
        try:
            merge_excel_files(excel_files, output_file)
            print(f"合并完成，结果已保存到 {output_file}")
        except Exception as e:
            print(f"合并过程中发生错误: {str(e)}")
