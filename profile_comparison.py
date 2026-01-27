import pandas as pd
import matplotlib.pyplot as plt
import platform
import os


def set_civil_style():
    """設定中文字體與環境環境"""
    system = platform.system()
    if system == 'Windows':
        plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
    elif system == 'Darwin':
        plt.rcParams['font.sans-serif'] = ['Heiti TC']
    else:
        plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei']
    plt.rcParams['axes.unicode_minus'] = False
    pd.set_option('display.unicode.east_asian_width', True)


def generate_profile_analysis(file_name):
    set_civil_style()

    # 設定比較對象
    TARGET_NEW = "民國114年CECI測製"
    TARGET_OLD = "民國110年像片基本圖"

    # 1. 讀取資料
    try:
        df = pd.read_excel(file_name, engine='xlrd')
    except:
        df = pd.read_excel(file_name)

    # 2. 數據清洗
    periods = df['SRC_NAME'].unique()
    dfs = {}
    for p in periods:
        dfs[p] = df[df['SRC_NAME'] == p][['FIRST_DIST', 'FIRST_Z']].copy()
        dfs[p] = dfs[p].sort_values('FIRST_DIST').drop_duplicates('FIRST_DIST')

    # 3. 執行內插 (以 110年 為基準)
    if TARGET_NEW not in dfs or TARGET_OLD not in dfs:
        print(f"【錯誤】找不到指定名稱。現有時期：{list(periods)}")
        return

    master_grid = dfs[TARGET_NEW][['FIRST_DIST']].copy()
    combined_df = master_grid.copy()
    for p in periods:
        p_data = dfs[p].rename(columns={'FIRST_Z': f'高程_{p}'})
        combined_df = pd.merge(combined_df, p_data, on='FIRST_DIST', how='outer').sort_values('FIRST_DIST')
        combined_df[f'高程_{p}'] = combined_df[f'高程_{p}'].interpolate(method='linear', limit_direction='both')

    final_df = pd.merge(master_grid, combined_df, on='FIRST_DIST', how='left')
    final_df['高程差(Delta)'] = final_df[f'高程_{TARGET_NEW}'] - final_df[f'高程_{TARGET_OLD}']

    # 4. 輸出數據對照表
    print(f"\n{'=' * 35} 地形剖面數據對照表 {'=' * 35}")
    print(f"現況基準: {TARGET_NEW}")
    print(f"歷史背景: {TARGET_OLD}")
    print("-" * 85)
    cols_to_show = ['FIRST_DIST', f'高程_{TARGET_OLD}', f'高程_{TARGET_NEW}', '高程差(Delta)']
    print(final_df[cols_to_show].to_string(index=False, max_rows=20, float_format=lambda x: f"{x:,.3f}"))
    print("-" * 85)

    # 5. 繪圖設定
    plt.figure(figsize=(14, 7))

    # 高對比顏色配置
    colors = {TARGET_NEW: '#e74c3c', TARGET_OLD: '#2c3e50'}

    for p in periods:
        color = colors.get(p, '#7f8c8d')
        lw = 2.5 if p == TARGET_NEW else 1.8
        ls = '-' if p == TARGET_NEW else '--'
        plt.plot(final_df['FIRST_DIST'], final_df[f'高程_{p}'],
                 label=f'{p}', color=color, linestyle=ls, linewidth=lw)

    # 6. 【核心修正：填挖方區塊與圖例】
    # 填方區
    plt.fill_between(final_df['FIRST_DIST'], final_df[f'高程_{TARGET_OLD}'], final_df[f'高程_{TARGET_NEW}'],
                     where=(final_df[f'高程_{TARGET_NEW}'] >= final_df[f'高程_{TARGET_OLD}']),
                     color='green', alpha=0.15, label='堆積區')

    # 挖方區
    plt.fill_between(final_df['FIRST_DIST'], final_df[f'高程_{TARGET_OLD}'], final_df[f'高程_{TARGET_NEW}'],
                     where=(final_df[f'高程_{TARGET_NEW}'] < final_df[f'高程_{TARGET_OLD}']),
                     color='brown', alpha=0.15, label='侵蝕區')

    # 7. 格式化與 1:1 比例
    plt.title(f"地形剖面套疊比較圖：{TARGET_OLD} vs {TARGET_NEW}", fontsize=15, fontweight='bold')
    plt.xlabel("水平距離 (m)", fontsize=12)
    plt.ylabel("高程 (m)", fontsize=12)
    plt.gca().set_aspect('equal', adjustable='box')
    plt.grid(True, linestyle=':', alpha=0.6)

    # 圖例設定：顯示四項（包含填挖方）
    plt.legend(loc='upper left', frameon=True, shadow=True, fontsize=10)
    plt.tight_layout()

    # 8. 高解析度存檔
    output_filename = f"剖面分析結果_{TARGET_NEW}_含圖例.png"
    plt.savefig(output_filename, dpi=300, bbox_inches='tight')
    print(f"\n【出圖成功】高品質圖檔已儲存至：{os.path.abspath(output_filename)} (DPI: 300)")

    plt.show()


if __name__ == "__main__":
    target_xlsx = "地形剖面分析_LINE2.xls"
    try:
        generate_profile_analysis(target_xlsx)
    except Exception as e:
        print(f"【系統錯誤】：{e}")