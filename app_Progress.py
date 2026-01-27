import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import FancyBboxPatch
from datetime import timedelta, datetime
import numpy as np


def plot_gantt_chart(file_name):
    """
    從 Excel 檔案讀取資料並繪製甘特圖。
    - 繪圖區以每季為間隔的顏色顯示 (從2024年Q4開始交替)。
    - 不同群組的進度條顏色要不一樣，以柔和顏色為主。
    - 工作項目以最長的工作項目為基準靠左對齊。
    - 進度條上不顯示文字。 (此版本會在進度條"內"顯示完成百分比)
    - 進度條增加黑色粗框線。
    - 修正中文顯示問題。
    - X 軸以月為單位顯示，並改為兩列文字。
    - 完成百分比以進度條方式顯示。
    - 不忽略時間空白的列。
    - 在群組1和群組2中間新增X軸格線。
    - **新增：進度條改為圓角樣式**
    - **新增：Y軸格線以細實線顯示**

    Args:
        file_name (str): Excel 檔案的名稱 (預期為 'gantt_data.xlsx'，假設在同一資料夾)。
    """
    try:
        # --- 設定中文顯示 ---
        plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'Microsoft YaHei', 'PingFang HK',
                                           'Noto Sans CJK JP']
        plt.rcParams['axes.unicode_minus'] = False

        # 讀取 Excel 檔案
        df = pd.read_excel(file_name, header=0)  # 直接使用 df

        # 確認必要的欄位是否存在
        required_columns = ['工作項目', '開始日期', '結束日期', '完成百分比', '項次']
        if not all(col in df.columns for col in required_columns):
            missing_cols = [col for col in required_columns if col not in df.columns]
            raise KeyError(
                f"Excel 檔案中缺少必要的欄位：{', '.join(missing_cols)}。請確認您的欄位名稱與 '工作項目', '開始日期', '結束日期', '完成百分比', '項次' 完全匹配。"
            )

        # --- 處理空白/無效數據 ---
        # 1. 將日期欄位轉換為日期時間格式，並處理無效日期 (NaT)
        # 這裡不移除 NaT 的行，讓它們的索引保留
        df['開始日期'] = pd.to_datetime(df['開始日期'], errors='coerce')
        df['結束日期'] = pd.to_datetime(df['結束日期'], errors='coerce')

        # 2. 清理完成百分比數據：確保它是數字，並將 NaN 或無效值設為 0
        df['完成百分比'] = pd.to_numeric(df['完成百分比'], errors='coerce').fillna(0)
        # 確保完成百分比在 0 到 1 之間
        df['完成百分比'] = df['完成百分比'].clip(0, 1)
        # 如果原始數據是百分數 (例如 85 而不是 0.85)，則轉換
        if df['完成百分比'].max() > 1:
            df['完成百分比'] = df['完成百分比'] / 100

        # 計算任務持續時間（以天為單位）
        # 如果結束日期早於開始日期，則設定為開始日期 + 1天 (最短任務)
        # 如果開始或結束日期是 NaT，則持續時間會是 0，因此不會繪製條形
        df['持續時間'] = df.apply(
            lambda row: (row['結束日期'] - row['開始日期']).days if pd.notna(row['開始日期']) and pd.notna(
                row['結束日期']) and row['結束日期'] >= row['開始日期'] else 0, axis=1)

        # --- 定義柔和的群組顏色 ---
        group_colors = {
            1: '#fec76f',  # 黃色
            2: '#b3be62',  # 綠色
            3: '#ADCBE3',  # 柔和藍
            4: '#FFB8C2',  # 柔和粉
            5: '#D7BDE2',  # 柔和紫
        }
        default_group_color = '#CCCCCC'  # 淺灰色，用於未定義的群組

        # --- 動態調整左側邊距以容納最長的工作項目名稱 ---
        # 這裡使用原始 df 來計算，因為所有行都會顯示
        df['label_length'] = df['工作項目'].apply(
            lambda x: sum([2 if '\u4e00' <= char <= '\u9fff' else 1 for char in str(x)]))
        max_label_length = df['label_length'].max() if not df.empty else 10
        dynamic_left_margin = 0.12 + (max_label_length * 0.005)
        if dynamic_left_margin > 0.4:
            dynamic_left_margin = 0.4

        # 檢查是否有任何有效日期以設定 X 軸範圍，否則使用預設值
        valid_start_dates = df['開始日期'].dropna()
        valid_end_dates = df['結束日期'].dropna()

        min_date_overall = valid_start_dates.min() if not valid_start_dates.empty else datetime(2024, 10, 1)
        max_date_overall = valid_end_dates.max() if not valid_end_dates.empty else datetime.now() + timedelta(days=90)

        # 如果數據範圍非常小，確保有一個合理的顯示範圍
        if (max_date_overall - min_date_overall).days < 30:
            max_date_overall = min_date_overall + timedelta(days=90)
            if not valid_start_dates.empty:  # 如果有至少一個有效開始日期，從它開始擴展
                min_date_overall = valid_start_dates.min()
            else:  # 否則使用預設的 Q4 2024
                min_date_overall = datetime(2024, 10, 1)

        # 繪製甘特圖
        fig, ax = plt.subplots(figsize=(16, 9))

        # --- 繪製季度背景色塊 ---
        start_year = min_date_overall.year
        start_month = ((min_date_overall.month - 1) // 3) * 3 + 1
        current_quarter_start = datetime(start_year, start_month, 1)

        reference_date_q4_2024 = datetime(2024, 10, 1)
        quarter_colors_bg = ['#F0F0F0', '#FFFFFF']

        ref_quarter_index = (reference_date_q4_2024.month - 1) // 3

        while current_quarter_start <= max_date_overall + timedelta(days=31):
            total_quarters_diff = (current_quarter_start.year - reference_date_q4_2024.year) * 4 + \
                                  ((current_quarter_start.month - 1) // 3) - ref_quarter_index

            color_idx_bg = 0 if total_quarters_diff % 2 == 0 else 1

            next_month = current_quarter_start.month + 3
            next_year = current_quarter_start.year
            if next_month > 12:
                next_month -= 12
                next_year += 1
            quarter_end_date = datetime(next_year, next_month, 1)

            ax.axvspan(current_quarter_start, quarter_end_date,
                       facecolor=quarter_colors_bg[color_idx_bg],
                       alpha=0.7,
                       zorder=0)

            current_quarter_start = quarter_end_date  # 修正變數名稱

        # 為每個任務繪製水平條
        # 遍歷所有行，即使開始/結束日期為 NaT
        for i, row in df.iterrows():
            # 只有當開始日期、結束日期都不是 NaT 且持續時間有效時才繪製條形
            if pd.notna(row['開始日期']) and pd.notna(row['結束日期']) and row['持續時間'] >= 0:
                current_group_color = group_colors.get(row['項次'], default_group_color)
                bar_height = 0.6

                # 將日期轉換為數值以便計算位置和寬度
                start_num = mdates.date2num(row['開始日期'])
                duration_num = row['持續時間']

                # 繪製主進度條 (整個任務時長) - 使用圓角矩形
                main_bar = FancyBboxPatch(
                    (start_num, i - bar_height / 2),  # (x, y) 左下角位置
                    duration_num,  # 寬度
                    bar_height,  # 高度
                    boxstyle="round,pad=0,rounding_size=0.5",  # 圓角樣式
                    facecolor=current_group_color,
                    edgecolor='black',
                    linewidth=1.0,
                    zorder=1
                )
                ax.add_patch(main_bar)

                # 繪製完成百分比條 (在主進度條上方)
                completed_duration = row['持續時間'] * row['完成百分比']

                # 計算縮短後的完成進度條長度
                # 每個尾巴縮短 2 天，所以總共縮短 4 天
                # 確保縮短後的長度不為負數
                shorten_by_days = 2  # 每邊縮短的天數
                adjusted_completed_duration = max(0, completed_duration - (shorten_by_days * 2))

                if adjusted_completed_duration > 0:  # 只在有有效完成進度時繪製
                    # 完成進度條的起始點向右偏移 shorten_by_days 天
                    adjusted_start_num = start_num + shorten_by_days

                    completed_bar_height = bar_height * 0.3  # 完成條的高度不變

                    # 繪製完成進度條 - 使用圓角矩形
                    completed_bar = FancyBboxPatch(
                        (adjusted_start_num, i - completed_bar_height / 2),  # (x, y) 左下角位置
                        adjusted_completed_duration,  # 寬度
                        completed_bar_height,  # 高度
                        boxstyle="round,pad=0,rounding_size=0.08",  # 圓角樣式，比主條稍小的圓角
                        facecolor='#FFFFFF',  # 完成進度條的顏色改為白色
                        edgecolor='none',  # 邊框移除
                        zorder=2
                    )
                    ax.add_patch(completed_bar)

                    # 在完成進度條內顯示完成百分比文字
                    percentage_text = f'{int(row["完成百分比"] * 100)}%'
                    # 文字置中點需要根據新的起始點和長度計算
                    text_x_pos = adjusted_start_num + adjusted_completed_duration / 2
                    ax.text(text_x_pos,
                            i,
                            percentage_text,
                            va='center',
                            ha='center',
                            color='black',  # 文字顏色改為黑色，因為背景是白色
                            fontsize=9,
                            weight='bold',
                            zorder=3
                            )

        # 設定 Y 軸
        # 由於不移除任何行，直接使用 df.index 和 df['工作項目']
        ax.set_yticks(df.index)
        ax.set_yticklabels(df['工作項目'], fontsize=14, ha='right', va='center')
        ax.invert_yaxis()  # 將第一個任務放在圖表頂部

        # 手動設定 Y 軸範圍，確保所有進度條都能完整顯示
        if not df.empty:
            y_min = df.index.min() - 0.8  # 給頂部留一些空間
            y_max = df.index.max() + 0.8  # 給底部留一些空間
            ax.set_ylim(y_max, y_min)  # 注意：因為已經 invert_yaxis，所以順序相反

        # 設定 X 軸為月份格式 (兩列文字)
        ax.xaxis.set_major_locator(mdates.MonthLocator())

        def two_line_date_formatter(x, pos):
            date_obj = mdates.num2date(x)
            if date_obj.month == 1 or (pos == 0):
                return f"{date_obj.year}年\n{date_obj.month:02d}月"
            else:
                return f"{date_obj.month:02d}月"

        ax.xaxis.set_major_formatter(plt.FuncFormatter(two_line_date_formatter))
        ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=range(1, 13, 1)))

        fig.autofmt_xdate(rotation=0, ha='center')

        # 設定 X 軸的顯示範圍
        ax.set_xlim(min_date_overall - timedelta(days=7), max_date_overall + timedelta(days=30))

        # 設定圖表標題和軸標籤
        ax.set_title('專案甘特圖', fontsize=18, pad=25, weight='bold')
        ax.set_xlabel('日期', fontsize=14, labelpad=15)
        ax.set_ylabel('工作項目', fontsize=14, labelpad=15)

        # 增加右邊和上邊的邊框線
        ax.spines['right'].set_visible(True)
        ax.spines['top'].set_visible(True)
        ax.spines['left'].set_visible(True)
        ax.spines['bottom'].set_visible(True)

        # 增加 X 軸網格線
        ax.grid(True, linestyle='--', alpha=0.6, linewidth=0.5, axis='x')

        # 新增 Y 軸格線 - 以細實線顯示
        ax.grid(True, linestyle='--', alpha=0.6, linewidth=0.5, axis='y')

        # 在群組 1 和 2 中間新增水平格線
        # 這裡可以直接使用 df 的原始索引，因為沒有行被移除
        group_1_max_index = df[df['項次'] == 1].index.max()
        group_2_min_index = df[df['項次'] == 2].index.min()

        if pd.notna(group_1_max_index) and pd.notna(group_2_min_index) and group_2_min_index > group_1_max_index:
            line_y_pos = (group_1_max_index + group_2_min_index) / 2
            ax.axhline(y=line_y_pos, color='gray', linestyle='-', linewidth=1.5, zorder=0)
        elif pd.notna(group_2_min_index):  # 如果群組1沒有數據，但在群組2之前（例如只有群組2和3），則在群組2前劃線
            ax.axhline(y=group_2_min_index - 0.5, color='black', linestyle='-', linewidth=2.0, zorder=0)

        # 調整圖表邊距
        plt.subplots_adjust(left=dynamic_left_margin, right=0.95, top=0.9, bottom=0.18)

        # 顯示圖表
        plt.show()

    except FileNotFoundError:
        print(f"錯誤：找不到檔案 '{file_name}'。請檢查檔案名稱是否正確，並確認它與 Python 腳本在同一資料夾。")
    except KeyError as e:
        print(f"錯誤：{e}")
    except Exception as e:
        print(f"發生未知錯誤: {e}")


# --- 如何使用 ---
excel_file_name = 'gantt_data.xlsx'
plot_gantt_chart(excel_file_name)