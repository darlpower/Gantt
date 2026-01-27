import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, html, dcc
from datetime import timedelta, datetime


def create_gantt_figure(file_name):
    """
    從 Excel 檔案讀取資料並創建 Plotly 甘特圖物件。
    包含 Matplotlib 版本中的所有定制功能。
    """
    # 在 try 區塊外部初始化 fig，確保它始終被定義
    fig = go.Figure()

    try:
        df_original = pd.read_excel(file_name, header=0)
        df = df_original.copy()

        required_columns = ['工作項目', '開始日期', '結束日期', '完成百分比', '項次']
        if not all(col in df.columns for col in required_columns):
            missing_cols = [col for col in required_columns if col not in df.columns]
            raise KeyError(
                f"Excel 檔案中缺少必要的欄位：{', '.join(missing_cols)}。請確認您的欄位名稱與 '工作項目', '開始日期', '結束日期', '完成百分比', '項次' 完全匹配。")

        df['開始日期'] = pd.to_datetime(df['開始日期'], errors='coerce')
        df['結束日期'] = pd.to_datetime(df['結束日期'], errors='coerce')
        df['持續時間'] = (df['結束日期'] - df['開始日期']).dt.days

        df['完成百分比'] = pd.to_numeric(df['完成百分比'], errors='coerce').fillna(0)
        if df['完成百分比'].max() > 1:
            df['完成百分比'] = df['完成百分比'] / 100

        # --- 定義柔和的群組顏色 ---
        group_colors = {
            1: '#fec76f',  # 黃
            2: '#b3be62',  # 綠
            3: '#ADCBE3',  # 柔和藍
            4: '#FFB8C2',  # 柔和粉
            5: '#D7BDE2',  # 柔和紫
        }
        default_group_color = '#CCCCCC'  # 淺灰色

        # --- 創建 Plotly 圖表物件 ---
        # 這裡不再需要 fig = go.Figure() 因為已經在 try 外部初始化

        # 獲取所有獨特的工作項目，並為其分配一個數值Y軸位置
        y_labels = df_original['工作項目'].tolist()
        y_labels_reversed = y_labels[::-1]

        # --- 繪製季度背景色塊 ---
        min_date = df['開始日期'].min() if not df['開始日期'].empty else datetime.now().replace(year=2024, month=10,
                                                                                                day=1)
        max_date = df['結束日期'].max() if not df['結束日期'].empty else datetime.now() + timedelta(days=90)

        if pd.isna(min_date): min_date = datetime.now().replace(year=2024, month=10, day=1)
        if pd.isna(max_date) or max_date < min_date: max_date = min_date + timedelta(days=90)

        reference_date = datetime(2024, 10, 1)
        start_year, start_month = min_date.year, min_date.month
        start_quarter_month = ((start_month - 1) // 3) * 3 + 1
        current_date = datetime(start_year, start_quarter_month, 1)

        quarter_colors_bg = ['#F5F5F5', 'white']

        ref_quarter_idx = (reference_date.month - 1) // 3
        total_quarters_since_ref = (current_date.year - reference_date.year) * 4 + \
                                   ((current_date.month - 1) // 3) - ref_quarter_idx
        color_idx_bg = total_quarters_since_ref % 2
        if total_quarters_since_ref < 0:
            color_idx_bg = abs(total_quarters_since_ref) % 2
            color_idx_bg = 1 - color_idx_bg

        shapes = []
        while current_date <= max_date + timedelta(days=31):
            next_quarter_month = current_date.month + 3
            next_year = current_date.year
            if next_quarter_month > 12:
                next_quarter_month -= 12
                next_year += 1
            quarter_end_date = datetime(next_year, next_quarter_month, 1)

            shapes.append(
                dict(
                    type="rect",
                    xref="x", yref="paper",
                    x0=current_date, y0=0,
                    x1=quarter_end_date, y1=1,
                    fillcolor=quarter_colors_bg[color_idx_bg % 2],
                    layer="below",
                    line_width=0,
                    opacity=0.8
                )
            )
            current_date = quarter_end_date
            color_idx_bg += 1

        # 將背景色塊添加到圖表佈局中
        fig.update_layout(shapes=shapes)

        # --- 為每個任務繪製水平條 ---
        for i, row in df.iterrows():
            if pd.notna(row['開始日期']) and pd.notna(row['結束日期']) and row['持續時間'] >= 0:
                current_group_color = group_colors.get(row['項次'], default_group_color)

                # 繪製主任務條
                fig.add_trace(go.Bar(
                    y=[row['工作項目']],
                    x=[timedelta(days=row['持續時間'])],
                    base=[row['開始日期']],
                    orientation='h',
                    marker=dict(
                        color=current_group_color,
                        line=dict(color='black', width=1.5)
                    ),
                    width=0.6,
                    hovertemplate=f"工作項目: {row['工作項目']}<br>開始日期: {row['開始日期'].strftime('%Y-%m-%d')}<br>結束日期: {row['結束日期'].strftime('%Y-%m-%d')}<br>持續時間: {row['持續時間']}天<br>完成百分比: {int(row['完成百分比'] * 100)}%",
                    showlegend=False,
                    name=f"任務 {row['工作項目']}"
                ))

                # 繪製完成百分比條
                completed_duration = row['持續時間'] * row['完成百分比']
                if completed_duration > 0:
                    fig.add_trace(go.Bar(
                        y=[row['工作項目']],
                        x=[timedelta(days=completed_duration)],
                        base=[row['開始日期']],
                        orientation='h',
                        marker=dict(
                            color='white',
                            line=dict(width=0)
                        ),
                        width=0.25,
                        hovertemplate=None,
                        showlegend=False,
                        name=f"完成 {row['工作項目']}"
                    ))

                    # 添加完成百分比文字標籤
                    text_x_pos = row['開始日期'] + timedelta(days=completed_duration)

                    if completed_duration == 0:
                        text_x_pos = row['開始日期']
                    elif completed_duration < row['持續時間'] * 0.15:
                        text_x_pos = row['開始日期'] + timedelta(days=row['持續時間'] * 0.15)

                    fig.add_annotation(
                        x=text_x_pos,
                        y=row['工作項目'],
                        text=f'{int(row["完成百分比"] * 100)}%',
                        showarrow=False,
                        xanchor='left',
                        yanchor='middle',
                        font=dict(size=10, color='black', weight='bold'),
                    )

        # --- 設定 Y 軸 (工作項目順序) ---
        fig.update_yaxes(
            categoryorder='array',
            categoryarray=y_labels_reversed,
            title_text='工作項目',
            tickfont=dict(size=10)
        )

        # --- 設定 X 軸 (月份格式) ---
        fig.update_xaxes(
            title_text='日期',
            tickformat="%Y年<br>%m月",
            gridcolor='lightgray',
            griddash='dash'
        )

        # --- 調整圖表佈局 ---
        max_label_length = max(df['工作項目'].apply(len)) if not df.empty else 10
        dynamic_left_margin = 80 + (max_label_length * 7)
        if dynamic_left_margin > 400:
            dynamic_left_margin = 400

        fig.update_layout(
            title_text='專案甘特圖',
            title_x=0.5,
            height=800,
            margin=dict(l=dynamic_left_margin, r=50, t=80, b=100),
            barmode='overlay',
            hovermode="x unified"
        )

        # --- 在群組 1 和 2 中間新增 X 軸格線 ---
        group_2_start_item_row = df_original[df_original['項次'] == 2]
        if not group_2_start_item_row.empty:
            group_2_start_item_name = group_2_start_item_row['工作項目'].iloc[0]
            group_2_y_index = y_labels_reversed.index(group_2_start_item_name)
            fig.add_hline(y=group_2_y_index - 0.5, line_width=1.5, line_color="gray", line_dash="solid")

        return fig

    except FileNotFoundError:
        print(f"錯誤：找不到檔案 '{file_name}'。請檢查檔案名稱是否正確，並確認它與 Python 腳本在同一資料夾。")
        # 即使發生錯誤，fig 也是一個空的 Figure 物件，可以直接返回
        return fig
    except KeyError as e:
        print(f"錯誤：{e}")
        # 即使發生錯誤，fig 也是一個空的 Figure 物件，可以直接返回
        return fig
    except Exception as e:
        print(f"發生未知錯誤: {e}")
        # 即使發生錯誤，fig 也是一個空的 Figure 物件，可以直接返回
        return fig


# --- Dash 應用程式 ---
app = Dash(__name__)

excel_file_name = 'gantt_data.xlsx'

app.layout = html.Div([
    html.H1("專案甘特圖", style={'textAlign': 'center'}),
    dcc.Graph(
        id='gantt-chart',
        figure=create_gantt_figure(excel_file_name)
    )
])

if __name__ == '__main__':
    # 在運行前，請確保你有 'gantt_data.xlsx' 檔案在相同目錄中
    # 若要運行 Dash 應用程式，請在終端機中執行此 Python 檔案
    # 然後打開瀏覽器，訪問 http://127.0.0.1:8050/
    app.run_server(debug=True)