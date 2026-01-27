import dash
from dash import html, dcc, dash_table
import pandas as pd
import plotly.figure_factory as ff
import plotly.colors as pc
from datetime import timedelta
import numpy as np

# 讀取 Excel 文件
df = pd.read_excel('長春祠復建工程甘特圖.xlsx', sheet_name=0)


# 轉換日期欄位
df['Start'] = pd.to_datetime(df['Start'])
df['Finish'] = pd.to_datetime(df['Finish'])

# 計算每個任務的持續時間（天數）
df['Duration_Days'] = (df['Finish'] - df['Start']).dt.days

# 計算持續時間（月）- 更精確的計算
df['Duration_Months'] = np.round(df['Duration_Days'] / 30.44, 1)  # 使用年平均每月天數

# 取得唯一任務種類
unique_tasks = df['Task'].unique()
num_colors = len(unique_tasks)

# 使用 Plotly 預設色盤
colors = pc.qualitative.Plotly
# 如果需要更多顏色，循環使用色盤
color_list = [colors[i % len(colors)] for i in range(num_colors)]

# 建立任務與顏色的對應字典
color_map = {task: color_list[i] for i, task in enumerate(unique_tasks)}

# 為每個任務分配顏色
df['Color'] = df['Task'].map(color_map)

# 建立甘特圖
fig = ff.create_gantt(
    df,
    index_col='Task',
    colors=df['Color'],
    show_colorbar=False,  # 不顯示顏色條
    group_tasks=True,
    showgrid_x=True,
    showgrid_y=True
)

# 獲取今天的日期以標記當前進度
today = pd.Timestamp.now().normalize()

# 添加當前日期的垂直線
fig.add_shape(
    type="line",
    x0=today,
    y0=0,
    x1=today,
    y1=len(df),
    line=dict(color="red", width=2, dash="dash"),
)

# 添加當前日期的標籤
fig.add_annotation(
    x=today,
    y=len(df),
    text="今日",
    showarrow=True,
    arrowhead=1,
    ax=0,
    ay=-40,
    font=dict(size=12, color="red"),
)

# 計算項目總時間範圍，擴展圖表顯示範圍
min_date = df['Start'].min() - timedelta(days=7)
max_date = df['Finish'].max() + timedelta(days=30)  # 增加右側空間以容納標籤

# 改進圖表版面設定
fig.update_layout(
    title={
        'text': "長春祠復建工程進度甘特圖",
        'font': {'size': 24, 'family': 'Arial, sans-serif', 'color': '#333333'},
        'x': 0.5,
        'xanchor': 'center'
    },
    paper_bgcolor='#f8f9fa',
    plot_bgcolor='#ffffff',
    margin=dict(l=150, r=120, t=80, b=50),  # 增加左側和右側邊距
    height=max(500, len(df) * 30),  # 動態調整高度
    legend_title="任務類型",
    font=dict(family="Arial, sans-serif", size=12),
    xaxis_range=[min_date, max_date],  # 設定 x 軸範圍
    hovermode="closest"
)

# 美化 x 軸，設定 dtick 為 "M1"（每月一個刻度）
fig.update_xaxes(
    dtick="M1",
    title_text="日期",
    title_font=dict(size=14, family="Arial, sans-serif"),
    showgrid=True,
    gridcolor='#e6e6e6',
    gridwidth=1,
    tickformat="%Y-%m",  # 顯示格式：西元年份-月份
    tickfont=dict(size=12),
    tickangle=-45,
    showline=True,
    linewidth=2,
    linecolor='#999999',
    mirror=True
)

# 美化 y 軸，設定任務標籤左對齊
fig.update_yaxes(
    title_text="工作項目",
    title_font=dict(size=14, family="Arial, sans-serif"),
    showgrid=True,
    gridcolor='#e6e6e6',
    gridwidth=1,
    tickfont=dict(size=12),
    showline=True,
    linewidth=2,
    linecolor='#999999',
    mirror=True,
    automargin=True,  # 自動調整邊距
    ticksuffix="  "  # 添加後綴空格，使文字與軸線保持一定距離
)

# 獲取工作項目的正確 y 值
# 當使用 group_tasks=True 時，Plotly 會按照任務在 dataframe 中的順序排列
task_positions = {}
y_position = 12
for i, task in enumerate(unique_tasks):
    task_positions[task] = i
    print(f"i等於{i}")
    # task_positions = len(unique_tasks) - 1
    task_df = df[df['Task'] == task]
    # 計算該工作項目的最早開始與最晚結束日期
    task_start = task_df['Start'].min()
    task_finish = task_df['Finish'].max()
    # 計算持續天數，並換算成月（以平均每月30.44天計算）
    duration_days = (task_finish - task_start).days
    duration_months = np.round(duration_days / 30.44, 1)

# 針對每個唯一工作項目添加持續時間標籤
# for task in unique_tasks:
#     # 取得該工作項目的所有記錄
#     task_df = df[df['Task'] == task]
#     # 計算該工作項目的最早開始與最晚結束日期
#     task_start = task_df['Start'].min()
#     task_finish = task_df['Finish'].max()
#     # 計算持續天數，並換算成月（以平均每月30.44天計算）
#     duration_days = (task_finish - task_start).days
#     duration_months = np.round(duration_days / 30.44, 1)

    # 獲取該任務在 y 軸上的位置
    y_position = y_position -1
    print(f"y_postion等於{y_position}")
    # y_position = task_positions - 1
    # 添加持續時間標籤
    fig.add_annotation(
        x=task_finish,
        y=y_position,  # 使用正確的 y 位置
        text=f"{duration_months} 個月",
        showarrow=False,
        xanchor="left",
        yanchor="middle",
        xshift=5,  # 向右移動一點
        font=dict(size=10, color="#333333"),
        bgcolor="rgba(255, 255, 255, 0.7)",
        bordercolor="rgba(0, 0, 0, 0.1)",
        borderwidth=1,
        borderpad=4,
        opacity=0.8
    )
    # fig.update_yaxes(autorange="reversed")

# 建立 Dash App
app = dash.Dash(__name__, external_stylesheets=[
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css'
])

# 建立數據表格顯示詳細信息
table_df = df[['Task', 'Start', 'Finish', 'Duration_Days', 'Duration_Months']].copy()
table_df['Start'] = table_df['Start'].dt.strftime('%Y-%m-%d')
table_df['Finish'] = table_df['Finish'].dt.strftime('%Y-%m-%d')
table_df.columns = ['工作項目', '開始日期', '結束日期', '持續天數', '持續月數']

# 應用程式布局
app.layout = html.Div([
    # 標題區塊
    html.Div([
        html.H1("長春祠復建工程進度管理", className="header-title"),
        html.P("工程進度追蹤與管理儀表板", className="header-description")
    ], className="header"),

    # 甘特圖與表格區塊
    html.Div([
        # 甘特圖
        html.Div([
            html.H2("進度甘特圖", className="chart-title"),
            html.Div([
                html.I(className="fas fa-calendar-alt"),
                html.Span(" 顯示各工作項目的時間安排與持續月數")
            ], className="chart-description"),
            dcc.Graph(
                id='gantt-chart',
                figure=fig,
                config={
                    'displayModeBar': True,
                    'modeBarButtonsToRemove': ['lasso2d', 'select2d'],
                    'displaylogo': False,
                    'toImageButtonOptions': {
                        'format': 'png',
                        'filename': '長春祠復建工程甘特圖',
                        'scale': 2
                    }
                }
            )
        ], className="chart-container"),

        # 表格
        html.Div([
            html.H2("工作項目詳細資訊", className="table-title"),
            dash_table.DataTable(
                id='project-table',
                columns=[{"name": i, "id": i} for i in table_df.columns],
                data=table_df.to_dict('records'),
                style_table={'overflowX': 'auto'},
                style_cell={
                    'textAlign': 'left',
                    'padding': '10px',
                    'fontFamily': 'Arial, sans-serif'
                },
                style_header={
                    'backgroundColor': '#f0f0f0',
                    'fontWeight': 'bold',
                    'border': '1px solid #ddd'
                },
                style_data={
                    'border': '1px solid #ddd'
                },
                style_data_conditional=[
                    {
                        'if': {'row_index': 'odd'},
                        'backgroundColor': '#f9f9f9'
                    }
                ]
            )
        ], className="table-container"),
    ], className="main-content"),

    # 頁尾
    html.Div([
        html.P("© 2025 長春祠復建工程管理系統")
    ], className="footer")
], className="dashboard-container")

# 添加 CSS 樣式
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>長春祠復建工程進度管理</title>
        {%favicon%}
        {%css%}
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 0;
                background-color: #f8f9fa;
            }
            .dashboard-container {
                max-width: 1200px;
                margin: 0 auto;
                padding: 20px;
            }
            .header {
                background-color: #3c6382;
                color: white;
                padding: 20px;
                border-radius: 8px;
                margin-bottom: 20px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            }
            .header-title {
                margin: 0;
                font-size: 28px;
            }
            .header-description {
                margin: 5px 0 0;
                font-size: 16px;
                opacity: 0.8;
            }
            .main-content {
                display: flex;
                flex-direction: column;
                gap: 20px;
            }
            .chart-container, .table-container, .summary-container {
                background-color: white;
                border-radius: 8px;
                padding: 20px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            .chart-title, .table-title, .summary-title {
                color: #333;
                margin-top: 0;
                margin-bottom: 10px;
                font-size: 20px;
            }
            .chart-description {
                color: #666;
                margin-bottom: 15px;
                font-size: 14px;
            }
            .footer {
                text-align: center;
                padding: 20px;
                color: #666;
                font-size: 14px;
                margin-top: 20px;
            }
            .stat-grid {
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
                gap: 15px;
                margin-top: 15px;
            }
            .stat-box {
                background-color: #f8f9fa;
                border-radius: 6px;
                padding: 15px;
                text-align: center;
                border-left: 4px solid #3c6382;
            }
            .stat-title {
                margin: 0;
                font-size: 14px;
                color: #666;
                font-weight: normal;
            }
            .stat-value {
                margin: 10px 0 0;
                font-size: 18px;
                color: #333;
                font-weight: bold;
            }
            .stat-detail {
                display: block;
                font-size: 12px;
                color: #666;
                margin-top: 5px;
            }
            @media (min-width: 768px) {
                .main-content {
                    flex-direction: column;
                }
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''

if __name__ == '__main__':
    app.run(debug=True)