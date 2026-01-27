import dash
from dash import html, dcc, dash_table
import pandas as pd
import plotly.express as px
from datetime import timedelta
import numpy as np

# ==========================================
# 1. 資料讀取與重組
# ==========================================

source_file = '布洛灣復建工程甘特圖(第二版)_20260105.xlsx'
try:
    df_raw = pd.read_excel(source_file, sheet_name='布洛灣')
except FileNotFoundError:
    print(f"找不到檔案 {source_file}，請確認檔案位置。")
    df_raw = pd.DataFrame()

formatted_data = []
task_order_list = []

# 遍歷原始資料的每一列
for index, row in df_raw.iterrows():
    task_name = str(row.iloc[0]).strip()

    # 記錄任務順序
    if task_name not in task_order_list:
        task_order_list.append(task_name)

    for i in range(1, len(df_raw.columns), 2):
        if i + 1 >= len(df_raw.columns):
            break
        start_date = row.iloc[i]
        finish_date = row.iloc[i + 1]

        if pd.notna(start_date) and pd.notna(finish_date):
            formatted_data.append({
                'Task': task_name,
                'Start': start_date,
                'Finish': finish_date
            })

df = pd.DataFrame(formatted_data)

# 確保日期格式正確
df['Start'] = pd.to_datetime(df['Start'])
df['Finish'] = pd.to_datetime(df['Finish'])

# 計算持續時間
df['Duration_Days'] = (df['Finish'] - df['Start']).dt.days
df['Duration_Months'] = np.round(df['Duration_Days'] / 30.44, 1)

# ==========================================
# 2. 設定顏色分組邏輯 (新增排水工程組)
# ==========================================

# 定義群組清單
group_slope = [
    '刷坡', '消能式落石防護網', '覆蓋式落石防護網',
    '開口隔幕型落防護網', '自由型框+岩栓'
]

group_structure = [
    '防落石柵', '半重力式擋土牆+防落石柵', '石籠擋土牆',
    '草種噴植', '漿砌卵石擋土牆', '石籠攔砂擋牆',
    '石籠護岸', '路面修復', '護欄修復','半重力式擋土牆',
]

# [新增] 排水工程組
group_drainage = [
    '梳子壩', '型鋼攔石柵', '拍漿溝',
    'RCP涵管Ф=1.5m', '集水井','鋼管壩','RC節制壩','砌石護岸','排水箱涵'
]


# 定義分類函式
def assign_color_group(task):
    # 清理一下任務名稱可能有的空白，避免對應失敗
    t = task.strip()
    if t in group_slope:
        return '邊坡防護工程'
    elif t in group_structure:
        return '擋土與修復工程'
    elif t in group_drainage:
        return '排水工程'
    else:
        return '其他/前置作業'


# 應用分類
df['Color_Group'] = df['Task'].apply(assign_color_group)

# 定義顏色映射 (色碼)
color_map = {
    '邊坡防護工程': '#3498db',  # 亮藍色
    '擋土與修復工程': '#2ecc71',  # 翠綠色
    '排水工程': '#f39c12',  # 橘色 (新增)
    '其他/前置作業': '#95a5a6'  # 灰色
}

# ==========================================
# 3. 建立甘特圖
# ==========================================

dynamic_height = max(400, len(task_order_list) * 35)

fig = px.timeline(
    df,
    x_start="Start",
    x_end="Finish",
    y="Task",
    color="Color_Group",
    color_discrete_map=color_map,
    hover_data={
        "Task": True,
        "Color_Group": True,
        "Start": "|%Y-%m-%d",
        "Finish": "|%Y-%m-%d",
        "Duration_Months": True
    },
    title="布洛灣復建工程進度甘特圖",
    height=dynamic_height
)

today = pd.Timestamp.now().normalize()

if not df.empty:
    min_date = df['Start'].min() - timedelta(days=15)
    max_date = df['Finish'].max() + timedelta(days=60)
else:
    min_date = today - timedelta(days=30)
    max_date = today + timedelta(days=30)

# ==========================================
# 4. 圖表版面與標籤設定
# ==========================================

# 手動反轉任務清單以符合視覺順序
task_order_list_reversed = list(reversed(task_order_list))

fig.update_yaxes(
    title_text="工作項目",
    categoryorder='array',
    categoryarray=task_order_list_reversed,
    showgrid=True,
    gridcolor='#e6e6e6',
    mirror=True
)

# # 添加垂直線
# fig.add_shape(
#     type="line",
#     x0=today, y0=0,
#     x1=today, y1=1,
#     xref="x", yref="paper",
#     line=dict(color="red", width=2, dash="dash")
# )
#
# # 添加「今日」文字標籤
# fig.add_annotation(
#     x=today,
#     y=1,
#     xref="x", yref="paper",
#     text="今日",
#     showarrow=False,
#     xanchor="right",
#     yanchor="bottom",
#     font=dict(color="red", size=12),
#     xshift=-5
# )

# 添加持續時間標籤
for i, row in df.iterrows():
    fig.add_annotation(
        x=row['Finish'],
        y=row['Task'],
        text=f"{row['Duration_Months']}",
        showarrow=False,
        xanchor="left",
        xshift=5,
        font=dict(size=9, color="#555555"),
        bgcolor="rgba(255, 255, 255, 0.4)",
        borderpad=1,
    )

# 全域版面調整
fig.update_layout(
    title_font_size=24,
    title_x=0.5,
    paper_bgcolor='#f8f9fa',
    plot_bgcolor='#ffffff',
    margin=dict(l=150, r=100, t=80, b=50),
    xaxis_range=[min_date, max_date],
    showlegend=True,
    legend_title_text='工程類別',
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    ),
    font=dict(family="Microsoft JhengHei, Arial, sans-serif"),
    bargap=0.4
)

fig.update_xaxes(
    dtick="M1",
    tickformat="%Y-%m",
    tickangle=-45,
    title_text="日期",
    showgrid=True,
    gridcolor='#e6e6e6',
    mirror=True,
    side="bottom"
)

# ==========================================
# 5. Dash App 介面
# ==========================================
app = dash.Dash(__name__, external_stylesheets=[
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css'
])

table_df = df[['Task', 'Start', 'Finish', 'Duration_Days', 'Duration_Months', 'Color_Group']].copy()
table_df['Start'] = table_df['Start'].dt.strftime('%Y-%m-%d')
table_df['Finish'] = table_df['Finish'].dt.strftime('%Y-%m-%d')
table_df.columns = ['工作項目', '開始日期', '結束日期', '持續天數', '持續月數', '工程類別']

app.layout = html.Div([
    html.Div([
        html.H1("布洛灣復建工程進度管理", className="header-title"),
        html.P("工程進度追蹤與管理儀表板", className="header-description")
    ], className="header"),

    html.Div([
        html.Div([
            html.H2("進度甘特圖", className="chart-title"),
            dcc.Graph(
                id='gantt-chart',
                figure=fig,
                config={'displayModeBar': True}
            )
        ], className="chart-container"),

        html.Div([
            html.H2("詳細資訊", className="table-title"),
            dash_table.DataTable(
                id='project-table',
                columns=[{"name": i, "id": i} for i in table_df.columns],
                data=table_df.to_dict('records'),
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'left', 'padding': '8px'},
                style_header={'backgroundColor': '#f0f0f0', 'fontWeight': 'bold'},
                style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': '#f9f9f9'}],
                page_size=15
            )
        ], className="table-container"),
    ], className="main-content"),

    html.Div([html.P("© 2025 布洛灣復建工程管理系統")], className="footer")
], className="dashboard-container")

app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>布洛灣復建工程進度管理</title>
        {%favicon%}
        {%css%}
        <style>
            body { font-family: "Microsoft JhengHei", Arial, sans-serif; margin: 0; padding: 0; background-color: #f8f9fa; }
            .dashboard-container { max-width: 1200px; margin: 0 auto; padding: 20px; }
            .header { background-color: #27ae60; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
            .header-title { margin: 0; font-size: 28px; }
            .header-description { margin: 5px 0 0; opacity: 0.8; }
            .chart-container, .table-container { background-color: white; border-radius: 8px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
            .chart-title, .table-title { color: #333; margin-top: 0; }
            .footer { text-align: center; color: #666; padding: 20px; }
            @media (min-width: 768px) { .main-content { flex-direction: column; } }
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