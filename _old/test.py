import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, callback_context

# 讀取 Excel 檔案及處理資料
xls = pd.ExcelFile(r"/工作進度安排_HHL_250214.xlsx")
df_all = {}
for sheet in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet)
    df_all[sheet] = df

# 依照每個工作表、每個工項及時間段數量，建立甘特圖所需資料
gantt_data = []
for sheet_name, df in df_all.items():
    for index, row in df.iterrows():
        num_times = int(row['次']) if not pd.isna(row['次']) else 0
        for i in range(1, num_times + 1):
            start_col = f'起始{i}'
            finish_col = f'結束{i}'
            if pd.notna(row[start_col]) and pd.notna(row[finish_col]):
                gantt_data.append({
                    'Task': row['工作項目'],
                    'Path': sheet_name,
                    'Team': row['團隊'],
                    'Start': pd.to_datetime(row[start_col]),
                    'Finish': pd.to_datetime(row[finish_col])
                })

gantt_df = pd.DataFrame(gantt_data)


# 初始甘特圖，呈現所有工項資料
def create_gantt_chart(df, title="專案任務甘特圖", height=1000, single_task=False):
    unique_paths = gantt_df['Path'].unique()
    color_palette = px.colors.qualitative.Set3
    path_colors = {path: color_palette[i % len(color_palette)] for i, path in enumerate(unique_paths)}

    fig = px.timeline(
        df,
        x_start="Start",
        x_end="Finish",
        y="Task",
        color="Path",
        text="Team",
        hover_data=["Team"],
        title=title,
        color_discrete_map=path_colors
    )

    if single_task:
        fig.update_yaxes(domain=[0.4, 0.6])
    else:
        fig.update_yaxes(autorange="reversed")

    fig.update_layout(
        title=dict(text=title, font=dict(size=24)),
        height=height
    )

    return fig


# 建立 Dash 應用程式
app = Dash(__name__)

app.layout = html.Div([
    html.H1("太魯閣案甘特圖"),
    html.Div([
        dcc.Dropdown(
            id='path-dropdown',
            options=[{'label': path, 'value': path} for path in gantt_df['Path'].unique()],
            placeholder="選擇步道",
            multi=False,
            style={'width': '300px', 'margin-left': '10px'}
        ),
        html.Button("顯示全部工項", id='show-all', n_clicks=0, style={'margin': '10px'}),
        html.Button("重置步道", id='reset-path', n_clicks=0, style={'margin': '10px'})
    ], style={'display': 'flex', 'align-items': 'center'}),

    dcc.Graph(id='gantt-chart', figure=create_gantt_chart(gantt_df))
])


# 建立回呼函式，監聽圖表點擊與按鈕事件
@app.callback(
    Output('gantt-chart', 'figure'),
    [Input('path-dropdown', 'value')],
    [Input('gantt-chart', 'clickData'),
     Input('show-all', 'n_clicks'),
     Input('reset-path', 'n_clicks')]
)
def update_gantt(selected_path, clickData, show_all_clicks, reset_clicks):
    ctx = callback_context

    # 如果按下「顯示全部工項」按鈕，就回復全部資料
    if ctx.triggered and ctx.triggered[0]['prop_id'] == 'show-all.n_clicks':
        return create_gantt_chart(gantt_df)

    # 如果按下「重置步道」按鈕，則清空選擇的步道
    if ctx.triggered and ctx.triggered[0]['prop_id'] == 'reset-path.n_clicks':
        return create_gantt_chart(gantt_df)  # 可以回傳完整資料或其他邏輯

    # 如果選擇了「步道」
    if selected_path:
        filtered_df = gantt_df[gantt_df['Path'] == selected_path]
        return create_gantt_chart(filtered_df, title=f"步道：{selected_path} 的甘特圖")

    # 如果有點擊事件，取得被點擊的工項名稱
    if clickData is not None:
        clicked_task = clickData['points'][0]['y']
        filtered_df = gantt_df[gantt_df['Task'] == clicked_task]
        return create_gantt_chart(filtered_df, title=f"工項：{clicked_task} 的甘特圖", height=500, single_task=True)

    # 若無任何點擊，則回傳初始圖表
    return create_gantt_chart(gantt_df)


if __name__ == '__main__':
    app.run_server(debug=True)
