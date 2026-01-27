from logging import debug
import os
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, callback_context, dash_table, State
import io
from flask import send_file, Response
import glob
import dash_bootstrap_components as dbc


# 定義通用的數據處理函數
def process_task_summary(gantt_df, task_keyword):
    """處理特定工作項目的摘要資料"""
    task_df = gantt_df[gantt_df['Task'].str.contains(task_keyword, na=False)]
    if not task_df.empty:
        # 依「Start」排序，取每個步道最後一筆（調查開始時間最大的那筆）
        summary = task_df.sort_values("Start").groupby("Path").tail(1).reset_index(drop=True)
        summary = summary[["Path", "Start", "Finish"]]
        summary.rename(columns={"Path": "步道", "Start": "調查開始", "Finish": "調查結束"}, inplace=True)

        # 將日期格式化，只顯示到「日」
        summary["調查開始"] = summary["調查開始"].dt.strftime("%Y-%m-%d")
        summary["調查結束"] = summary["調查結束"].dt.strftime("%Y-%m-%d")
    else:
        summary = pd.DataFrame(columns=["步道", "調查開始", "調查結束"])

    return summary, task_df


def load_latest_excel(folder_path):
    """載入最新的Excel檔案"""
    try:
        # 取得資料夾中所有 .xlsx 檔案的路徑
        excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

        if not excel_files:
            return None, "找不到Excel檔案。請確保指定資料夾中有.xlsx檔案。"

        # 找出最新的檔案，依據最後修改時間
        file_path = max(excel_files, key=os.path.getmtime)
        print(f"正在讀取檔案: {file_path}")

        # 讀取 Excel 檔案及處理資料
        xls = pd.ExcelFile(file_path)
        df_all = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            df_all[sheet] = df

        return df_all, None

    except Exception as e:
        return None, f"讀取Excel檔案時發生錯誤: {str(e)}"


def prepare_gantt_data(df_all):
    """從Excel資料準備甘特圖所需的資料"""
    gantt_data = []
    for sheet_name, df in df_all.items():
        for index, row in df.iterrows():
            num_times = int(row['次']) if not pd.isna(row['次']) else 0

            if num_times > 0:
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
            else:
                gantt_data.append({
                    'Task': row['工作項目'],
                    'Path': sheet_name,
                    'Team': row['團隊'],
                    'Start': pd.NaT,
                    'Finish': pd.NaT
                })

    return pd.DataFrame(gantt_data)


def create_gantt_chart(df, title="專案任務甘特圖", height=1000, single_task=False, path_colors=None):
    """創建甘特圖"""
    if df.empty:
        # 返回一個帶有錯誤訊息的空圖表
        fig = px.scatter(title="沒有可用的甘特圖資料")
        fig.update_layout(
            title=dict(text="沒有可用的甘特圖資料", font=dict(size=26, color="red")),
            height=height
        )
        return fig

    if path_colors is None:
        unique_paths = df['Path'].unique()
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

    fig.update_layout(
        title=dict(text=title, font=dict(size=26, color="black", weight="bold")),
        xaxis_title="時間",
        yaxis_title="工項",
        xaxis_title_font=dict(size=20, family="Arial", color="black", weight="bold"),
        yaxis_title_font=dict(size=20, family="Arial", color="black", weight="bold"),
        yaxis=dict(autorange="reversed", tickfont=dict(size=16)),
        xaxis=dict(tickfont=dict(size=16)),
        showlegend=True,
        legend_title="步道",
        height=height
    )
    return fig


def process_uav_model_data(gantt_df):
    """處理UAV測繪與三維模型建模資料"""
    # 篩選 Task 中包含「高精度UAV測繪」或「三維模型建模」的資料
    uav_model_df = gantt_df[
        gantt_df['Task'].str.contains("高精度UAV測繪", na=False) |
        gantt_df['Task'].str.contains("三維模型建模", na=False)
        ]

    # 針對同一步道進行分組處理
    def filter_path(group):
        # 如果該步道中有任一列有有效的開始時間（非空值）
        if group['Start'].notnull().any():
            # 則只保留開始時間不為空的資料
            return group[group['Start'].notnull()]
        else:
            # 否則（即全部都是空值）保留全部資料
            return group

    if not uav_model_df.empty:
        uav_model_df = uav_model_df.groupby("Path", group_keys=False).apply(filter_path)

        # 依開始時間排序，利用 na_position="last" 把空值排到最後
        uav_model_df = uav_model_df.sort_values(by="Start", na_position="last").reset_index(drop=True)

        # 將日期格式化為只顯示「年-月-日」
        uav_model_df["Start"] = uav_model_df["Start"].apply(lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else "")
        uav_model_df["Finish"] = uav_model_df["Finish"].apply(lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else "")

        # 將 "Path" 欄位重新命名為 "步道" 以便於顯示
        uav_model_df.rename(columns={"Path": "步道"}, inplace=True)
        return uav_model_df
    else:
        return pd.DataFrame(columns=["Team", "步道", "Task", "Start", "Finish"])


# 資料載入與處理
def initialize_data():
    """初始化應用程式所需的資料"""
    # 設定相對路徑（假設 Excel 檔案位於 "app/xlsfile" 資料夾內）
    folder_path = "app/xlsfile"

    # 載入Excel資料
    df_all, error_message = load_latest_excel(folder_path)

    if error_message:
        return None, None, None, None, None, None, None, None, None, error_message

    # 準備甘特圖資料
    gantt_df = prepare_gantt_data(df_all)

    # 處理各種工作項目的摘要資料及原始資料
    survey_summary, survey_df = process_task_summary(gantt_df, "現況調查")
    geology_summary, geology_df = process_task_summary(gantt_df, "路線地質")
    qslope_summary, qslope_df = process_task_summary(gantt_df, "岩體評分")

    # 處理UAV測繪與三維模型建模資料
    uav_model_df = process_uav_model_data(gantt_df)

    return gantt_df, survey_summary, geology_summary, qslope_summary, uav_model_df, df_all, survey_df, geology_df, qslope_df, None


# 初始化資料
gantt_df, survey_summary, geology_summary, qslope_summary, uav_model_df, df_all, survey_df, geology_df, qslope_df, error_message = initialize_data()

# 建立顏色映射（確保同步道在不同視圖中有相同顏色）
path_colors = None
if gantt_df is not None:
    unique_paths = gantt_df['Path'].unique()
    color_palette = px.colors.qualitative.Set3
    path_colors = {path: color_palette[i % len(color_palette)] for i, path in enumerate(unique_paths)}

# 建立 Dash 應用程式
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)
server = app.server


# 定義應用程式的佈局
def get_app_layout():
    if error_message:
        return dbc.Container([
            html.H1("錯誤", className="text-danger text-center my-4"),
            html.Div(error_message, className="alert alert-danger"),
            html.Button("重新整理", id="refresh-btn", className="btn btn-primary mt-3"),
            dcc.Store(id="store-data"),
            html.Div(id="status-message"),
            dcc.Dropdown(id='path-dropdown', style={'display': 'none'}),
            html.Button("顯示全部工項", id='show-all', style={'display': 'none'}),
            html.Button("重新載入資料", id="reload-btn", style={'display': 'none'})
        ], fluid=True)

    return dbc.Container([
        html.H1("太魯閣案甘特圖", className="text-center my-4"),

        dbc.Row([
            dbc.Col([
                dbc.Button("顯示全部工項", id='show-all', className="me-2"),
                dcc.Dropdown(
                    id='path-dropdown',
                    options=[{'label': path, 'value': path} for path in
                             gantt_df['Path'].unique()] if gantt_df is not None else [],
                    placeholder="選擇步道",
                    multi=False,
                    className="d-inline-block",
                    style={'width': '300px'}
                ),
                dbc.Button("重新載入資料", id="reload-btn", className="ms-2", color="secondary"),
                html.Button("重新整理", id="refresh-btn", style={'display': 'none'})
            ], width=12, className="d-flex align-items-center mb-3")
        ]),

        dbc.Tabs([
            dbc.Tab(label="主要甘特圖", children=[
                dbc.Row([
                    dbc.Col([
                        dcc.Loading(
                            id="loading-gantt",
                            type="circle",
                            children=dcc.Graph(
                                id='gantt-chart',
                                figure=create_gantt_chart(gantt_df,
                                                          path_colors=path_colors) if gantt_df is not None else {}
                            )
                        )
                    ], width=12)
                ])
            ]),

            dbc.Tab(label="現況調查甘特圖", children=[
                dbc.Row([
                    dbc.Col([
                        dcc.Loading(
                            id="loading-survey-gantt",
                            type="circle",
                            children=dcc.Graph(
                                id='survey-gantt-chart',
                                figure=create_gantt_chart(
                                    survey_df,
                                    title="現況調查工項甘特圖",
                                    path_colors=path_colors,
                                    height=600
                                ) if survey_df is not None else {}
                            )
                        )
                    ], width=12)
                ])
            ]),

            dbc.Tab(label="路線地質甘特圖", children=[
                dbc.Row([
                    dbc.Col([
                        dcc.Loading(
                            id="loading-geology-gantt",
                            type="circle",
                            children=dcc.Graph(
                                id='geology-gantt-chart',
                                figure=create_gantt_chart(
                                    geology_df,
                                    title="路線地質工項甘特圖",
                                    path_colors=path_colors,
                                    height=600
                                ) if geology_df is not None else {}
                            )
                        )
                    ], width=12)
                ])
            ]),

            dbc.Tab(label="岩體評分甘特圖", children=[
                dbc.Row([
                    dbc.Col([
                        dcc.Loading(
                            id="loading-qslope-gantt",
                            type="circle",
                            children=dcc.Graph(
                                id='qslope-gantt-chart',
                                figure=create_gantt_chart(
                                    qslope_df,
                                    title="岩體評分工項甘特圖",
                                    path_colors=path_colors,
                                    height=600
                                ) if qslope_df is not None else {}
                            )
                        )
                    ], width=12)
                ])
            ])
        ], id="tabs", active_tab="tab-0"),

        dbc.Row([
            dbc.Col([
                html.H2("各步道現況調查時間摘要", className="mt-4 mb-2"),
                dash_table.DataTable(
                    id="survey-table",
                    columns=[{"name": col, "id": col} for col in
                             survey_summary.columns] if survey_summary is not None else [],
                    data=survey_summary.to_dict('records') if survey_summary is not None else [],
                    style_table={'overflowX': 'auto'},
                    page_size=10,
                    style_cell={'textAlign': 'center'},
                    style_header={'fontWeight': 'bold'},
                    sort_action="native",
                    filter_action="native"
                )
            ], width=12)
        ]),

        dbc.Row([
            dbc.Col([
                html.H2("各步道路線地質時間摘要", className="mt-4 mb-2"),
                dash_table.DataTable(
                    id="geology-table",
                    columns=[{"name": col, "id": col} for col in
                             geology_summary.columns] if geology_summary is not None else [],
                    data=geology_summary.to_dict('records') if geology_summary is not None else [],
                    style_table={'overflowX': 'auto'},
                    page_size=10,
                    style_cell={'textAlign': 'center'},
                    style_header={'fontWeight': 'bold'},
                    sort_action="native",
                    filter_action="native"
                )
            ], width=12)
        ]),

        dbc.Row([
            dbc.Col([
                html.H2("各步道岩體評分時間摘要", className="mt-4 mb-2"),
                dash_table.DataTable(
                    id="qslope-table",
                    columns=[{"name": col, "id": col} for col in
                             qslope_summary.columns] if qslope_summary is not None else [],
                    data=qslope_summary.to_dict('records') if qslope_summary is not None else [],
                    style_table={'overflowX': 'auto'},
                    page_size=10,
                    style_cell={'textAlign': 'center'},
                    style_header={'fontWeight': 'bold'},
                    sort_action="native",
                    filter_action="native"
                )
            ], width=12)
        ]),

        dbc.Row([
            dbc.Col([
                html.H2("各步道高精度UAV測繪與三維模型建模時間摘要", className="mt-4 mb-2"),
                dash_table.DataTable(
                    id="uav-model-table",
                    columns=[{"name": col, "id": col} for col in
                             ["Team", "步道", "Task", "Start", "Finish"]] if uav_model_df is not None else [],
                    data=uav_model_df[["Team", "步道", "Task", "Start", "Finish"]].to_dict(
                        'records') if uav_model_df is not None else [],
                    style_table={'overflowX': 'auto'},
                    page_size=15,
                    style_cell={'textAlign': 'center'},
                    style_header={'fontWeight': 'bold'},
                    sort_action="native",
                    filter_action="native"
                )
            ], width=12)
        ]),

        dbc.Row([
            dbc.Col([
                html.Div(id="status-message", className="mt-3"),
                dcc.Store(id="store-data")
            ], width=12)
        ])
    ], fluid=True)


app.layout = get_app_layout


# 回調函數，用於更新甘特圖
@app.callback(
    Output('gantt-chart', 'figure'),
    [Input('path-dropdown', 'value'),
     Input('gantt-chart', 'clickData'),
     Input('show-all', 'n_clicks'),
     Input('store-data', 'data')]
)
def update_gantt(selected_path, clickData, n_clicks, stored_data):
    ctx = callback_context

    if gantt_df is None:
        # 返回一個帶有錯誤訊息的空圖表
        fig = px.scatter(title="沒有可用的甘特圖資料")
        fig.update_layout(
            title=dict(text="沒有可用的甘特圖資料", font=dict(size=26, color="red")),
            height=600
        )
        return fig

    if ctx.triggered and ctx.triggered[0]['prop_id'] == 'show-all.n_clicks':
        return create_gantt_chart(gantt_df, path_colors=path_colors, height=1000, single_task=False)

    if selected_path:
        filtered_df = gantt_df[gantt_df['Path'] == selected_path]
        return create_gantt_chart(filtered_df, title=f"步道：{selected_path} 的甘特圖", path_colors=path_colors)

    if clickData is not None:
        clicked_task = clickData['points'][0]['y']
        filtered_df = gantt_df[gantt_df['Task'] == clicked_task]
        title = f"工項：{clicked_task} 的甘特圖"
        return create_gantt_chart(filtered_df, title=title, path_colors=path_colors, height=500, single_task=True)

    return create_gantt_chart(gantt_df, path_colors=path_colors)


# 回調函數，用於根據選擇的步道更新所有甘特圖
@app.callback(
    [Output('survey-gantt-chart', 'figure'),
     Output('geology-gantt-chart', 'figure'),
     Output('qslope-gantt-chart', 'figure')],
    [Input('path-dropdown', 'value'),
     Input('store-data', 'data')]
)
def update_task_specific_gantt(selected_path, stored_data):
    # 處理無數據情況
    if survey_df is None or geology_df is None or qslope_df is None:
        empty_fig = px.scatter(title="沒有可用的資料")
        empty_fig.update_layout(height=600)
        return empty_fig, empty_fig, empty_fig

    # 為每個工作類型創建甘特圖
    if selected_path:
        # 篩選現況調查數據
        filtered_survey = survey_df[survey_df['Path'] == selected_path]
        survey_fig = create_gantt_chart(
            filtered_survey,
            title=f"步道：{selected_path} 的現況調查甘特圖",
            path_colors=path_colors,
            height=600
        )

        # 篩選路線地質數據
        filtered_geology = geology_df[geology_df['Path'] == selected_path]
        geology_fig = create_gantt_chart(
            filtered_geology,
            title=f"步道：{selected_path} 的路線地質甘特圖",
            path_colors=path_colors,
            height=600
        )

        # 篩選岩體評分數據
        filtered_qslope = qslope_df[qslope_df['Path'] == selected_path]
        qslope_fig = create_gantt_chart(
            filtered_qslope,
            title=f"步道：{selected_path} 的岩體評分甘特圖",
            path_colors=path_colors,
            height=600
        )
    else:
        # 不篩選，顯示所有數據
        survey_fig = create_gantt_chart(
            survey_df,
            title="現況調查工項甘特圖",
            path_colors=path_colors,
            height=600
        )

        geology_fig = create_gantt_chart(
            geology_df,
            title="路線地質工項甘特圖",
            path_colors=path_colors,
            height=600
        )

        qslope_fig = create_gantt_chart(
            qslope_df,
            title="岩體評分工項甘特圖",
            path_colors=path_colors,
            height=600
        )

    return survey_fig, geology_fig, qslope_fig


# 回調函數，用於重新載入資料
@app.callback(
    [Output('store-data', 'data'),
     Output('status-message', 'children'),
     Output('path-dropdown', 'options')],
    [Input('reload-btn', 'n_clicks'),
     Input('refresh-btn', 'n_clicks')],
    prevent_initial_call=True
)
def reload_data(n_clicks, refresh_clicks):
    from dash import no_update
    import dash

    if n_clicks is None and refresh_clicks is None:
        raise dash.exceptions.PreventUpdate

    # 重新初始化資料
    global gantt_df, survey_summary, geology_summary, qslope_summary, uav_model_df, df_all, path_colors, error_message
    global survey_df, geology_df, qslope_df
    gantt_df, survey_summary, geology_summary, qslope_summary, uav_model_df, df_all, survey_df, geology_df, qslope_df, error_message = initialize_data()

    if error_message:
        return {}, dbc.Alert(error_message, color="danger"), []

    # 更新顏色映射
    unique_paths = gantt_df['Path'].unique()
    color_palette = px.colors.qualitative.Set3
    path_colors = {path: color_palette[i % len(color_palette)] for i, path in enumerate(unique_paths)}

    # 更新步道下拉選單選項
    dropdown_options = [{'label': path, 'value': path} for path in gantt_df['Path'].unique()]

    message = dbc.Alert("資料已成功重新載入！", color="success", duration=4000)
    return {"timestamp": pd.Timestamp.now().isoformat()}, message, dropdown_options


# 主程式
if __name__ == '__main__':
    app.run(debug=True)