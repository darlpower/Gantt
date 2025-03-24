from logging import debug
import os
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, callback_context, dash_table
import io
from flask import send_file

# 讀取 Excel 檔案及處理資料
try:
    xls = pd.ExcelFile(r"X:\63882\gantt\工作進度安排_HHL_250214.xlsx")
    df_all = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df_all[sheet] = df
except Exception as e:
    print(f"讀取 Excel 檔案錯誤：{e}")
    exit()

# 依照每個工作表、每個工項及時間段數量，建立甘特圖所需資料
gantt_data = []
for sheet_name, df in df_all.items():
    for index, row in df.iterrows():
        num_times = int(row['次'])
        if pd.isna(num_times):
            num_times = 0  # 若為 NaN，視為 0

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

gantt_df = pd.DataFrame(gantt_data)

# 篩選出「現況調查」的資料
survey_df = gantt_df[gantt_df['Task'].str.contains("現況調查", na=False)]
if not survey_df.empty:
    # 依「Start」排序，取每個步道最後一筆（調查開始時間最大的那筆）
    survey_summary = survey_df.sort_values("Start").groupby("Path").tail(1).reset_index(drop=True)
    survey_summary = survey_summary[["Path", "Start", "Finish"]]
    survey_summary.rename(columns={"Path": "步道","Start": "調查開始", "Finish": "調查結束"}, inplace=True)

    # 將日期格式化，只顯示到「日」
    survey_summary["調查開始"] = survey_summary["調查開始"].dt.strftime("%Y-%m-%d")
    survey_summary["調查結束"] = survey_summary["調查結束"].dt.strftime("%Y-%m-%d")

else:
    survey_summary = pd.DataFrame(columns=["Path", "調查開始", "調查結束"])

# 篩選出「路線地質」的資料
geology_df = gantt_df[gantt_df['Task'].str.contains("路線地質", na=False)]
if not geology_df.empty:
    # 依「Start」排序，取每個步道最後一筆（調查開始時間最大的那筆）
    geology_summary = geology_df.sort_values("Start").groupby("Path").tail(1).reset_index(drop=True)
    geology_summary = geology_summary[["Path", "Start", "Finish"]]
    geology_summary.rename(columns={"Path": "步道","Start": "調查開始", "Finish": "調查結束"}, inplace=True)

    # 將日期格式化，只顯示到「日」
    geology_summary["調查開始"] = geology_summary["調查開始"].dt.strftime("%Y-%m-%d")
    geology_summary["調查結束"] = geology_summary["調查結束"].dt.strftime("%Y-%m-%d")

else:
    geology_summary = pd.DataFrame(columns=["Path", "調查開始", "調查結束"])

# 篩選出「岩體評分」的資料
qslope_df = gantt_df[gantt_df['Task'].str.contains("岩體評分", na=False)]
if not qslope_df.empty:
    # 依「Start」排序，取每個步道最後一筆（調查開始時間最大的那筆）
    qslope_summary = qslope_df.sort_values("Start").groupby("Path").tail(1).reset_index(drop=True)
    qslope_summary = qslope_summary[["Path", "Start", "Finish"]]
    qslope_summary.rename(columns={"Path": "步道","Start": "調查開始", "Finish": "調查結束"}, inplace=True)

    # 將日期格式化，只顯示到「日」
    qslope_summary["調查開始"] = qslope_summary["調查開始"].dt.strftime("%Y-%m-%d")
    qslope_summary["調查結束"] = qslope_summary["調查結束"].dt.strftime("%Y-%m-%d")

else:
    qslope_summary = pd.DataFrame(columns=["Path", "調查開始", "調查結束"])


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

    fig.update_layout(
        title=dict(text=title, font=dict(size=26, color="black", weight="bold")),
        xaxis_title="時間",
        yaxis_title="工項",
        xaxis_title_font=dict(size=20, family="Arial", color="black", weight="bold"),
        yaxis_title_font=dict(size=20, family="Arial", color="black", weight="bold"),
        yaxis=dict(autorange="reversed", tickfont=dict(size=16)),
        xaxis=dict(tickfont=dict(size=16)),
        showlegend=True, legend_title="步道",
        height=height
    )
    return fig


# 建立 Dash 應用程式
app = Dash(__name__)
server = app.server

app.layout = html.Div([
    html.H1("太魯閣案甘特圖", style={'fontSize': '32px', 'textAlign': 'center'}),
    html.Div([
        html.Button("顯示全部工項", id='show-all', n_clicks=0, style={'margin': '10px', 'fontSize': '18px'}),
        dcc.Dropdown(
            id='path-dropdown',
            options=[{'label': path, 'value': path} for path in gantt_df['Path'].unique()],
            placeholder="選擇步道",
            multi=False,
            style={'width': '300px', 'margin-left': '10px', 'fontSize': '18px'}
        )
    ], style={'display': 'flex', 'align-items': 'center'}),
    dcc.Graph(id='gantt-chart', figure=create_gantt_chart(gantt_df)),
    html.H2("各步道現況調查時間摘要", style={'fontSize': '28px'}),
    dash_table.DataTable(
        id="survey-table",
        columns=[{"name": col, "id": col} for col in survey_summary.columns],
        data=survey_summary.to_dict('records'),
        style_table={'overflowX': 'auto'},
        page_size=29,
        style_cell={'textAlign': 'center', 'fontSize': '21px'}
    ),

    html.H2("各步道路線地質時間摘要", style={'fontSize': '28px'}),
    dash_table.DataTable(
        id="geology-table",
        columns=[{"name": col, "id": col} for col in geology_summary.columns],
        data=geology_summary.to_dict('records'),
        style_table={'overflowX': 'auto'},
        page_size=29,
        style_cell={'textAlign': 'center', 'fontSize': '21px'}
    ),

    html.H2("各步道岩體評分時間摘要", style={'fontSize': '28px'}),
    dash_table.DataTable(
        id="qslope-table",
        columns=[{"name": col, "id": col} for col in qslope_summary.columns],
        data=qslope_summary.to_dict('records'),
        style_table={'overflowX': 'auto'},
        page_size=29,
        style_cell={'textAlign': 'center', 'fontSize': '21px'}
    )

])


@app.callback(
    Output('gantt-chart', 'figure'),
    [Input('path-dropdown', 'value')],
    [Input('gantt-chart', 'clickData'),
     Input('show-all', 'n_clicks')]
)
def update_gantt(selected_path, clickData, n_clicks):
    ctx = callback_context

    if ctx.triggered and ctx.triggered[0]['prop_id'] == 'show-all.n_clicks':
        return create_gantt_chart(gantt_df, height=1000, single_task=False)

    if selected_path:
        filtered_df = gantt_df[gantt_df['Path'] == selected_path]
        return create_gantt_chart(filtered_df, title=f"步道：{selected_path} 的甘特圖")

    if clickData is not None:
        clicked_task = clickData['points'][0]['y']
        filtered_df = gantt_df[gantt_df['Task'] == clicked_task]
        title = f"工項：{clicked_task} 的甘特圖"
        return create_gantt_chart(filtered_df, title=title, height=500, single_task=True)

    return create_gantt_chart(gantt_df)


if __name__ == '__main__':
    app.run_server(debug=True)
