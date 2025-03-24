from logging import debug
import os
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, callback_context, dash_table
import io
import zipfile
import logging

# 設定日誌
logging.basicConfig(level=logging.INFO)


# 讀取 Excel 檔案
def read_excel_file(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        return xls
    except Exception as e:
        logging.error(f"Failed to read Excel file: {e}")
        return None


# 處理資料
def process_data(xls):
    df_all = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df_all[sheet] = df
    return df_all


# 建立甘特圖所需資料
def create_gantt_data(df_all):
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


# 篩選出「現況調查」的資料
def filter_survey_data(gantt_df):
    survey_df = gantt_df[gantt_df['Task'].str.contains("現況調查", na=False)]
    if not survey_df.empty:
        survey_summary = survey_df.sort_values("Start").groupby("Path").tail(1).reset_index(drop=True)
        survey_summary = survey_summary[["Path", "Start", "Finish"]]
        survey_summary.rename(columns={"Start": "調查開始", "Finish": "調查結束"}, inplace=True)

        survey_summary["調查開始"] = survey_summary["調查開始"].dt.strftime("%Y-%m-%d")
        survey_summary["調查結束"] = survey_summary["調查結束"].dt.strftime("%Y-%m-%d")
    else:
        survey_summary = pd.DataFrame(columns=["Path", "調查開始", "調查結束"])
    return survey_summary


# 生成甘特圖
def create_gantt_chart(df, title="專案任務甘特圖", height=1000, single_task=False):
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

    if single_task:
        fig.update_yaxes(domain=[0.4, 0.6])
    else:
        fig.update_yaxes(autorange="reversed")

    fig.update_traces(textposition="inside", insidetextanchor="middle")
    fig.update_layout(
        title=dict(text=title, font=dict(size=24, color="black", weight="bold")),
        xaxis_title="時間",
        yaxis_title="工項",
        xaxis_title_font=dict(size=18, family="Arial", color="black", weight="bold"),
        yaxis_title_font=dict(size=18, family="Arial", color="black", weight="bold"),
        yaxis=dict(autorange="reversed"),
        showlegend=True, legend_title="步道",
        height=height
    )
    fig.update_xaxes(range=["2024-10-24", "2025-12-31"], dtick="M1", tickfont=dict(size=14))
    fig.update_yaxes(tickfont=dict(size=14), tickangle=0)
    fig.update_xaxes(showline=True, linewidth=2, linecolor='black', mirror=True,
                     showgrid=True, gridwidth=1, gridcolor='black')
    fig.update_yaxes(showline=True, linewidth=2, linecolor='black', mirror=True)
    return fig


# 建立 Dash 應用程式
app = Dash(__name__)
server = app.server

# 讀取 Excel 檔案及處理資料
xls_file_path = r"X:\63882\gantt\工作進度安排_HHL_250214.xlsx"
xls = read_excel_file(xls_file_path)
if xls is not None:
    df_all = process_data(xls)
    gantt_df = create_gantt_data(df_all)
    survey_summary = filter_survey_data(gantt_df)
else:
    logging.error("Failed to load Excel file.")
    exit()

app.layout = html.Div([
    html.H1("太魯閣案甘特圖"),
    html.Div([
        html.Button("顯示全部工項", id='show-all', n_clicks=0, style={'margin': '10px'}),
        dcc.Dropdown(
            id='path-dropdown',
            options=[{'label': path, 'value': path} for path in gantt_df['Path'].unique()],
            placeholder="選擇步道",
            multi=False,
            style={'width': '300px', 'marginleft': '10px'}
        ),
        html.Button("下載所有步道 PDF (ZIP)", id="download-pdf-zip", n_clicks=0, style={'marginleft': '10px'})
    ], style={'display': 'flex', 'align-items': 'center'}),
    dcc.Graph(id='gantt-chart', figure=create_gantt_chart(gantt_df)),

    html.H2("各步道現況調查時間摘要"),
    dash_table.DataTable(
        id="survey-table",
        columns=[{"name": col, "id": col} for col in survey_summary.columns],
        data=survey_summary.to_dict('records'),
        style_table={'overflowX': 'auto'},
        page_size=10,
        style_cell={
            'textAlign': 'center',
            'fontSize': '18px'  # 調整單元格文字大小
        },
        style_header={
            'fontSize': '20px',  # 調整表頭文字大小
            'fontWeight': 'bold'
        }
    )
    ,
    dcc.Download(id="download-zip-component")
])


# 更新甘特圖的回呼函式
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


# 下載所有步道的 PDF (ZIP) 回呼函式
@app.callback(
    Output("download-zip-component", "data"),
    Input("download-pdf-zip", "n_clicks"),
    prevent_initial_call=True
)
def download_all_paths_as_zip(n_clicks):
    try:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for path in gantt_df['Path'].unique():
                filtered_df = gantt_df[gantt_df['Path'] == path]
                fig = create_gantt_chart(filtered_df, title=f"{path} 甘特圖", height=800)
                pdf_bytes = fig.to_image(format="pdf")
                zip_file.writestr(f"{path}.pdf", pdf_bytes)
        zip_buffer.seek(0)
        return dcc.send_bytes(lambda b: b.write(zip_buffer.getvalue()), "Gantt_Charts.zip")
    except Exception as e:
        logging.error(f"Error in download_all_paths_as_zip: {e}")
        return None


if __name__ == '__main__':
    app.run_server(debug=True)
