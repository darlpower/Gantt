from logging import debug
import os
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, callback_context, dash_table, State
import glob
import dash_bootstrap_components as dbc


# 定義通用的數據處理函數
def process_task_summary(gantt_df, task_keyword):
    """處理特定工作項目的摘要資料"""
    task_df = gantt_df[gantt_df['Task'].str.contains(task_keyword, na=False)]
    if not task_df.empty:
        summary = task_df.sort_values("Start").groupby("Path").tail(1).reset_index(drop=True)
        summary = summary[["Path", "Start", "Finish"]]
        summary.rename(columns={"Path": "步道", "Start": "調查開始", "Finish": "調查結束"}, inplace=True)
        summary["調查開始"] = summary["調查開始"].dt.strftime("%Y-%m-%d")
        summary["調查結束"] = summary["調查結束"].dt.strftime("%Y-%m-%d")
    else:
        summary = pd.DataFrame(columns=["步道", "調查開始", "調查結束"])
    return summary, task_df


def load_latest_excel(folder_path):
    """載入最新的Excel檔案"""
    try:
        excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))
        if not excel_files:
            return None, "找不到Excel檔案。請確保指定資料夾中有.xlsx檔案。"
        file_path = max(excel_files, key=os.path.getmtime)
        print(f"正在讀取檔案: {file_path}")
        xls = pd.ExcelFile(file_path)
        df_all = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
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


def create_gantt_chart(df, title="專案任務甘特圖", height=1000, path_colors=None):
    """創建甘特圖，Y 軸改為步道"""
    if df.empty:
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
        y="Path",  # 修改為 Path（步道）
        color="Path",
        text="Team",
        hover_data=["Team"],
        title=title,
        color_discrete_map=path_colors
    )

    fig.update_layout(
        title=dict(text=title, font=dict(size=26, color="black", weight="bold")),
        xaxis_title="時間",
        yaxis_title="步道",  # 更新 Y 軸標題
        xaxis_title_font=dict(size=20, family="Arial", color="black", weight="bold"),
        yaxis_title_font=dict(size=20, family="Arial", color="black", weight="bold"),
        yaxis=dict(autorange="reversed", tickfont=dict(size=16)),
        xaxis=dict(tickfont=dict(size=16)),
        showlegend=True,
        legend_title="步道",
        height=height
    )
    # 新增格線設定
    fig.update_xaxes(showgrid=True, gridcolor="lightgray", gridwidth=1)
    fig.update_yaxes(showgrid=True, gridcolor="lightgray", gridwidth=1)
    return fig


# 初始化資料
def initialize_data():
    folder_path = "app/xlsfile"
    df_all, error_message = load_latest_excel(folder_path)
    if error_message:
        return None, None, None, None, None, error_message
    gantt_df = prepare_gantt_data(df_all)
    survey_summary, survey_df = process_task_summary(gantt_df, "現況調查")
    geology_summary, geology_df = process_task_summary(gantt_df, "路線地質")
    qslope_summary, qslope_df = process_task_summary(gantt_df, "岩體評分")
    return gantt_df, survey_summary, geology_summary, qslope_summary, survey_df, geology_df, qslope_df


gantt_df, survey_summary, geology_summary, qslope_summary, survey_df, geology_df, qslope_df = initialize_data()

# 建立 Dash 應用程式
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)

# 定義應用程式佈局
app.layout = dbc.Container([
    html.H1("太魯閣案甘特圖", className="text-center my-4"),

    dbc.Tabs([
        dbc.Tab(label="現況調查甘特圖", children=[
            dcc.Graph(
                id='survey-gantt-chart',
                figure=create_gantt_chart(
                    survey_df,
                    title="現況調查工項甘特圖",
                    height=600
                ) if survey_df is not None else {}
            )
        ]),

        dbc.Tab(label="路線地質甘特圖", children=[
            dcc.Graph(
                id='geology-gantt-chart',
                figure=create_gantt_chart(
                    geology_df,
                    title="路線地質工項甘特圖",
                    height=600
                ) if geology_df is not None else {}
            )
        ]),

        dbc.Tab(label="岩體評分甘特圖", children=[
            dcc.Graph(
                id='qslope-gantt-chart',
                figure=create_gantt_chart(
                    qslope_df,
                    title="岩體評分工項甘特圖",
                    height=600
                ) if qslope_df is not None else {}
            )
        ])
    ])
], fluid=True)

if __name__ == "__main__":
    app.run(debug=True)
