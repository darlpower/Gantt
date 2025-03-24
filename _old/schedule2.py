import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.graph_objects as go

# read Excel 
xls = pd.ExcelFile(r"/工作進度安排_HHL_250214.xlsx")

# 讀取每個worksheets
df_all = {}
for sheet in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet)
    df_all[sheet] = df
n_sheet=len(df_all)     # count of worksheets

# Data Processing
gantt_data = []     # 存放繪製各步道甘特圖資料
date_se=[]          # 存放繪製總工特圖資料

# 歷遍每個DataFrame
for sheet_name, df in df_all.items():   
    start_times = []  #存放每個DataFrame的所有起始時間
    end_times = []    #存放每個DataFrame的所有結束時間
    
    # 歷遍DataFrame中的每列
    for index, row in df.iterrows():
        gantt_data.append({'Task':row['工作項目'], 'Path':sheet_name, 'Team':row['團隊']})
        num_times = int(row['次'])  # 紀錄每個工項的時間段數量       
        
        # 擷取各工項的時間
        for i in range(1, num_times + 1):
            start_col = f'起始{i}'
            finish_col = f'結束{i}'
            gantt_data.append({'Start': row[f'起始{i}'],'Finish': row[f'結束{i}']})
            if pd.notna(row[start_col]) and pd.notna(row[finish_col]):  # 確認欄位不為空
                start_time = pd.to_datetime(row[start_col])
                finish_time = pd.to_datetime(row[finish_col])
                
                start_times.append(pd.to_datetime(row[start_col]))
                end_times.append(pd.to_datetime(row[finish_col]))                
                gantt_data.append({'Task':row['工作項目'], 'Path':sheet_name,'Team':row['團隊'], 'Start':row[f'起始{i}'], 'Finish':row[f'結束{i}']})
    
    # 各步道所有工項的最早起始時間與最晚結束時間
    date_se.append({'Path':sheet_name,'Start':min(start_times),'Finish':max(end_times)})
gantt_df = pd.DataFrame(gantt_data)  #把gantt_data list轉為DataFrame
date_se_df=pd.DataFrame(date_se)     #把date_se list轉為DataFrame


# color map (後續使用)    
path_color_map=['rgb(136,204,238)', 'rgb(82,106,131)', 'rgb(141,160,203)', 'rgb(128,177,211)'] # color map for paths
team_color_map={'世曦':'rgb(251,180,174)', '北科':'rgb(255,242,174)', '交大':'#B6E880', '空資':'rgb(203,213,232)', '青山':'rgb(141,160,203)', '中興':'rgb(222,203,228)'} #color map for teams


#初始化 Dash 應用-----------------------------------------------------------------------------------------------------------------------------------------------------
app = dash.Dash(__name__)

# 繪製總甘特圖
fig = px.timeline(date_se_df, x_start = "Start", x_end = "Finish", y = "Path", color = "Path", color_discrete_sequence = path_color_map, labels = {"Path": "步道"}, title="甘特圖總覽")
fig.write_html("gantt_chart.html")      # 將總甘特圖存成html

# 設置 Dash layout
app.layout = html.Div([html.Div([dcc.Dropdown(
            id = 'path-dropdown',                                                                   # 下拉式選單
            options = [{'label': sheet, 'value': sheet} for sheet in xls.sheet_names],              # 選單選項設置
            value = xls.sheet_names[0],
            multi = False,
            style = {'width': '200px','height': '40px','padding': '0 10px','font-size': '16px'})],  # 下拉式選單尺寸設置
            style = {'text-align': 'left', 'margin': '10px 0'}),                                    # 下拉式選單位置設置
            dcc.Graph(id = "gantt-chart", figure = fig)])

# 回應互動 (根據點擊選擇的步道，顯示對應的工項甘特圖)
@app.callback(Output('gantt-chart', 'figure'), [Input('path-dropdown', 'value')])
def update_gantt(selected_path): 
    # 篩選選擇的步道數據
    filtered_df = gantt_df[gantt_df['Path'] == selected_path]
    
    # 使用 plotly 表示甘特圖
    fig = px.timeline(filtered_df, x_start = "Start", x_end = "Finish", y = "Task", color = "Team",color_discrete_map = team_color_map,labels = {"Task": "工項"})

    # 版面美化
    fig.update_layout(title = dict(text = f'步道 {selected_path} 的工項甘特圖',font = dict(size = 24, color = "black", weight = "bold")),  # 調整標題字體          

                      xaxis_title = "時間",
                      yaxis_title = "工項",
                      xaxis_title_font = dict(size = 18, family = "Arial", color = "black", weight = "bold"),  # 調整X軸標籤
                      yaxis_title_font = dict(size = 18, family = "Arial", color = "black", weight = "bold"),  # 調整Y軸標籤
                      yaxis = dict(autorange = "reversed"),                                                    # 翻轉 Y 軸
                      showlegend = True, legend_title = "團隊",                                                 # legend設定                       
                      height = 1000)
    
    fig.update_xaxes(range = ["2024-10-24", "2025-12-31"],  # 固定時間軸範圍
                     dtick="M1",                            # 每個標籤間距為 1 個月 (M1 表示 1 month)
                     tickfont=dict(size=14))
    
    fig.update_yaxes(tickfont=dict(size=14), tickangle=0)

    # 設定 x 軸的屬性
    fig.update_xaxes(
        showline=True,  # 顯示 x 軸線
        linewidth=2,  # 軸線寬度
        linecolor='black',  # 軸線顏色
        mirror=True,  # 在圖表上方也顯示 x 軸線
        showgrid = True,  # 顯示格線
        gridwidth = 1,  # 線寬
        gridcolor = 'black'# 線的顏色
    )

    # 設定 y 軸的屬性
    fig.update_yaxes(
        showline=True,  # 顯示 y 軸線
        linewidth=2,
        linecolor='black',
        mirror=True,  # 在圖表右側也顯示 y 軸線
    )

    # 在每次更新後儲存 HTML
    fig.write_html("gantt_chart_updated.html")
    return fig

# 啟動 Dash 應用
if __name__ == '__main__':
    app.run_server(debug=True)