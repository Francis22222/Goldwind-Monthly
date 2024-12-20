import dash
from dash import dcc, html, Input, Output
import dash_bootstrap_components as dbc
import pandas as pd  # 数据处理
import numpy as np  # 数值计算
from openpyxl import load_workbook  # 用于读取 Excel 文件
import plotly.graph_objects as go  # Plotly 图表
import plotly.express as px
from dash import Dash, dcc, html, Input, Output
from dash import dash_table

# 1. 数据预处理
# Excel 文件路径
file_path = r"C:\Users\Administrator\Desktop\Tableau\2.Data\Vietnam Data-20241118.xlsx"
sheet_name = "Production statistics"

# 读取 Excel 文件
df = pd.read_excel(file_path, sheet_name=sheet_name)

# 数据清洗和处理
# 1. 转换日期格式
df["Statistical time"] = pd.to_datetime(df["Statistical time"], errors="coerce")

# 2. 将所有变量列转换为数值类型
columns_to_numeric = [
    "Average ambient temperature (°C)",
    "Active Energy Imported(kWh)",
    "Active Energy Exported(kWh)",
    "Energy production time (h)",
    "Equivalent Utilization Hours (H)",
    "Loss due to curtailment (kWh)",
    "Curtailment duration (h)",
    "Loss due to fault (kWh)",
    "Fault duration (h)",
    "Average wind speed (m/s)"
]

# 转换为数值类型，遇到错误数据会变成 NaN，但不删除整行
for col in columns_to_numeric:
    df[col] = pd.to_numeric(df[col], errors="coerce")  # 仅将无法转换的值变为 NaN

# 删除关键列中为 NaN 的数据，而不是整行
df = df.dropna(subset=["Statistical time", "Power plant name"])  # 确保关键列没有 NaN

# 3. 添加时间相关列（年、月、周等）
df["Year"] = df["Statistical time"].dt.year
df["Month"] = df["Statistical time"].dt.month
df["Week"] = df["Statistical time"].dt.isocalendar().week
df["Day"] = df["Statistical time"].dt.date

# 获取所有可用年份和电厂
available_years = sorted(df["Year"].unique(), reverse=True)
available_projects = sorted(df["Power plant name"].unique())

# 2. 初始化 Dash 应用
app = dash.Dash(__name__)

# 定义表格组件（移除 'style' 参数）
table_layout = dash_table.DataTable(
    id="data-table",
    style_table={
        "maxHeight": "185px",  # 最大高度，出现滚动条
        "width": "950px",  # 固定宽度
        "overflowY": "auto",
        "overflowX": "auto",
    },
    style_cell={
        "textAlign": "center",
        "fontFamily": "Calibri",
        "fontSize": "12px",
        "padding": "5px",
    },
    style_header={
        "backgroundColor": "#0d3057",#"lightgrey",
        "fontWeight": "bold",
        "textAlign": "center",
        "color": "white"  # 字体颜色为白色
    },
    fixed_rows={"headers": True},
    page_action="none",
    export_format="xlsx",
)

# 3. 定义布局
layout = html.Div("Asia Region Project Monthly Report")
app.layout = html.Div(
    style={
        "width": "1600px",  # 设置仪表盘的宽度为 1600px
        "height": "900px",  # 设置仪表盘的高度为 900px
        "margin": "0 auto",  # 居中显示
        "padding": "20px",  # 添加内边距
        "boxSizing": "border-box",  # 包括内边距在内的总宽度
        "overflow": "hidden",  # 防止超出内容显示滚动条
    },
    children=[
        html.H1(
            "Asia Region Project Monthly Report",
            style={
                "textAlign": "center",  # 文本居中
                "height": "43px",  # 高度设置为 43px
                "lineHeight": "43px",  # 行高设置为 43px，确保文本在垂直方向居中
                "fontFamily": "Calibri",  # 字体设置为 Calibri
                "fontSize": "22px",  # 字体大小设置为 22px
                "fontWeight": "bold",  # 字体加粗
                "color": "#0d3057",  # 字体颜色为深蓝色
                "margin": "0" , # 移除默认的上下外边距
            }
        ),

    # 筛选器
        html.Div([
            html.Div([
                html.Label("Year:"),
                dcc.Dropdown(
                    id="year-selector",
                    options=[{"label": str(year), "value": year} for year in available_years],
                    value=available_years[0],
                    clearable=False
                )
            ], style={"marginLeft": "50px", "width": "200px", "display": "inline-block", "height": "35px",
                      "font-size": "10px"}),

            html.Div([
                html.Label("Month:"),
                dcc.Dropdown(
                    id="month-selector",
                    options=[{"label": month, "value": idx} for idx, month in enumerate(
                        ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                         "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)],
                    value=1,
                    clearable=False
                )
            ], style={"marginLeft": "100px", "width": "200px", "display": "inline-block", "height": "35px",
                      "font-size": "10px"}),
            html.Div([
                html.Label("Project Name:"),
                dcc.Dropdown(
                    id="plant-selector",
                    options=[{"label": plant, "value": plant} for plant in available_projects],
                    value=available_projects[0],
                    clearable=False
                )
            ], style={"marginLeft": "100px", "width": "200px", "display": "inline-block", "height": "35px",
                      "font-size": "10px"})
        ], style={"textAlign": "left", "marginBottom": "30px"}),

    #图表布局
    html.Div(
            style={"position": "relative", "width": "1600px", "height": "900px"},  # 设置整体容器的宽度和高度
            children=[
                html.Div(
                    dcc.Graph(
                    # 第一行三个图表
                        id="monthly-wind-chart",
                        style={"width": "550px", "height": "140px"},  # 修正宽和高
                    ),
                    style={"position": "absolute", "left": "10px", "top": "40px"},  # 设置位置
                ),
                html.Div(
                    dcc.Graph(
                    id="weekly-wind-chart",
                    style={"width": "550px", "height": "216px"},  # 修正宽和高
                    ),
                    style={"position": "absolute", "left": "10px", "top": "200px"},  # 设置位置
                ),
                html.Div(
                    dcc.Graph(
                    id="annual-production-chart",
                    style={"width": "550px", "height": "100px"},  # 修正宽和高
                    ),
                    style={"position": "absolute", "left": "10px", "top": "430px"},  # 设置位置
                ),
                html.Div(
                    dcc.Graph(
                    id="combined-chart",
                    style={"width": "550px", "height": "225px"},  # 修正宽和高
                    ),
                    style={"position": "absolute", "left": "10px", "top": "550px"},  # 设置位置
                ),
                html.Div(
                    dcc.Graph(
                    id="combined-chart2",
                    style={"width": "1020px", "height": "550px"},  # 修正宽和高
                    ),
                    style={"position": "absolute", "left": "570px", "top": "40px"},  # 设置位置
                ),
                html.Div(
                    table_layout,
                    style={"position": "absolute", "left": "590px", "top": "550px"},  # 修正宽和高
                ),
            ]
        ),
    ],
)

# 回调函数更新表格
@app.callback(
    Output("data-table", "data"),
    Output("data-table", "columns"),
    [
        Input("year-selector", "value"),
        Input("month-selector", "value"),
        Input("plant-selector", "value")
    ]
)
def update_table(selected_year, selected_month, selected_project):
    # 筛选数据
    filtered_df = df[
        (df["Year"] == selected_year) &
        (df["Month"] == selected_month) &
        (df["Power plant name"] == selected_project)
    ]

    # 按 Device Name 和日期分组，统计变量的日数据
    grouped_df = filtered_df.groupby(["Device Name", "Day"]).agg({
        "Average ambient temperature (°C)": "mean",
        "Active Energy Imported(kWh)": "sum",
        "Active Energy Exported(kWh)": "sum",
        "Energy production time (h)": "sum",
        "Equivalent Utilization Hours (H)": "sum",
        "Loss due to curtailment (kWh)": "sum",
        "Curtailment duration (h)": "sum",
        "Loss due to fault (kWh)": "sum",
        "Fault duration (h)": "sum",
        "Average wind speed (m/s)": "mean"
    }).reset_index()

    # 转换为适合 DataTable 的格式
    data = grouped_df.to_dict("records")  # 转换为字典列表
    columns = [{"name": col, "id": col} for col in grouped_df.columns]  # 列名和列 ID

    return data, columns

# 4. 定义回调函数
@app.callback(
    Output("monthly-wind-chart", "figure"),
    Input("plant-selector", "value")
)
def update_chart(selected_project):
    filtered_df = df[df["Power plant name"] == selected_project]
    # 图表 1：仅使用 project 筛选
    monthly_data = df[df["Power plant name"] == selected_project].groupby(["Year", "Month"]).agg({
        "Average wind speed (m/s)": "mean"
    }).reset_index()

    # 2. 构造完整的月份数据框架，确保所有月份（1-12）都有记录
    all_years = df["Year"].unique()
    all_months = list(range(1, 13))  # 1 到 12 月
    full_index = pd.MultiIndex.from_product([all_years, all_months], names=["Year", "Month"])
    monthly_data = monthly_data.set_index(["Year", "Month"]).reindex(full_index).reset_index()

    # 3. 将月份数字转换为英文名称的前三个字母
    monthly_data["Month"] = monthly_data["Month"].apply(lambda x: pd.to_datetime(f"2024-{x}-01").strftime("%b"))

    # 4. 固定月份顺序为 "Jan" 到 "Dec"
    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    monthly_data["Month"] = pd.Categorical(monthly_data["Month"], categories=month_order, ordered=True)
    # 5. 创建对比年份的折线图
    monthly_wind_chart = go.Figure()
    for year in monthly_data["Year"].unique():
        yearly_data = monthly_data[monthly_data["Year"] == year]
        monthly_wind_chart.add_trace(
            go.Scatter(
                x=yearly_data["Month"],
                y=yearly_data["Average wind speed (m/s)"],
                mode="lines+markers",
                name=str(year),
                line = dict(width=2)  # 设置线条宽度
            )
        )
    #原先代码monthly_wind_chart.update_layout(title="Monthly Wind Speed", xaxis_title="Month", yaxis_title="Wind Speed (m/s)")
    # 更新图表布局
    monthly_wind_chart.update_layout(
        title=dict(
            text="Comparison of Annual Average Wind Speed",  # 标题文本
            font=dict(
                family="Arial",  # 字体
                size=14,  # 字体大小
                color="black"  # 字体颜色
            ),
            x=0.5,  # 标题水平居中
            xanchor="center",  # 锚点设置为左侧
            pad=dict(
                t=30  # 标题和图表的间距
            )
        ),
        xaxis=dict(
            categoryorder="array",
            categoryarray=month_order,  # 按固定顺序排列月份
            showgrid = False  # 取消 X 轴网格线
        ),
        yaxis=dict(
            showgrid=False  # 取消 Y 轴网格线
        ),
        margin=dict(
            t=20,  # 图表整体距离顶部的外边距（标题包含在此范围内）
            b=10, # 图表下边距
            l=10, # 设置左侧边距为 15px
            r=10 # 图表右边距
        ),
        legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.2),
        plot_bgcolor = "white"  # 设置背景为白色
        )

    return monthly_wind_chart


@app.callback(
    Output("weekly-wind-chart", "figure"),
      Input("plant-selector", "value")
)
def update_wind_speed_year_graph(selected_project):
    # 筛选数据
    filtered_df = df[df["Power plant name"] == selected_project]
    # 图表 2：周风速变化区域图
    # 1. 计算整个数据集的开始日期和结束日期（不受年份筛选器影响）
    start_date = df["Statistical time"].min()  # 数据集的最早日期
    end_date = df["Statistical time"].max()  # 数据集的最晚日期

    # 2. 按每周分组，计算每周的平均风速
    weekly_data = df[df["Power plant name"] == selected_project].resample("W-Mon", on="Statistical time").agg({
        "Average wind speed (m/s)": "mean"
    }).reset_index()

    # 3. 创建完整的日期范围（从开始日期到结束日期，每周一为时间点）
    full_date_range = pd.date_range(start=start_date, end=end_date, freq="W-Mon")

    # 4. 将完整的日期范围与每周数据对齐
    weekly_data = weekly_data.set_index("Statistical time").reindex(full_date_range).reset_index()
    weekly_data.columns = ["Statistical time", "Average wind speed (m/s)"]

    # 5. 绘制区域图
    weekly_wind_chart = go.Figure(
        data=go.Scatter(
            x=weekly_data["Statistical time"],# X 轴为完整的每周时间点
            y=weekly_data["Average wind speed (m/s)"],# Y 轴为每周的平均风速
            mode="lines",# 显示线条模式
            fill="tozeroy",# 填充区域到 Y=0
            name="Weekly Average Wind Speed",
            line = dict(color="#0d3057", width=2),  # 设置线条颜色
            fillcolor = "#0d3057"  # 区域填充颜色

        )
    )
    #原先代码weekly_wind_chart.update_layout(title="Weekly Wind Speed", xaxis_title="Week", yaxis_title="Wind Speed (m/s)")
    # 6. 更新图表布局
    weekly_wind_chart.update_layout(
        title=dict(
            text="Variation of Weekly Average Wind Speed",  # 标题文本
            font=dict(
                family="Arial",  # 字体
                size=14,  # 字体大小
                color="black"  # 字体颜色
            ),
            x=0.5,  # 标题水平居中
            xanchor="center",
            pad=dict(
                t=1  # 标题和图表的间距
            )
        ),
        xaxis=dict(
            showgrid=False,  # 不显示网格线
            tickformat="%Y-%m-%d",  # 日期格式化为 年-月-日
        ),
        yaxis=dict(
            showgrid=True  # 显示网格线
        ),
        margin=dict(
            t=20,  # 图表整体距离顶部的外边距（标题包含在此范围内）
            b=10,
            l=10,
            r=10
        ),
        legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.2),
        plot_bgcolor="white"  # 设置背景为白色
    )

    return weekly_wind_chart


@app.callback(
        Output("annual-production-chart", "figure"),
        Input("plant-selector", "value")
    )
def update_wind_speed_year_graph(selected_project):
        # 筛选数据
    filtered_df = df[df["Power plant name"] == selected_project]
    # 图表 3：年度产量柱状图（纵坐标为年份，横坐标为年度产量）
    annual_data = df[df["Power plant name"] == selected_project].groupby("Year").agg({
        "Active Energy Exported(kWh)": "sum"
    }).reset_index()

    annual_production_chart = go.Figure(
        data=go.Bar(
            y=annual_data["Year"],  # 纵坐标为年份
            x=annual_data["Active Energy Exported(kWh)"],  # 横坐标为年度产量
            orientation="h",  # 将柱状图设置为水平
            marker_color="#0d3057",  # 图形颜色与图表 2 一致
            name="Annual Production (kWh)"
        )
    )
    #annual_production_chart.update_layout(title="Annual Production", xaxis_title="Year", yaxis_title="Production (kWh)")

    annual_production_chart.update_layout(
        title=dict(
            text="Statistics of Annual Power Production",  # 标题文本
            font=dict(
                family="Arial",  # 字体
                size=14,  # 字体大小
                color="black"  # 字体颜色
            ),
            x=0.5,  # 标题水平居中
            xanchor="center",
            pad=dict(
                t=10  # 标题和图表的间距
            )
        ),
        yaxis=dict(
            tickmode="linear",  # 仅显示整数年份
            tick0=annual_data["Year"].min(),  # 起始年份
            dtick=1  # 每次递增 1
        ),
        margin=dict(
            t=30,  # 图表整体距离顶部的外边距（标题包含在此范围内）
            b=10,
            l=10,
            r=10
        ),
        legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.2),
        plot_bgcolor="white"  # 设置背景为白色
    )

    return annual_production_chart


@app.callback(
    Output("combined-chart", "figure"),
    [
        Input("year-selector", "value"),
        Input("plant-selector", "value")
     ]
)
def update_monthly_energy_production_graph(selected_year, selected_project):
    # 筛选数据
    filtered_df = df[
        (df["Year"] == selected_year) &
        (df["Power plant name"] == selected_project)
    ]
    #图表4风速与功率双Y轴图表
    # 按年份分组，计算每年的月度产量和平均风

    monthly_data_combined = filtered_df.groupby("Month").agg({
        "Active Energy Exported(kWh)": "sum",
        "Average wind speed (m/s)": "mean"
    }).reset_index()

    # 将月份数字转换为英文的前三个字母
    monthly_data_combined["Month"] = monthly_data_combined["Month"].apply(
        lambda x: pd.to_datetime(f"2024-{x}-01").strftime("%b")  # 使用 2024 年的任意日期转换月份
    )

    # 固定月份顺序为 "Jan" 到 "Dec"
    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    monthly_data_combined["Month"] = pd.Categorical(monthly_data_combined["Month"], categories=month_order,
                                                    ordered=True)

    # 绘制柱状图（年度产量）
    bar_chart = go.Bar(
        x=monthly_data_combined["Month"],# X 轴为月份
        y=monthly_data_combined["Active Energy Exported(kWh)"],# Y 轴为年度产量
        name="Production (kWh)",
        yaxis="y1",
        marker_color = "#0d3057",  # 柱体颜色为深蓝色
        showlegend = False  # 隐藏图例
    )

    # 绘制折线图（平均风速）
    line_chart = go.Scatter(
        x=monthly_data_combined["Month"],# X 轴为月份
        y=monthly_data_combined["Average wind speed (m/s)"],# Y 轴为平均风速
        name="Average Wind Speed (m/s)",
        yaxis="y2",
        mode="lines+markers",
        line = dict(color="red", width=2),  # 折线颜色为红色
        showlegend = False  # 隐藏图例
    )
    combined_chart = go.Figure(data=[bar_chart, line_chart])

    combined_chart.update_layout(
        title=dict(
            text="Monthly Production and Wind Speed Variation",  # 标题文本
            font=dict(
                family="Arial",  # 字体
                size=14,  # 字体大小
                color="black"  # 字体颜色
            ),
            x=0.5,  # 标题水平居中
            xanchor="center",
            pad=dict(
                t=30  # 标题和图表的间距
            )
        ),
        xaxis=dict(
            #title="Month",  # X 轴标题
            categoryorder="array",  # 按固定顺序排列月份
            categoryarray=month_order,
            showgrid = False,  # 取消 X 轴网格线

        ),

        yaxis=dict(
            # title="Production (kWh)",  # 左侧 Y 轴标题
            titlefont=dict(color="#0d3057"),  # 左侧 Y 轴标题颜色
            tickfont=dict(color="#0d3057"), # 左侧 Y 轴刻度颜色
            showgrid = False  # 取消 Y 轴网格线
        ),
        yaxis2=dict(
            # title="Wind Speed (m/s)",  # 右侧 Y 轴标题
            titlefont=dict(color="red"),  # 右侧 Y 轴标题颜色
            tickfont=dict(color="red"),  # 右侧 Y 轴刻度颜色
            overlaying="y",  # 右侧 Y 轴与左侧 Y 轴重叠
            side="right",  # 右侧显示
            showgrid = False  # 取消 Y 轴网格线
        ),
        margin=dict(
            t=20,  # 图表整体距离顶部的外边距（标题包含在此范围内）
            b=10,
            l=10,
            r=10
        ),
        legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.2),
        plot_bgcolor="white"  # 设置背景为白色
    )

    return combined_chart




# 图表 5：使用 project、year 和 month 筛选
@app.callback(
    Output("combined-chart2", "figure"),
    [
        Input("year-selector", "value"),
        Input("month-selector", "value"),
        Input("plant-selector", "value")
    ]
)
def update_combined_chart2(selected_year, selected_month, selected_project):
    # 筛选数据
    filtered_df = df[
        (df["Year"] == selected_year) &
        (df["Month"] == selected_month) &
        (df["Power plant name"] == selected_project)
        ]

    # 生成当前月的完整日期范围
    start_date = pd.Timestamp(year=selected_year, month=selected_month, day=1)
    end_date = start_date + pd.offsets.MonthEnd(0)
    date_range = pd.date_range(start=start_date, end=end_date, freq="D")

    # 数据分组并重新索引至完整日期范围
    daily_data = filtered_df.groupby("Statistical time").agg({
        "Active Energy Exported(kWh)": "sum",
        "Average wind speed (m/s)": "mean"
    })

    # 将日期设置为索引并重设索引为完整的日期范围
    daily_data = daily_data.reindex(date_range).reset_index()
    daily_data.rename(columns={"index": "Statistical time"}, inplace=True)

    # 填充缺失值为 0
    daily_data["Active Energy Exported(kWh)"] = daily_data["Active Energy Exported(kWh)"].fillna(0)
    daily_data["Average wind speed (m/s)"] = daily_data["Average wind speed (m/s)"].fillna(0)

    # 创建图表 5
    daily_bar_chart = go.Bar(
        x=daily_data["Statistical time"],  # X 轴为日期
        y=daily_data["Active Energy Exported(kWh)"],
        name="Production (kWh)",
        yaxis="y1",
        marker_color="#0d3057",  # 柱体颜色为深蓝色
        showlegend=False  # 隐藏图例
    )
    daily_line_chart = go.Scatter(
        x=daily_data["Statistical time"],  # X 轴为日期
        y=daily_data["Average wind speed (m/s)"],
        name="Wind Speed",
        yaxis="y2",
        mode="lines+markers",
        line=dict(color="red"),
        showlegend=False  # 隐藏图例
    )
    combined_chart2 = go.Figure(data=[daily_bar_chart, daily_line_chart])

    combined_chart2.update_layout(
        title=dict(
            text="Daily Production and Wind Speed Variation",  # 标题文本
            font=dict(
                family="Arial",  # 字体
                size=14,  # 字体大小
                color="black"  # 字体颜色
            ),
            x=0.5,  # 标题水平居中
            xanchor="center",
            pad=dict(
                t=30  # 标题和图表的间距
            )
        ),
        xaxis=dict(
            #title="",  # 可根据需要添加 X 轴标题
            showgrid=False,  # 取消 X 轴网格线
            tickformat="%d",  # 显示日期为日数字
            dtick="D1",  # 每天一个刻度
            tickangle=-45  # 旋转刻度标签，以防止重叠
        ),
        yaxis=dict(
            #title="Production (kWh)",
            titlefont=dict(color="#0d3057"),  # 左侧 Y 轴标题颜色
            tickfont=dict(color="#0d3057"),  # 左侧 Y 轴刻度颜色
            showgrid=False  # 取消 Y 轴网格线
        ),
        yaxis2=dict(
            #title="Wind Speed (m/s)",  # 右侧 Y 轴标题
            titlefont=dict(color="red"),  # 右侧 Y 轴标题颜色
            tickfont=dict(color="red"),  # 右侧 Y 轴刻度颜色
            overlaying="y",  # 右侧 Y 轴与左侧 Y 轴重叠
            side="right",  # 右侧显示
            showgrid=False  # 取消 Y 轴网格线
        ),
        margin=dict(
            t=50,  # 图表整体距离顶部的外边距（标题包含在此范围内）
            b=80,
            l=60,
            r=60
        ),
        legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.2),
        plot_bgcolor="white"  # 设置背景为白色
    )

    # 返回图表
    return combined_chart2

if __name__ == "__main__":
    app.run_server(debug=True, port=8051)