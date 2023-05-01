
import pandas as pd
df = pd.read_excel("通话记录_分析.xlsx", sheet_name=0)
grouped = df.groupby("主叫号码")["呼叫类型"].apply(lambda x: x[x == "OUTBOUND"].count())
max_number = grouped.idxmax()
max_count = grouped.max()
df["呼出时间"] = df["开始时间"].dt.hour + df["开始时间"].dt.minute / 60 + df["开始时间"].dt.second / 3600
hourly_count = df.groupby("主叫号码")["呼出时间"].apply(lambda x: x[x.notna()].nunique())
avg_hourly_count = max_count / hourly_count[max_number]
avg_interval = df.groupby("主叫号码")["上一通间隔时间(秒)"].mean()[max_number]


if max_count > 50 and avg_interval < 60:
    risk_message = "注意：{}是高风险客户，请谨慎处理！".format(max_number)
else:
    risk_message = ""
hourly_grouped = df.groupby(["主叫号码", df["呼出时间"].astype(int)])["呼叫类型"].apply(lambda x: x[x == "OUTBOUND"].count())
hourly_interval = df.groupby(["主叫号码", df["呼出时间"].astype(int)])["上一通间隔时间(秒)"].mean()
abnormal_numbers = grouped[(grouped > 50) & (df.groupby("主叫号码")["上一通间隔时间(秒)"].mean() < 120)].index.tolist() 

for number in hourly_grouped.index.get_level_values(0).unique(): 
    for hour in range(23):
        if (number, hour) in hourly_grouped.index and (number, hour + 1) in hourly_grouped.index: 
            if hourly_grouped[(number, hour)] > 20 and hourly_grouped[(number, hour + 1)] > 20: 
                abnormal_numbers.append(number) 
                break 
abnormal_numbers = list(set(abnormal_numbers)) 
output = pd.DataFrame({
    "文本结果": []
})
for number in grouped.index:
    count = grouped[number]
    interval = df.groupby("主叫号码")["上一通间隔时间(秒)"].mean()[number]
    output.loc[len(output)] = "{}共计呼出{}个电话号码，总平均通话间隔{:.2f}秒。".format(number, count, interval)
    if number in abnormal_numbers:
        output.loc[len(output)] = "注意：{}是高风险客户，请谨慎处理！".format(number)
        output.style.applymap(lambda x: Styler(bg_color='red', font_color='white'), subset=pd.IndexSlice[len(output)-1, '文本结果'])
    for hour in range(24):
        if (number, hour) in hourly_grouped.index:
            output.loc[len(output)] = "{}在{}点呼出{}个电话，平均通话间隔{:.2f}秒。".format(number, hour, hourly_grouped[(number, hour)], hourly_interval[(number, hour)])
    output.loc[len(output)] = ""
with pd.ExcelWriter("通话记录_分析.xlsx", mode="a", engine="openpyxl") as writer:
    output.to_excel(writer, sheet_name="文本结果", index=False)
