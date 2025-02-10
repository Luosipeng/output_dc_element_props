import csv
import re
import os
import networkx as nx
import matplotlib.pyplot as plt
import pandas as pd

path = r'C:\Codes\etap'
data = r'\data\test_data.TXT'
output_dir =  path + r'\data'
# Define file paths


input_file_path = path + data
output_file_path = path + r'\data\test_data.xlsx'

# Read the content of the text file
with open(input_file_path, 'r', encoding='gbk', errors='replace') as file:
    lines = file.readlines()

# Define section headers
headers = {
    "switch": ["编号", "开关状态", "P量测", "Q量测", "IA量测", "IB量测", "IC量测",
               "P计算", "Q计算", "IA计算", "IB计算", "IC计算", "Uab计算", "相角",
               "基准电压U", "满量测", "设备ID", "表号", "名称", "厂站名称",
               "首端点号", "末端点号", "首端节点", "末端节点"],
    "load": ["编号", "母线编号", "P量测", "Q量测", "I量测", "P计算", "Q计算",
             "I计算", "Vab计算", "Pa计算", "Qa计算", "Pb计算", "Qb计算", "Pc计算",
             "Qc计算", "基准电压U", "满量测", "节点ID", "设备ID", "表号", "设备名", "厂站名"],
    "bus": ["母线编号", "busno3", "V量测", "V标幺", "Va", "Vb", "Vc", "Ubase",
            "PhaseA", "PhaseB", "PhaseC", "节点数", "节点描述", "母线描述", "所属馈线"],
    "branch": ["序号", "编号", "pos", "fBus", "tBus", "fBus3", "tBus3", "Pij", "Qij", "Pji", "Qji",
               "Vf", "Vt", "IA", "IB", "IC", "Imax", "rate", "ind", "jnd", "frLayer",
               "toLayer", "r", "x", "ploss", "qloss", "支路ID", "表号", "名称", "功率方向"],
    "source": ["BusNo3", "PG", "QG", "Vmea", "Vmag", "母线id", "负荷描述", "负荷节点号"]
}

# Initialize containers for each section's data
data = {
    "switch": [],
    "load": [],
    "bus": [],
    "branch": [],
    "source": []
}

# Function to parse a line into sections
def parse_line(line, current_section):
    parts = re.split(r'[,\s]\s*', line)
    if current_section == "switch":
        return [parts[0].split(':')[0], parts[0].split(':')[1], parts[1][1:]] + parts[2].strip('()').split(',') + parts[3].strip('()').split(',') + parts[4:5] + [parts[5][:-1], parts[6][1:], parts[7]] + parts[8:12] + [parts[12][:-1]] + parts[13:18] + [parts[18][1:], parts[19][:-1], parts[20][1:],parts[21][:-1]]
    elif current_section == "load":
        return [parts[0].split(':')[0], parts[0].split(':')[1], parts[1][1:]] + parts[2].strip('[]').split(',') + parts[3].strip(']').split(',') + parts[4].strip('[').split(',') + parts[5:7] + parts[7].strip(']').split(',') + parts[8].strip('[').split(',') + parts[9:13] + parts[13].strip(']').split(',') + parts[14:]
    elif current_section == "bus":
        return [parts[0].split(':')[0], parts[0].split(':')[1], parts[1], parts[2][1:]] + parts[3:6] + parts[6].strip(')').split(',') + parts[7].strip('(').split(',') + [parts[8]] + parts[9].strip(')').split(',') + parts[10:]
    elif current_section == "branch":
        return [parts[0].split(':')[0], parts[0].split(':')[1], parts[1], parts[2][1:], parts[3][:-1], parts[4][1:], parts[5][:-1],parts[6][1:], parts[7][:-1], parts[8][1:], parts[9][:-1], parts[10][1:], parts[11][:-1], parts[12][3:], parts[13], parts[14][:4], parts[14][6:12], parts[15], parts[16][1:], parts[17][:-1], parts[18][1:], parts[19][:-1], parts[20][1:], parts[21][:-1], parts[22][1:], parts[23][:-1], parts[24:]]
    elif current_section == "source":
        return [parts[0], parts[1][1:], parts[2][:-1], parts[3][1:], parts[4][:-1], parts[5], parts[6], parts[7], parts[8]]
    else:
        return parts

current_section = None
skip_lines = 0

# Read the content of the text file using 'gbk' encoding
with open(input_file_path, 'r', encoding='gbk', errors='replace') as file:
    for line in file:
        line = line.strip()
        if line.startswith("-------------开始打印开关数据"):
            current_section = "switch"
            skip_lines = 2
        elif line.startswith("-------------开始打印负荷数据"):
            current_section = "load"
            skip_lines = 2
        elif line.startswith("-----------------开始打印母线数据"):
            current_section = "bus"
            skip_lines = 2
        elif line.startswith("-------------开始打印支路数据"):
            current_section = "branch"
            skip_lines = 2
        elif line.startswith("-----------------开始打印电源数据"):
            current_section = "source"
            skip_lines = 2
        elif skip_lines > 0:
            skip_lines -= 1
        elif current_section and ':' in line:
            data[current_section].append(parse_line(line, current_section))

# Write the parsed data to CSV files
# In some cases, the length extracted might not be the same as the length of the header.

# Write the parsed data to CSV files
for section, rows in data.items():
    if rows:
        section_file_path = os.path.join(output_dir, f'test_data_{section}.csv')
        with open(section_file_path, 'w', newline='', encoding='gbk') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(headers[section])  # Write the header
            csvwriter.writerows(rows)

# # Convert each section's data to a DataFrame and write to an Excel file with multiple sheets
# with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
#     for section, rows in data.items():
#         if rows:
#             df = pd.DataFrame(rows, columns=headers[section])
#             df.to_excel(writer, sheet_name=section, index=False)

# Create a graph object
Network = nx.Graph()

nb = len(data["bus"])
nl = len(data["branch"])
for n in range(nb):
    Network.add_node(data["bus"][n][0])
for l in range(nl):
    Network.add_edge(data["branch"][l][3], data["branch"][l][4])

pos = nx.spring_layout(Network)  # you can also use other layouts like circular_layout, shell_layout, etc.
nx.draw(Network, pos, with_labels=True, node_color='lightblue', node_size=100, edge_color='grey', width=2, linewidths=10, font_size=15)

# Draw the graph
# nx.draw(Network, with_labels=True, node_color='skyblue', node_size=1, edge_color='black', linewidths=100, font_size=15)
# Show the plot
plt.title("Distribution Network Graph")
plt.show()


print("Conversion complete. CSV files have been created for each section.")
