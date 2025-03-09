# ETAP 项目数据处理工具技术文档

## 概述

本文档详细介绍了一套用于 ETAP 电力系统分析软件的 Python 工具集。这些工具可以连接 ETAP 软件、执行各种电力系统分析（如潮流计算、短路计算等）、导出分析结果以及处理项目数据。

## 系统架构

该工具集由以下几个主要模块组成：

1. **连接模块**：负责与 ETAP 软件建立连接
2. **分析模块**：执行各种电力系统分析
3. **数据导出模块**：从 ETAP 导出各类元件数据
4. **结果处理模块**：处理和可视化分析结果
5. **数据转换模块**：在不同格式之间转换数据

## 核心功能模块

### 1. ETAP 连接与基础分析 (`runpf.py`, `runupf.py`, `runtdpf.py`)

这些文件提供了与 ETAP 软件连接并执行不同类型潮流计算的功能。

#### 主要类：`Main`

共同功能：
- 初始化与 ETAP 的连接
- 验证连接状态
- 执行不同类型的电力系统分析
- 导出分析结果

#### 支持的分析类型：

1. **平衡潮流计算** (`runpf.py`)
   ```python
   def run_power_flow(self, revision_name, config_name, study_case, presentation, output_report, get_online_data)
   ```

2. **不平衡潮流计算** (`runupf.py`)
   ```python
   def run_unbalanced_power_flow(self, revision_name, config_name, study_case, presentation, output_report, get_online_data)
   ```

3. **时域潮流计算** (`runtdpf.py`)
   ```python
   def run_time_domain_load_flow(self, revision_name, config_name, study_case, presentation, output_report, get_online_data, online_config_only, what_if_commands)
   ```

4. **短路计算**
   ```python
   def run_sc_cal(self, revision_name, config_name, study_case, presentation, output_report, get_online_data)
   ```

### 2. 直流系统元件数据导出 (`dc_element_output.py`)

专门用于导出 ETAP 项目中的直流系统元件数据。

#### 主要类：`ETAPExporter`

```python
class ETAPExporter:
    """ETAP 项目数据导出工具"""
```

#### 支持导出的直流元件：

- 逆变器 (`INVERTER`)
- 直流负载 (`DCLUMPLOAD`)
- 电池储能系统 (`BATTERY`)
- 直流电缆 (`DCCABLE`)
- 直流母线 (`DCBUS`)
- 直流阻抗 (`DCIMPEDANCE`)

#### 工作流程：

1. 初始化 ETAP 连接
2. 验证连接状态
3. 执行 `export_project_data()` 导出所有直流元件数据
4. 将数据保存到 Excel 文件

### 3. 模型验证工具 (`model_validate.py`)

用于通过参数扫描验证 ETAP 模型的正确性。

#### 主要功能：

- 改变指定元件的参数（如负载功率、变压器阻抗等）
- 执行潮流计算
- 记录关键节点的电压幅值和相角
- 分析参数变化对系统的影响

```python
# 参数扫描示例
start_value = 40
iterations = 10
for i in range(iterations):
    current_value = start_value + (i * 10)
    current_value_str = str(current_value)
    # 修改参数
    main.change_parameters("LUMPEDLOAD", "Lump1", "MVA", current_value_str)
    # 运行潮流计算
    main.run_power_flow(revision_name, config_name, study_case, presentation, output_report, get_online_data)
    # 导出结果
    result = main.export_pfreport()
    # 记录结果
    x_array.append(current_value)
    volt_mag1_array.append(result.volt_mag.values[0])
    # ...
```

### 4. 结果导出模块 (`export_data.py`, `export_pfdata.py`, `export_result.py`)

这些模块负责从 ETAP 生成的数据库文件中提取分析结果。

#### 主要功能：

1. **基本潮流结果导出** (`export_pfdata.py`)
   - 导出母线电压幅值和相角

2. **不平衡潮流结果导出** (`export_data.py`)
   - 计算三相电压的平均值
   - 计算三相角度的平均值

3. **时域潮流结果导出与可视化** (`export_result.py`)
   - 导出时序潮流计算结果
   - 生成系统总负荷和损耗图表
   - 生成系统最低电压图表
   - 生成系统最大支路负载图表
   - 生成母线电压和相角图表
   - 生成负荷特性图表
   - 生成负荷功率因数图表
   - 生成负荷与事件关系图表

### 5. 数据转换工具 (`convert_json.py`, `convert_xml_to_xls.py`, `data_import.py`)

这些工具用于在不同数据格式之间进行转换。

#### 主要功能：

1. **JSON 转 Excel** (`convert_json.py`)
   ```python
   def json_to_excel(json_file, excel_file)
   ```

2. **XML 转 Excel** (`convert_xml_to_xls.py`)
   ```python
   def xml_to_excel(xml_file, excel_file)
   ```

3. **文本数据导入与网络可视化** (`data_import.py`)
   - 解析特定格式的文本数据
   - 提取开关、负荷、母线、支路和电源数据
   - 使用 NetworkX 创建和可视化电网拓扑

## 使用示例

### 1. 执行平衡潮流计算并导出结果

```python
from configuration.configuration import base_address, revision_name, config_name, study_case, presentation, output_report, get_online_data

# 初始化连接
main = Main(base_address)

# 运行潮流计算
main.run_power_flow(revision_name, config_name, study_case, presentation, output_report, get_online_data)

# 导出结果
result = main.export_pfreport()

# 保存到Excel文件
result.to_excel("result.xlsx", index=False)
```

### 2. 导出直流系统元件数据

```python
from configuration.configuration import base_address

# 初始化导出工具
exporter = ETAPExporter(base_address)

# 导出所有直流元件数据
exporter.export_project_data()
```

### 3. 执行时域潮流计算并可视化结果

```python
from configuration.configuration import base_address, revision_name, config_name, study_case, presentation, output_report, get_online_data, online_config_only, what_if_commands

# 初始化连接
main = Main(base_address)

# 指定要监视的母线和负荷
custom_buses = ['BUS_都天元居开闭所（自）_220', 'BUS_kV爱涛线腾亚环网柜_207']
custom_loads = ["LOAD_南京原野制衣有限公司_939"]

# 运行时域潮流计算
main.run_time_domain_load_flow(revision_name, config_name, study_case, presentation, output_report, get_online_data, online_config_only, what_if_commands)

# 导出并可视化结果
result = main.export_output(custom_buses=custom_buses, custom_loads=custom_loads)

# 保存到Excel文件
result.to_excel("result.xlsx", index=False)
```

## 数据结构

### 1. 潮流计算结果数据结构

```
{
    "bus_ID": [母线ID列表],
    "volt_mag": [电压幅值列表],
    "volt_ang": [电压相角列表]
}
```

### 2. 时域潮流计算结果数据结构

```
{
    "Time": [时间点列表],
    "Total_Load_MW": [总负荷列表],
    "Total_Loss_MW": [总损耗列表],
    "Min_Voltage_Percent": [最低电压列表],
    "Max_Branch_Loading_Percent": [最大支路负载列表],
    "Bus_{bus_id}_Voltage": [指定母线电压列表],
    "Bus_{bus_id}_Angle": [指定母线相角列表],
    "Load_{load_id}_MW": [指定负荷有功功率列表],
    "Load_{load_id}_Mvar": [指定负荷无功功率列表],
    "Load_{load_id}_A": [指定负荷电流列表]
}
```

## 注意事项

1. 使用前需确保 ETAP 软件已启动并开启了 DataHub 服务
2. 需要正确配置 `configuration.py` 文件中的连接参数
3. 导出大量数据时可能需要较长时间，请耐心等待
4. 时域潮流计算结果的可视化需要足够的内存资源

## 依赖库

- etap.api：ETAP API 接口
- pandas：数据处理
- matplotlib：数据可视化
- numpy：数值计算
- networkx：网络可视化（仅用于 `data_import.py`）
- sqlite3：数据库操作
- xml.etree.ElementTree：XML 解析

## 总结

这套工具集提供了全面的 ETAP 数据处理功能，可以满足电力系统分析、数据导出和结果可视化的需求。通过这些工具，用户可以更高效地进行电力系统建模、分析和验证工作。