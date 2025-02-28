# ETAP数据导出工具说明文档

## 1. etap_to_excel.py

### 概述
这个Python模块提供了一个用于从ETAP项目中导出数据的工具类`ETAPExporter`。该工具可以连接ETAP系统，获取各种电力系统组件的数据，并将其导出为Excel格式。

### 主要功能
- 连接ETAP系统并验证连接状态
- 执行各类分析任务（负荷流、短路等）
- 导出以下组件的详细数据：
  - 逆变器（Inverter）
  - 直流负载（DC Load）
  - 电池系统（Battery）
  - 直流电缆（DC Cable）
  - 直流母线（DC Bus）
  - 直流阻抗（DC Impedance）

### 类结构

#### ETAPExporter
主要的导出工具类，包含以下核心方法：

##### 初始化和连接
```python
def __init__(self, base_address)
def _connect_etap(self, address)
def _verify_connection(self)
```
- 功能：建立与ETAP系统的连接并验证连接状态
- 参数：
  - `base_address`: ETAP系统的基础地址

##### 分析功能
```python
def run_analysis(self, study_type, **kwargs)
def _run_short_circuit(self, **kwargs)
```
- 功能：执行各类分析任务
- 支持的分析类型：
  - LF（负荷流）
  - ULF（统一负荷流）
  - SC（短路计算）

##### 数据导出
```python
def export_project_data(self)
```
- 功能：执行主要的数据导出流程
- 处理所有支持的组件类型
- 将成功处理的数据保存到Excel文件

##### 组件处理方法
每种组件都有专门的处理方法：
```python
def _process_inverters(self)
def _process_dc_loads(self)
def _process_batteries(self)
def _process_dc_cables(self)
def _process_dc_buses(self)
def _process_dc_impedances(self)
```

### 数据处理特性
- 错误处理：所有组件处理方法都包含完整的异常处理
- 数据验证：检查组件数据是否存在
- 跳过处理：当找不到特定组件或处理失败时，会跳过该组件
- 部分导出：即使部分数据处理失败，仍然会导出其他成功处理的数据

### 导出数据格式

#### 逆变器数据
- ID：唯一标识符
- 基本参数：服务状态、总线ID、网络连接等
- 电气参数：AC电压、DC电压、功率等
- PV曲线数据：发电和充电特性

#### 直流负载数据
- ID：唯一标识符
- 基本参数：服务状态、连接母线等
- 额定参数：电压、功率

#### 电池系统数据
- ID：唯一标识符
- 基本参数：服务状态、连接母线等
- 配置参数：电池单元数、组数、串数

#### 直流电缆数据
- ID：唯一标识符
- 连接信息：起始母线、终止母线
- 物理参数：长度、阻抗等

#### 直流母线数据
- ID：唯一标识符
- 基本参数：额定电压、服务状态等

#### 直流阻抗数据
- ID：唯一标识符
- 连接信息：起始母线、终止母线
- 电气参数：电阻值、电感值

### 使用示例
```python
from configuration.configuration import base_address

# 创建导出器实例
exporter = ETAPExporter(base_address)

# 执行数据导出
exporter.export_project_data()
```

### 输出文件
- 文件名：parameters.xlsx
- 格式：Excel工作簿
- 工作表：每种组件类型对应一个工作表
- 数据组织：每个组件类型的所有实例按行排列，属性按列排列

### 错误处理
- 连接错误：显示详细的连接状态信息
- 数据处理错误：跳过问题组件，继续处理其他组件
- 文件保存错误：提供错误详情

### 注意事项
1. 确保ETAP系统可访问且配置正确
2. 检查输出目录的写入权限
3. 对于大型项目，预留足够的处理时间
4. 定期检查日志输出，了解处理状态

## 2. time_domain_pf.py

### 概述
这个Python模块用于执行ETAP时域潮流计算并导出结果。它包含两个主要部分：Main类用于与ETAP交互执行计算，export_time_series_power_flow函数用于处理和可视化计算结果。

### Main类功能
```python
class Main():
    def __init__(self, base_address)
    def run_power_flow(self, revision_name, config_name, study_case, presentation, output_report, get_online_data)
    def run_unbalanced_power_flow(self, revision_name, config_name, study_case, presentation, output_report, get_online_data)
    def run_sc_cal(self, revision_name, config_name, study_case, presentation, output_report, get_online_data)
    def run_time_domain_load_flow(self, revision_name, config_name, study_case, presentation, output_report, get_online_data, online_config_only, what_if_commands)
    def export_report(self)
    def change_parameters(self, elementType, elementName, fieldName, value)
    def export_output(self)
```

- 初始化：建立与ETAP系统的连接
- 分析功能：支持多种电力系统分析（负荷流、不平衡负荷流、短路计算、时域负荷流）
- 参数修改：允许修改ETAP模型中的元素参数
- 结果导出：调用export_time_series_power_flow函数处理结果

### export_time_series_power_flow函数功能
```python
def export_time_series_power_flow(database_path, output_file="time_series_results.txt", custom_buses=None, custom_loads=None)
```

该函数处理时域潮流计算的SQLite数据库结果，提供以下功能：

1. **数据提取与处理**：
   - 系统总负荷和损耗
   - 系统最低电压
   - 系统最大支路负载
   - 指定母线的电压和相角
   - 指定负荷的有功功率、无功功率和电流
   - 系统事件信息

2. **可视化图表生成**：
   - 系统总览图（负荷和损耗、最低电压、最大支路负载、关键母线电压）
   - 母线电压和相角图
   - 负荷特性图（有功功率、无功功率、电流）
   - 负荷功率因数图
   - 负荷与事件关联图

3. **数据导出**：
   - 文本报告（详细数据）
   - Excel表格（便于进一步分析）
   - 高质量图表（用于报告和展示）

### 使用方法
```python
if __name__ == "__main__":
    # 从配置文件导入必要参数
    main = Main(base_address)
    # 执行时域潮流计算
    main.run_time_domain_load_flow(revision_name, config_name, study_case, presentation, output_report, get_online_data, online_config_only, what_if_commands)
    # 导出和处理结果
    result = main.export_output()
    # 保存结果到Excel
    result.to_excel("result.xlsx", index=False)
```

### 输出内容
- 文本报告：time_series_results.txt
- Excel数据：time_series_results.xlsx
- 系统总览图：system_overview.png
- 母线电压和相角图：bus_voltage_angle.png
- 负荷特性图：load_characteristics.png
- 负荷功率因数图：load_power_factor.png
- 负荷与事件关系图：load_events_correlation.png

### 自定义分析
可以通过指定参数来分析特定母线和负荷：
```python
custom_buses = ['BUS_CNODE_JCT__1333', 'BUS_都天元居开闭所（自）_220', 'BUS_kV爱涛线腾亚环网柜_207']
custom_loads = ['Load_1', 'Load_2', 'Load_3']
```

### 负荷数据分析特性
- **有功功率分析**：追踪负荷有功功率随时间的变化
- **无功功率分析**：监测负荷无功功率的波动
- **电流分析**：观察负荷电流的变化趋势
- **功率因数计算**：基于有功和无功功率自动计算功率因数
- **多负荷对比**：在同一图表中比较多个负荷的特性

### 错误处理
该模块包含完整的异常处理机制，确保在数据处理过程中出现问题时能够提供有用的错误信息，并安全关闭数据库连接。