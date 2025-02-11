# ETAP数据导出工具技术文档

## 1. 概述

该工具用于从ETAP工程项目中导出关键设备参数和配置数据，并将其保存为Excel格式。支持多种电力系统设备的数据导出，包括逆变器、直流负载、电池、电缆、母线和阻抗等组件。

## 2. 主要功能

- ETAP项目连接与验证
- 支持多种分析类型运行(LF、ULF、SC)
- 导出以下设备的详细参数:
  - 逆变器(Inverter)
  - 直流负载(DC Load)
  - 电池储能系统(Battery)
  - 直流电缆(DC Cable)
  - 直流母线(DC Bus)
  - 直流阻抗(DC Impedance)

## 3. 类结构

### ETAPExporter类

主要类，负责所有数据导出操作。

#### 核心方法

```python
def __init__(self, base_address)
def export_project_data(self)
def run_analysis(self, study_type, **kwargs)
```

## 4. 数据导出明细

### 4.1 逆变器数据
- 基本信息：ID、运行状态、连接母线等
- 电气参数：AC电压、DC电压、功率等
- PV曲线数据：发电和充电模式下的V-P特性

### 4.2 直流负载数据
- ID标识
- 运行状态
- 额定电压
- 额定功率

### 4.3 电池储能系统数据
- 基本配置信息
- 电池组结构参数（电池数、组数、串数）

### 4.4 直流电缆数据
- 首末端母线
- 长度及单位
- 阻抗参数

### 4.5 直流母线数据
- 标识信息
- 额定电压
- 运行状态

### 4.6 直流阻抗数据
- 连接信息
- 阻抗参数(R值、L值)

## 5. 使用示例

```python
from configuration.configuration import base_address

# 创建导出器实例
exporter = ETAPExporter(base_address)

# 执行数据导出
exporter.export_project_data()
```

## 6. 输出格式

数据以Excel文件格式输出，包含多个工作表：
- INVERTER
- DCLUMPLOAD
- BATTERY
- DCCABLE
- DCBUS
- DC_IMPEDANCE

## 7. 注意事项

1. 使用前需确保ETAP服务正常运行
2. 需要正确配置base_address
3. 确保有足够的权限访问ETAP项目
4. 导出过程中请勿关闭ETAP程序

## 8. 依赖项

- etap.api
- pandas
- json
- xml.etree.ElementTree

## 9. 错误处理

工具包含基本的连接验证机制，会在初始化时检查：
- ETAP连接状态
- 文件路径有效性
- DataHub状态