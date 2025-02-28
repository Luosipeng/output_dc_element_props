import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.dates import DateFormatter
import datetime
import matplotlib.font_manager as fm
from matplotlib import rcParams


def export_time_series_power_flow(database_path, output_file="time_series_results.txt",
                                  custom_buses=None,custom_loads=None):
    """
    导出时序潮流计算结果并生成可视化图表

    参数:
    database_path: SQLite数据库文件的路径
    output_file: 输出文本文件的路径
    custom_buses: 用户指定的要查看电压和相角的母线ID列表，例如 ['Bus_1', 'Bus_2']
    custom_loads: 用户指定的要查看有功无功和电流的负荷ID列表，例如 ['Load_1', 'Load_2']

    返回:
    pandas.DataFrame: 包含关键结果数据的DataFrame
    """
    try:
        # 设置中文字体支持
        plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans', 'Arial Unicode MS', 'sans-serif']
        plt.rcParams['axes.unicode_minus'] = False  # 正确显示负号

        # 连接到数据库
        conn = sqlite3.connect(database_path)
        cur = conn.cursor()

        # 测试连接
        result = cur.execute('SELECT SQLITE_VERSION()')
        version = result.fetchone()
        print(f"SQLite版本: {version[0]}")

        # 获取时间点信息
        cur.execute("SELECT ResultID, Time FROM TDTimeID ORDER BY TimeID;")
        time_points = cur.fetchall()
        time_data = [(rid, datetime.datetime.strptime(t, '%m-%d-%Y %H:%M:%S.%f'))
                     for rid, t in time_points]

        # 创建时间和结果ID的映射
        result_ids = [rid for rid, _ in time_data]
        timestamps = [t for _, t in time_data]

        # 打开输出文件
        with open(output_file, 'w', encoding='utf-8') as f:
            # 写入数据库信息头
            f.write("===== 时序潮流计算结果 =====\n")
            f.write(f"SQLite版本: {version[0]}\n")
            f.write(f"数据库路径: {database_path}\n\n")

            # 写入时间点信息
            f.write(f"总时间点数量: {len(time_points)}\n")
            f.write("时间点列表: " + ", ".join([t[1] for t in time_points[:10]]) + "...\n\n")

            # 1. 系统总体负荷和损耗随时间变化
            f.write("===== 系统总体负荷和损耗随时间变化 =====\n")
            cur.execute("""
                SELECT t.Time, s.TotalLoadMWPhA + s.TotalLoadMWPhB + s.TotalLoadMWPhC as TotalLoadMW,
                s.MWLossPhA + s.MWLossPhB + s.MWLossPhC as TotalLossMW
                FROM TDTimeID t 
                JOIN TDSysResult s ON t.ResultID = s.ResultID
                ORDER BY t.TimeID
            """)
            system_load_data = cur.fetchall()

            f.write("时间, 总负荷(MW), 总损耗(MW)\n")
            for time, load, loss in system_load_data:
                f.write(f"{time}, {load:.4f}, {loss:.4f}\n")
            f.write("\n")

            # 2. 获取系统最低电压随时间变化
            f.write("===== 系统最低电压随时间变化 =====\n")
            cur.execute("""
                SELECT t.Time, s.MinLNBusVPhA, s.MinLNBusVDeviceID
                FROM TDTimeID t 
                JOIN TDSysResult s ON t.ResultID = s.ResultID
                ORDER BY t.TimeID
            """)
            min_voltage_data = cur.fetchall()

            f.write("时间, 最低电压(%), 母线ID\n")
            for time, voltage, bus_id in min_voltage_data:
                f.write(f"{time}, {voltage:.4f}, {bus_id}\n")
            f.write("\n")

            # 3. 获取系统最大支路负载随时间变化
            f.write("===== 系统最大支路负载随时间变化 =====\n")
            cur.execute("""
                SELECT t.Time, s.MaxBranchLoadingPhA, s.MaxBranchLoadingDeviceID
                FROM TDTimeID t 
                JOIN TDSysResult s ON t.ResultID = s.ResultID
                ORDER BY t.TimeID
            """)
            max_loading_data = cur.fetchall()

            f.write("时间, 最大支路负载(%), 设备ID\n")
            for time, loading, device_id in max_loading_data:
                f.write(f"{time}, {loading:.4f}, {device_id}\n")
            f.write("\n")

            # 4. 获取母线电压和相角随时间变化
            # 处理用户指定的母线和关键母线
            buses_to_process = []
            if custom_buses:
                buses_to_process.extend(custom_buses)
            else:
                # 如果用户没有指定母线，则使用欠压母线
                cur.execute("""
                    SELECT DISTINCT DeviceID FROM TDAlert 
                    WHERE Condition='Under Voltage' AND AlertType='Critical'
                    LIMIT 5
                """)
                critical_buses = cur.fetchall()
                buses_to_process.extend([bus[0] for bus in critical_buses])

            # 确保buses_to_process中没有重复项
            buses_to_process = list(set(buses_to_process))

            # 存储母线电压和相角数据
            bus_voltages = {}
            bus_angles = {}

            if buses_to_process:
                f.write("===== 母线电压和相角随时间变化 =====\n")
                f.write(f"分析的母线: {', '.join(buses_to_process)}\n\n")

                # 为每个母线获取电压和相角随时间变化
                for bus_id in buses_to_process:
                    f.write(f"母线 {bus_id} 的电压和相角变化:\n")
                    f.write("时间, 电压(%), 相角(度)\n")

                    # 获取母线的内部ID
                    cur.execute(f"SELECT BusIID FROM TDBusInfo WHERE BusName=?", (bus_id,))
                    bus_iid_result = cur.fetchone()

                    if bus_iid_result:
                        bus_iid = bus_iid_result[0]

                        # 一次性获取所有时间点的数据
                        cur.execute("""
                            SELECT t.TimeID, t.ResultID, t.Time, b.VPhA, b.AngPhA 
                            FROM TDTimeID t
                            LEFT JOIN TDBusResult b ON t.ResultID = b.ResultID AND b.BusIID = ?
                            ORDER BY t.TimeID
                        """, (bus_iid,))

                        all_results = cur.fetchall()

                        # 初始化数据列表
                        voltages = []
                        angles = []
                        valid_data_count = 0

                        # 处理查询结果
                        for result in all_results:
                            time_id, result_id, time_str, voltage, angle = result

                            if voltage is not None:
                                voltages.append(voltage)
                                angles.append(angle)
                                valid_data_count += 1
                                f.write(f"{time_str}, {voltage:.4f}, {angle:.4f}\n")
                            else:
                                voltages.append(None)
                                angles.append(None)
                                f.write(f"{time_str}, 数据缺失, 数据缺失\n")

                        # 调试信息
                        print(f"母线 {bus_id}: 有效数据点 {valid_data_count}/{len(all_results)}")

                        # 存储数据用于绘图 - 确保长度与时间点一致
                        bus_voltages[bus_id] = voltages
                        bus_angles[bus_id] = angles
                    else:
                        f.write(f"未找到母线 {bus_id} 的信息\n")

                    f.write("\n")

            # 5.获取指定负荷信息
            loads_to_process = []
            if custom_loads:
                loads_to_process.extend(custom_loads)
                # 确保load_to_process中没有重复项

            loads_to_process = list(set(loads_to_process))
            #存储负荷的有功无功和
            load_active_power = {}
            load_reactive_power = {}
            load_current = {}

            if loads_to_process:
                f.write("===== 负荷有功无功和电流随时间变化 =====\n")
                f.write(f"分析的负荷: {', '.join(loads_to_process)}\n\n")

                # 为每个负荷获取有功无功和电流随时间变化
                for load_id in loads_to_process:
                    f.write(f"负荷 {load_id} 的有功功率、无功功率和电流变化:\n")
                    f.write("时间, 有功功率(MW), 无功功率(Mvar), 电流(A)\n")

                    # 获取负荷的内部ID
                    cur.execute(f"SELECT DeviceIID FROM TDOneTermDevicesInfo WHERE DeviceName=?", (load_id,))
                    load_iid_result = cur.fetchall()

                    if load_iid_result:
                        load_iid = load_iid_result[0]

                        # 确保 load_iid 是整数
                        load_iid = load_iid[0]  # 根据您之前提到的 DeviceIID

                        # 注意：在 SQLite 中，参数绑定需要使用元组形式，即使只有一个参数
                        cur.execute("""
                            SELECT t.TimeID, t.Time, t.ResultID, s.DeviceIID, s.TotalMWPhA, s.TotalMvarPhA, s.AmpPhA 
                            FROM TDTimeID t
                            LEFT JOIN TDSourceandLoadResult s ON t.ResultID = s.ResultID
                            WHERE s.DeviceIID = ?
                            ORDER BY t.TimeID
                        """, (load_iid,))  # 注意这里使用元组 (load_iid,)

                        results = cur.fetchall()

                        # 初始化数据列表
                        active_power_data = []
                        reactive_power_data = []
                        current_data = []
                        valid_data_count = 0

                        # 处理查询结果
                        for result in results:
                            time_id, time_str, result_id, device_iid, activepower, reactivepower, current = result
                            if activepower is not None:
                                active_power_data.append(activepower)
                                reactive_power_data.append(reactivepower)
                                current_data.append(current)
                                valid_data_count += 1
                                f.write(f"{time_str}, {activepower:.4f}, {reactivepower:.4f}, {current:.4f}\n")
                            else:
                                active_power_data.append(None)
                                reactive_power_data.append(None)
                                current_data.append(None)
                                f.write(f"{time_str}, 数据缺失, 数据缺失, 数据缺失\n")

                        # 调试信息
                        print(f"负荷 {load_id}: 有效数据点 {valid_data_count}/{len(results)}")

                        # 存储数据用于绘图 - 确保长度与时间点一致
                        load_active_power[load_id] = active_power_data
                        load_reactive_power[load_id] = reactive_power_data
                        load_current[load_id] = current_data
                    else:
                        f.write(f"未找到负荷 {load_id} 的信息\n")

                    f.write("\n")


            # 6. 获取事件信息
            f.write("===== 系统事件信息 =====\n")
            cur.execute("""
                SELECT Time, DeviceType, DeviceID, Action, ActionPercent
                FROM TDActions
                ORDER BY Time
            """)
            events_data = cur.fetchall()

            f.write("时间, 设备类型, 设备ID, 操作, 操作百分比\n")
            for time, device_type, device_id, action, action_percent in events_data:
                f.write(f"{time}, {device_type}, {device_id}, {action}, {action_percent}\n")

        print(f"时序潮流结果已成功导出到 {output_file}")

        # 准备绘图数据
        times = [datetime.datetime.strptime(t[1], '%m-%d-%Y %H:%M:%S.%f') for t in time_points]

        # 1. 系统总负荷和损耗图
        total_load = [row[1] for row in system_load_data]
        total_loss = [row[2] for row in system_load_data]

        # 2. 系统最低电压图
        min_voltages = [row[1] for row in min_voltage_data]

        # 3. 系统最大支路负载图
        max_loadings = [row[1] for row in max_loading_data]

        # 绘图 - 系统总览
        plt.figure(figsize=(15, 10))

        # 1. 系统总负荷和损耗图
        plt.subplot(2, 2, 1)
        plt.plot(times, total_load, 'b-', label='Total Load (MW)')
        plt.plot(times, total_loss, 'r-', label='Total Loss (MW)')
        plt.title('System Load and Loss vs Time')
        plt.xlabel('Time')
        plt.ylabel('Power (MW)')
        plt.legend()
        plt.grid(True)
        plt.xticks(rotation=45)
        plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))

        # 2. 系统最低电压图
        plt.subplot(2, 2, 2)
        plt.plot(times, min_voltages, 'g-')
        plt.axhline(y=90, color='r', linestyle='--', label='Lower Limit (90%)')
        plt.title('Minimum System Voltage vs Time')
        plt.xlabel('Time')
        plt.ylabel('Voltage (%)')
        plt.grid(True)
        plt.xticks(rotation=45)
        plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))
        plt.legend()

        # 3. 系统最大支路负载图
        plt.subplot(2, 2, 3)
        plt.plot(times, max_loadings, 'm-')
        plt.axhline(y=100, color='r', linestyle='--', label='Rated Value (100%)')
        plt.title('Maximum Branch Loading vs Time')
        plt.xlabel('Time')
        plt.ylabel('Loading (%)')
        plt.grid(True)
        plt.xticks(rotation=45)
        plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))
        plt.legend()

        # 4. 关键母线电压图
        plt.subplot(2, 2, 4)
        if bus_voltages:
            for bus_id, voltages in bus_voltages.items():
                # 只使用母线ID的数字部分作为标签
                label = bus_id.split('_')[1] if '_' in bus_id else bus_id
                plt.plot(times[:len(voltages)], voltages, label=f"Bus {label}")
            plt.axhline(y=90, color='r', linestyle='--', label='Lower Limit (90%)')
            plt.title('Bus Voltages vs Time')
            plt.xlabel('Time')
            plt.ylabel('Voltage (%)')
            plt.legend()
            plt.grid(True)
            plt.xticks(rotation=45)
            plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))

        plt.tight_layout()
        plt.savefig('system_overview.png', dpi=300)
        print("系统总览图表已保存为 system_overview.png")

        # 绘制母线电压和相角图 (单独的图表)
        if bus_voltages:
            plt.figure(figsize=(15, 10))

            # 电压图
            plt.subplot(2, 1, 1)
            for bus_id, voltages in bus_voltages.items():
                label = bus_id.split('_')[1] if '_' in bus_id else bus_id
                plt.plot(times[:len(voltages)], voltages, label=f"Bus {label}")

            plt.axhline(y=90, color='r', linestyle='--', label='Lower Limit (90%)')
            plt.title('Bus Voltages vs Time')
            plt.xlabel('Time')
            plt.ylabel('Voltage (%)')
            plt.legend()
            plt.grid(True)
            plt.xticks(rotation=45)
            plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))

            # 相角图
            plt.subplot(2, 1, 2)
            for bus_id, angles in bus_angles.items():
                label = bus_id.split('_')[1] if '_' in bus_id else bus_id
                plt.plot(times[:len(angles)], angles, label=f"Bus {label}")

            plt.title('Bus Angles vs Time')
            plt.xlabel('Time')
            plt.ylabel('Angle (degrees)')
            plt.legend()
            plt.grid(True)
            plt.xticks(rotation=45)
            plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))

            plt.tight_layout()
            plt.savefig('bus_voltage_angle.png', dpi=300)
            print("母线电压和相角图表已保存为 bus_voltage_angle.png")

        # 新增：绘制负荷有功功率、无功功率和电流图表
        if load_active_power:
            # 创建一个新的图表，包含三个子图：有功功率、无功功率和电流
            plt.figure(figsize=(15, 15))

            # 1. 有功功率图
            plt.subplot(3, 1, 1)
            for load_id, active_power in load_active_power.items():
                # 简化负荷ID作为标签
                label = load_id.split('_')[1] if '_' in load_id else load_id
                plt.plot(times[:len(active_power)], active_power, label=f"Load {label}")

            plt.title('Load Active Power vs Time')
            plt.xlabel('Time')
            plt.ylabel('Active Power (MW)')
            plt.legend()
            plt.grid(True)
            plt.xticks(rotation=45)
            plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))

            # 2. 无功功率图
            plt.subplot(3, 1, 2)
            for load_id, reactive_power in load_reactive_power.items():
                label = load_id.split('_')[1] if '_' in load_id else load_id
                plt.plot(times[:len(reactive_power)], reactive_power, label=f"Load {label}")

            plt.title('Load Reactive Power vs Time')
            plt.xlabel('Time')
            plt.ylabel('Reactive Power (Mvar)')
            plt.legend()
            plt.grid(True)
            plt.xticks(rotation=45)
            plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))

            # 3. 电流图
            plt.subplot(3, 1, 3)
            for load_id, current in load_current.items():
                label = load_id.split('_')[1] if '_' in load_id else load_id
                plt.plot(times[:len(current)], current, label=f"Load {label}")

            plt.title('Load Current vs Time')
            plt.xlabel('Time')
            plt.ylabel('Current (A)')
            plt.legend()
            plt.grid(True)
            plt.xticks(rotation=45)
            plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))

            plt.tight_layout()
            plt.savefig('load_characteristics.png', dpi=300)
            print("负荷特性图表已保存为 load_characteristics.png")

            # 额外创建一个功率因数图表
            plt.figure(figsize=(15, 8))
            for load_id in load_active_power.keys():
                if load_id in load_reactive_power:
                    active = load_active_power[load_id]
                    reactive = load_reactive_power[load_id]

                    # 计算视在功率和功率因数
                    apparent_power = []
                    power_factor = []

                    for p, q in zip(active, reactive):
                        if p is not None and q is not None:
                            s = np.sqrt(p ** 2 + q ** 2)
                            pf = p / s if s > 0 else 0
                            apparent_power.append(s)
                            power_factor.append(pf)
                        else:
                            apparent_power.append(None)
                            power_factor.append(None)

                    # 绘制功率因数
                    label = load_id.split('_')[1] if '_' in load_id else load_id
                    plt.plot(times[:len(power_factor)], power_factor, label=f"Load {label}")

            plt.title('Load Power Factor vs Time')
            plt.xlabel('Time')
            plt.ylabel('Power Factor')
            plt.legend()
            plt.grid(True)
            plt.xticks(rotation=45)
            plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))
            plt.tight_layout()
            plt.savefig('load_power_factor.png', dpi=300)
            print("负荷功率因数图表已保存为 load_power_factor.png")

        # 5. 事件标记图
        plt.figure(figsize=(15, 8))
        plt.plot(times, total_load, 'b-', label='Total Load (MW)')

        # 添加事件标记
        event_times = [datetime.datetime.strptime(event[0], '%m-%d-%Y %H:%M:%S.%f') for event in events_data]

        # 在事件发生时添加垂直线，使用简化的事件标签
        for i, event_time in enumerate(event_times):
            event = events_data[i]
            # 简化事件标签，避免中文问题
            action_label = f"{event[3]} {event[2].split('_')[-1] if '_' in event[2] else event[2]}"
            plt.axvline(x=event_time, color='r', linestyle='--', alpha=0.5)
            # 添加事件标签，交替放置在上下位置以避免重叠
            y_pos = max(total_load) * (0.9 if i % 2 == 0 else 0.8)
            plt.text(event_time, y_pos, action_label, rotation=90, fontsize=8)

        plt.title('System Load and Events Correlation')
        plt.xlabel('Time')
        plt.ylabel('Power (MW)')
        plt.grid(True)
        plt.xticks(rotation=45)
        plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))
        plt.tight_layout()
        plt.savefig('load_events_correlation.png', dpi=300)
        print("负荷与事件关系图已保存为 load_events_correlation.png")

        plt.close('all')

        # 创建一个包含关键结果的DataFrame用于导出到Excel
        result_data = {
            'Time': [t[1] for t in time_points],
            'Total_Load_MW': total_load,
            'Total_Loss_MW': total_loss,
            'Min_Voltage_Percent': min_voltages,
            'Max_Branch_Loading_Percent': max_loadings
        }

        # 添加母线电压和相角数据
        for bus_id, voltages in bus_voltages.items():
            # 确保所有数据长度一致
            padded_voltages = voltages + [None] * (len(time_points) - len(voltages))
            result_data[f'Bus_{bus_id}_Voltage'] = padded_voltages

            # 添加相角数据
            if bus_id in bus_angles:
                angles = bus_angles[bus_id]
                padded_angles = angles + [None] * (len(time_points) - len(angles))
                result_data[f'Bus_{bus_id}_Angle'] = padded_angles

        # 新增：添加负荷数据到Excel导出
        for load_id, active_power in load_active_power.items():
            # 确保所有数据长度一致
            padded_active = active_power + [None] * (len(time_points) - len(active_power))
            result_data[f'Load_{load_id}_MW'] = padded_active

            # 添加无功功率数据
            if load_id in load_reactive_power:
                reactive_power = load_reactive_power[load_id]
                padded_reactive = reactive_power + [None] * (len(time_points) - len(reactive_power))
                result_data[f'Load_{load_id}_Mvar'] = padded_reactive

            # 添加电流数据
            if load_id in load_current:
                current = load_current[load_id]
                padded_current = current + [None] * (len(time_points) - len(current))
                result_data[f'Load_{load_id}_A'] = padded_current

        # 创建DataFrame
        result_df = pd.DataFrame(result_data)

        # 导出到Excel
        excel_file = "time_series_results.xlsx"
        result_df.to_excel(excel_file, index=False)
        print(f"结果数据已导出到Excel文件: {excel_file}")

        # 返回结果DataFrame
        return result_df

    except Exception as e:
        print(f"错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return None
    finally:
        # 关闭数据库连接
        if 'conn' in locals() and conn:
            conn.close()
            print("数据库连接已关闭")
