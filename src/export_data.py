"""
   Exp
"""

import sqlite3
import pandas as pd


def export_report(database_path):
    try:
        conn = sqlite3.connect(database_path)
        cur = conn.cursor()

        # 执行查询
        print("\n执行数据查询...")
        cur.execute("""
            SELECT IDBus, NomlkV, VMagA, VMagB, VMagC, VAngA, VAngB, VAngC
            FROM IBusLF3PH
            WHERE NomlkV > 0;
        """)
        rows = cur.fetchall()

        if not rows:
            raise Exception("未找到符合条件的数据")

        # 计算平均相电压和平均相角
        result_bus = {
            "bus_ID": [],
            "volt_mag": [],
            "volt_ang": []
        }

        for row in rows:
            bus_id, nominal_kv, vmag_a, vmag_b, vmag_c, vang_a, vang_b, vang_c = row
            # 计算三相电压的平均值（百分比转换为标幺值）
            avg_voltage = (vmag_a + vmag_b + vmag_c) / 300 if all(
                v is not None for v in [vmag_a, vmag_b, vmag_c]) else None
            # 计算三相角度的平均值
            avg_angle = (vang_a + vang_b + vang_c) / 300 if all(a is not None for a in [vang_a, vang_b, vang_c]) else None

            result_bus["bus_ID"].append(bus_id)
            result_bus["volt_mag"].append(avg_voltage)
            result_bus["volt_ang"].append(avg_angle)

        result_bus = pd.DataFrame(result_bus)
        print(f"\n成功获取数据，共 {len(result_bus)} 条记录")
        return result_bus

    except sqlite3.Error as e:
        print(f"SQLite 错误: {e}")
        raise
    except Exception as e:
        print(f"发生错误: {e}")
        raise
    finally:
        if 'cur' in locals():
            cur.close()
        if 'conn' in locals():
            conn.close()


