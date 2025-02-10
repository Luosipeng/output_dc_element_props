import etap.api
import json
import pandas as pd
import xml.etree.ElementTree as ET
from configuration.configuration import base_address
from src.export_data import export_report


class ETAPExporter:
    """ETAP 项目数据导出工具"""

    def __init__(self, base_address):
        """初始化ETAP连接并验证"""
        self._connect_etap(base_address)
        self._verify_connection()
        self.path_result = None

    def _connect_etap(self, address):
        """建立ETAP连接"""
        print("Initializing ETAP connection...")
        self.etap = etap.api.connect(address)

    def _verify_connection(self):
        """验证连接状态"""
        print("Verifying connection...")
        print("File paths:", json.loads(self.etap.application.filepaths()))
        print("PID:", self.etap.application.pid())
        print("DataHub status:", self.etap.application.ping())

    def run_analysis(self, study_type, **kwargs):
        """执行分析任务"""
        study_runner = {
            'LF': self.etap.studies.runLF,
            'ULF': self.etap.studies.runULF,
            'SC': self._run_short_circuit
        }
        return study_runner[study_type](**kwargs)

    def _run_short_circuit(self, **kwargs):
        """短路计算专用方法"""
        kwargs.update({'output_report': "SC", 'study_case': "SC-A"})
        return self.etap.studies.runSC(studyType="IEC Transient Fault Current", **kwargs)

    def export_project_data(self):
        """主数据导出流程"""
        inverter_df = self._process_inverters()
        dc_load_df = self._process_dc_loads()
        battery_df = self._process_batteries()
        self._save_to_excel({"INVERTER": inverter_df, "DCLUMPLOAD": dc_load_df, "BATTERY": battery_df})

    def _process_inverters(self):
        """处理逆变器数据"""
        print("Processing inverter data...")
        root = ET.fromstring(self.etap.projectdata.getallelementdata("INVERTER"))
        return self._build_inverter_df(root.findall("INVERTER"))

    def _build_inverter_df(self, inverters):
        """构建逆变器DataFrame"""
        # 基础属性提取
        base_attributes = {
            'IID': 'ID',
            'Felement': 'BusID',
            'Telement': 'CZNetwork',
            'ACkV': 'KV',
            'DCV': 'DcV',
            'KVA': 'KVA',
            'DCKW': 'DckW',
            'DCEff': 'DcPercentEFF',
            'Vref': 'Vref',
            'Smax': 'KVAMax',
            'Pmax': 'KWMax',
            'Qmax': 'KvarMax'
        }

        data = {col: [] for col in ['ID'] + list(base_attributes.keys())}
        pv_data = {f"{t}_{v}{i}": []
                   for t in ['generate', 'charge']
                   for v in ['V', 'P']
                   for i in range(1, 4)}

        for idx, inv in enumerate(inverters, 1):
            data['ID'].append(idx)
            for df_col, attr in base_attributes.items():
                data[df_col].append(inv.attrib.get(attr))

            # PV曲线处理
            self._process_pv_curves(inv, pv_data, idx)

        return pd.DataFrame({**data, **pv_data})

    def _process_pv_curves(self, inv, data_dict, idx):
        """处理PV曲线数据"""
        if user_points := inv.attrib.get("UserDefPoints"):
            points = ET.fromstring(user_points).findall(".//UserDefPVPoint")
            for i, point in enumerate(points[:3], 1):  # 只取前3个点
                data_dict[f'generate_V{i}'].append(point.attrib.get("percentVGenerating"))
                data_dict[f'generate_P{i}'].append(point.attrib.get("percentPGenerating"))
                data_dict[f'charge_V{i}'].append(point.attrib.get("percentVCharging"))
                data_dict[f'charge_P{i}'].append(point.attrib.get("percentPCharging"))
        else:
            for key in data_dict:
                data_dict[key].append(None)

    def _process_dc_loads(self):
        """处理直流负载数据"""
        print("Processing DC load data...")
        root = ET.fromstring(self.etap.projectdata.getallelementdata("DCLUMPLOAD"))
        return pd.DataFrame([
            {
                "ID": idx,
                "IID": ld.attrib.get("ID"),
                "BUS_I": ld.attrib.get("Bus"),
                "RatedV": ld.attrib.get("DCV"),
                "RatedKW": ld.attrib.get("KW")
            }
            for idx, ld in enumerate(root.findall("DCLUMPLOAD"), 1)
        ])

    def _process_batteries(self):
        """处理电池储能系统数据"""
        print("Processing battery data...")
        root = ET.fromstring(self.etap.projectdata.getallelementdata("BATTERY"))
        return pd.DataFrame([
            {
                "ID": idx,
                "IID": bat.attrib.get("ID"),
                "BusID": bat.attrib.get("Bus"),
                "Cells": bat.attrib.get("NrOfCells"),
                "Packs": bat.attrib.get("NoOfPacks"),
                "Strings": bat.attrib.get("NrOfStrings"),
            }
            for idx, bat in enumerate(root.findall("BATTERY"), 1)
        ])

    def _process_dc_cables(self):
        """处理直流电缆数据"""
        print("Processing DC cable data...")
    def _save_to_excel(self, data_dict):
        """保存数据到Excel"""
        print("Exporting to Excel...")
        with pd.ExcelWriter("parameters.xlsx") as writer:
            for sheet_name, df in data_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print("Export completed.")



if __name__ == "__main__":
    exporter = ETAPExporter(base_address)

    # 示例执行分析
    # exporter.run_analysis('LF',
    #     revision_name=revision_name,
    #     config_name=config_name,
    #     study_case=study_case,
    #     presentation=presentation,
    #     output_report=output_report,
    #     get_online_data=get_online_data
    # )

    exporter.export_project_data()