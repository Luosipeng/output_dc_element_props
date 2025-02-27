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
        # 创建一个字典来存储成功处理的数据框
        processed_data = {}

        # 定义要处理的元件及其对应的处理方法
        components = {
            "INVERTER": self._process_inverters,
            "DCLUMPLOAD": self._process_dc_loads,
            "BATTERY": self._process_batteries,
            "DCCABLE": self._process_dc_cables,
            "DCBUS": self._process_dc_buses,
            "DCIMPEDANCE": self._process_dc_impedances
        }

        # 遍历处理每个元件
        for component_name, process_func in components.items():
            try:
                df = process_func()
                if df is not None and not df.empty:
                    processed_data[component_name] = df
                else:
                    print(f"No data found for {component_name}, skipping...")
            except Exception as e:
                print(f"Error processing {component_name}: {str(e)}, skipping...")
                continue

        # 只保存成功处理的数据
        if processed_data:
            self._save_to_excel(processed_data)
        else:
            print("No data was successfully processed.")

    def _process_component_data(self, component_type):
        """通用的元件数据处理函数"""
        try:
            root = ET.fromstring(self.etap.projectdata.getallelementdata(component_type))
            if root is None:
                print(f"No {component_type} data found.")
                return None
            return root
        except Exception as e:
            print(f"Error getting {component_type} data: {str(e)}")
            return None

    def _process_inverters(self):
        """处理逆变器数据"""
        print("Processing inverter data...")
        try:
            root = self._process_component_data("INVERTER")
            if root is None:
                return pd.DataFrame()

            inverters = root.findall("INVERTER")
            if not inverters:
                print("No inverter elements found.")
                return pd.DataFrame()

            return self._build_inverter_df(inverters)
        except Exception as e:
            print(f"Error in processing inverters: {str(e)}")
            return pd.DataFrame()

    def _build_inverter_df(self, inverters):
        """构建逆变器DataFrame"""
        try:
            # 基础属性提取
            base_attributes = {
                'IID': 'ID',
                'InService': 'InService',
                'STATE': 'InServiceState',
                'Felement': 'BusID',
                'Telement': 'CZNetwork',
                'ACkV': 'KV',
                'DCV': 'DcV',
                'KVA': 'KVA',
                'DCKW': 'DckW',
                'DCEff': 'DcPercentEFF',
                'ACkW': 'GenCat0ACkW',
                'ACkVar': 'GenCat0kVar',
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
                self._process_pv_curves(inv, pv_data)

            return pd.DataFrame({**data, **pv_data})
        except Exception as e:
            print(f"Error in building inverter DataFrame: {str(e)}")
            return pd.DataFrame()

    def _process_pv_curves(self, inv, data_dict):
        """处理PV曲线数据"""
        try:
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
        except Exception as e:
            print(f"Error in processing PV curves: {str(e)}")
            for key in data_dict:
                data_dict[key].append(None)

    def _process_dc_loads(self):
        """处理直流负载数据"""
        print("Processing DC load data...")
        try:
            root = self._process_component_data("DCLUMPLOAD")
            if root is None:
                return pd.DataFrame()

            loads = root.findall("DCLUMPLOAD")
            if not loads:
                print("No DC load elements found.")
                return pd.DataFrame()

            return pd.DataFrame([
                {
                    "ID": idx,
                    "IID": ld.attrib.get("ID"),
                    "InService": ld.attrib.get("InService"),
                    "STATE": ld.attrib.get("InServiceState"),
                    "BUS_I": ld.attrib.get("Bus"),
                    "RatedV": ld.attrib.get("DCV"),
                    "RatedKW": ld.attrib.get("KW"),
                    "Percent_P": ld.attrib.get("MTLoadPercent"),
                    "Percent_Z": ld.attrib.get("StaticLoadPercent"),
                }
                for idx, ld in enumerate(loads, 1)
            ])
        except Exception as e:
            print(f"Error in processing DC loads: {str(e)}")
            return pd.DataFrame()

    def _process_batteries(self):
        """处理电池储能系统数据"""
        print("Processing battery data...")
        try:
            root = self._process_component_data("BATTERY")
            if root is None:
                return pd.DataFrame()

            batteries = root.findall("BATTERY")
            if not batteries:
                print("No battery elements found.")
                return pd.DataFrame()

            return pd.DataFrame([
                {
                    "ID": idx,
                    "IID": bat.attrib.get("ID"),
                    "InService": bat.attrib.get("InService"),
                    "STATE": bat.attrib.get("InServiceState"),
                    "BusID": bat.attrib.get("Bus"),
                    "Cells": bat.attrib.get("NrOfCells"),
                    "Packs": bat.attrib.get("NoOfPacks"),
                    "Strings": bat.attrib.get("NrOfStrings"),
                }
                for idx, bat in enumerate(batteries, 1)
            ])
        except Exception as e:
            print(f"Error in processing batteries: {str(e)}")
            return pd.DataFrame()

    def _process_dc_cables(self):
        """处理直流电缆数据"""
        print("Processing DC cable data...")
        try:
            root = self._process_component_data("CABLE")
            if root is None:
                return pd.DataFrame()

            cables = root.findall("CABLE")
            if not cables:
                print("No DC cable elements found.")
                return pd.DataFrame()

            return pd.DataFrame([
                {
                    "ID": idx,
                    "IID": cable.attrib.get("ID"),
                    "InService": cable.attrib.get("InService"),
                    "STATE": cable.attrib.get("InServiceState"),
                    "FromBus": cable.attrib.get("FromBus"),
                    "ToBus": cable.attrib.get("ToBus"),
                    "LENGTH": cable.attrib.get("LengthValue"),
                    "LENGTH_Unit": cable.attrib.get("ImpedanceUnits"),
                    "OhmsPerLengthValue": cable.attrib.get("OhmsPerLengthValue"),
                    "OhmsPerLengthUnit": cable.attrib.get("OhmsPerLengthUnit"),
                    "RPosValue": cable.attrib.get("RPosValue"),
                    "XPosValue": cable.attrib.get("XPosValue"),
                }
                for idx, cable in enumerate(cables, 1)
            ])
        except Exception as e:
            print(f"Error in processing DC cables: {str(e)}")
            return pd.DataFrame()

    def _process_dc_buses(self):
        """处理直流母线数据"""
        print("Processing DC bus data...")
        try:
            root = self._process_component_data("DCBUS")
            if root is None:
                return pd.DataFrame()

            buses = root.findall("DCBUS")
            if not buses:
                print("No DC bus elements found.")
                return pd.DataFrame()

            return pd.DataFrame([
                {
                    "ID": idx,
                    "IID": bus.attrib.get("ID"),
                    "NominalV": bus.attrib.get("NominalV"),
                    "InService": bus.attrib.get("InService"),
                    "STATE": bus.attrib.get("InServiceState"),
                }
                for idx, bus in enumerate(buses, 1)
            ])
        except Exception as e:
            print(f"Error in processing DC buses: {str(e)}")
            return pd.DataFrame()

    def _process_dc_impedances(self):
        """处理直流阻抗数据"""
        print("Processing DC impedance data...")
        try:
            root = self._process_component_data("DCIMPEDANCE")
            if root is None:
                return pd.DataFrame()

            impedances = root.findall("DCIMPEDANCE")
            if not impedances:
                print("No DC impedance elements found.")
                return pd.DataFrame()

            return pd.DataFrame([
                {
                    "ID": idx,
                    "IID": imp.attrib.get("ID"),
                    "InService": imp.attrib.get("InService"),
                    "STATE": imp.attrib.get("InServiceState"),
                    "FromBus": imp.attrib.get("FromBus"),
                    "ToBus": imp.attrib.get("ToBus"),
                    "RValue": imp.attrib.get("RValue"),
                    "LValue": imp.attrib.get("LValue"),
                }
                for idx, imp in enumerate(impedances, 1)
            ])
        except Exception as e:
            print(f"Error in processing DC impedances: {str(e)}")
            return pd.DataFrame()

    def _save_to_excel(self, data_dict):
        """保存数据到Excel"""
        try:
            print("Exporting to Excel...")
            with pd.ExcelWriter("parameters.xlsx") as writer:
                for sheet_name, df in data_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print("Export completed successfully.")
        except Exception as e:
            print(f"Error saving to Excel: {str(e)}")


if __name__ == "__main__":
    exporter = ETAPExporter(base_address)
    exporter.export_project_data()
