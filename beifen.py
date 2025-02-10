#This function is aimed to output the essential parameter of etap project to an excel file
#Usage
#Step1: open etap project and start datahub
#Step2: run this file

import etap.api
import json
import etap
import pandas as pd

from configuration.configuration import base_address, revision_name, config_name, study_case, presentation, output_report, get_online_data

from src.export_data import export_report
import xml.etree.ElementTree as ET
class Main():
    def __init__(self,base_address):
        print("Test connection...")
        e = etap.api.connect(base_address)
        response = e.application.filepaths()
        print(response)

        response = e.application.pid()
        print(response)
        # ping
        print("Test DataHub...")
        ping_result = e.application.ping()
        print(str(ping_result))

        # obtain the project path
        response = e.application.filepaths()
        paths = json.loads(response)
        # convert the str to dictionary
        self.paths = paths
        self.etap = e
    def run_power_flow(self, revision_name, config_name, study_case, presentation, output_report, get_online_data):
        print("Run power flow...")

        response = self.etap.studies.runLF(revision_name, config_name, study_case, presentation, output_report, get_online_data)
        print("Save power flow result...")
        paths = json.loads(response)
        self.path_result = paths["ReportPath"]
    def run_unbalanced_power_flow(self, revision_name, config_name, study_case, presentation, output_report, get_online_data):
        print("Run unbalanced power flow...")
        response = self.etap.studies.runULF(revision_name, config_name, study_case, presentation, output_report, get_online_data)
        print("Save unbalanced power flow result...")
        paths = json.loads(response)
        self.path_result = paths["ReportPath"]
    def run_sc_cal(self, revision_name, config_name, study_case, presentation, output_report, get_online_data):
        print("Run short circuit calculation...")
        output_report = "SC"
        study_case = "SC-A"
        studyType = "IEC Transient Fault Current"
        response = self.etap.studies.runSC(revision_name, config_name, study_case, presentation, output_report, studyType, get_online_data)
        print("Save short circuit flow result...")
        paths = json.loads(response)
        self.path_result = paths["ReportPath"]
    def export_report(self):
        result_bus = export_report(self.path_result)
        return result_bus
    def change_parameters(self, elementType, elementName, fieldName, value):
        print("Change parameters...")
        self.etap.projectdata.setelementprop(elementType, elementName, fieldName, value)

    # def run_power_flow(self):

if __name__ == "__main__":
    main = Main(base_address)
    inverter_data=main.etap.projectdata.getallelementdata("INVERTER")
    root = ET.fromstring(inverter_data)
    count = root.attrib.get("Count")
    ID= list(range(1, int(count) + 1))

    # 获取所有 INVERTER 元素
    inverters = root.findall("INVERTER")

    # 提取设定的交流侧输出功率
    genCat0ACkW_values = [inv.attrib.get("GenCat0ACkW") for inv in inverters]
    genCat0kVar_values = [inv.attrib.get("GenCat0kVar") for inv in inverters]

    #提取逆变器ID
    ID_values = [inv.attrib.get("ID") for inv in inverters]

    #提取逆变器连接元素
    connected_Felement=[inv.attrib.get("BusID") for inv in inverters]
    connected_Telement = [inv.attrib.get("CZNetwork") for inv in inverters]

    #提取逆变器的效率因素
    dcPercentEff=[inv.attrib.get("DcPercentEFF") for inv in inverters]

    #提取额定值
    ratedKV=[inv.attrib.get("KV") for inv in inverters]
    rateddcV=[inv.attrib.get("DcV") for inv in inverters]

    #提取额定功率
    ratedKVA=[inv.attrib.get("KVA") for inv in inverters]
    ratedDCKW=[inv.attrib.get("DckW") for inv in inverters]

    #参考电压
    vref=[inv.attrib.get("Vref") for inv in inverters]

    #功率限幅
    smax=[inv.attrib.get("KVAMax") for inv in inverters]
    pmax=[inv.attrib.get("KWMax") for inv in inverters]
    qmax=[inv.attrib.get("KvarMax") for inv in inverters]


    #获取自定义PV曲线
    # 提取 UserDefPoints 并解析 UserDefPVPoints
    generating_V_values1 = []
    generating_P_values1 = []
    generating_V_values2 = []
    generating_P_values2 = []
    generating_V_values3 = []
    generating_P_values3 = []

    charging_V_values1 = []
    charging_P_values1 = []
    charging_V_values2 = []
    charging_P_values2 = []
    charging_V_values3 = []
    charging_P_values3 = []
    for inv in inverters:
        userDefPoints = inv.attrib.get("UserDefPoints")

        if userDefPoints:  # 如果 UserDefPoints 存在
            # 解析 UserDefPoints
            userDefPoints_tree = ET.ElementTree(ET.fromstring(userDefPoints))
            userDefPVPoints = userDefPoints_tree.findall(".//UserDefPVPoint")

            # 定义两个列表来存储生成和充电的部分
            generating_V_values = []
            generating_P_values = []
            charging_V_values = []
            charging_P_values = []

            # 遍历所有 UserDefPVPoint 元素并提取数据
            for point in userDefPVPoints:
                generating_V_values.append(point.attrib.get("percentVGenerating"))
                generating_P_values.append(point.attrib.get("percentPGenerating"))
                charging_V_values.append(point.attrib.get("percentVCharging"))
                charging_P_values.append(point.attrib.get("percentPCharging"))

            generating_V_values1.append(generating_V_values[0])
            generating_V_values2.append(generating_V_values[1])
            generating_V_values3.append(generating_V_values[2])
            generating_P_values1.append(generating_P_values[0])
            generating_P_values2.append(generating_P_values[1])
            generating_P_values3.append(generating_P_values[2])

            charging_V_values1.append(charging_V_values[0])
            charging_V_values2.append(charging_V_values[1])
            charging_V_values3.append(charging_V_values[2])
            charging_P_values1.append(charging_P_values[0])
            charging_P_values2.append(charging_P_values[1])
            charging_P_values3.append(charging_P_values[2])

    #obtain the DC lumped lad parameter
    dclumpedload_data = main.etap.projectdata.getallelementdata("DCLUMPLOAD")
    root = ET.fromstring(dclumpedload_data)
    count = root.attrib.get("Count")
    IDld = list(range(1, int(count) + 1))
    # 获取所有 INVERTER 元素
    dclumpedload = root.findall("DCLUMPLOAD")
    loadiid=[ld.attrib.get("ID") for ld in dclumpedload]
    #提取连接节点信息
    connectedbus_load = [ld.attrib.get("Bus") for ld in dclumpedload]
    #提取额定电压
    ratedloadV=[ld.attrib.get("DCV") for ld in dclumpedload]

    ratedloadKW=[ld.attrib.get("KW") for ld in dclumpedload]

    print(IDld)
    print(connectedbus_load)
    print(ratedloadV)
    print(ratedloadKW)
    print(loadiid)
    dl=pd.DataFrame({
        "ID": IDld,
        "IID": loadiid,
        "BUS_I": connectedbus_load,
        "RatedV": ratedloadV,
        "RatedKW": ratedloadKW,
    })

    # Create a DataFrame to store the results
    df = pd.DataFrame({
        'ID': ID,
        'IID': ID_values,
        'Felement': connected_Felement,
        'Telement': connected_Telement,
        'ACkV': ratedKV,
        'DCV': rateddcV,
        'KVA': ratedKVA,
        'DCKW': ratedDCKW,
        'DCEff': dcPercentEff,
        'Vref': vref,
        'Smax': smax,
        'Pmax': pmax,
        'Qmax': qmax,
        'generate_V1': generating_V_values1,
        'generate_P1': generating_P_values1,
        'generate_V2': generating_V_values2,
        'generate_P2': generating_P_values2,
        'generate_V3': generating_V_values3,
        'generate_P3': generating_P_values3,
        'charging_V1': charging_V_values1,
        'charging_P1': charging_P_values1,
        'charging_V2': charging_V_values2,
        'charging_P2': charging_P_values2,
        'charging_V3': charging_V_values3,
        'charging_P3': charging_P_values3,
    })

    # Save the results to an Excel file
    print("Export report to Excel file")
    with pd.ExcelWriter("parameters.xlsx") as writer:
        df.to_excel(writer, sheet_name="INVERTER", index=False)
        dl.to_excel(writer, sheet_name="DCLUMPLOAD", index=False)
    print("Done.")