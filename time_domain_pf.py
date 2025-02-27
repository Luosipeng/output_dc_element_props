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
    # configuration/configuration.py
    base_address = "http://localhost:60000"  # 修改为正确的端口号
    revision_name = "Base"
    config_name = "Normal"  # API 返回的配置名称
    study_case = "ULF"
    presentation = "OLV1"
    output_report = "Results"
    get_online_data = False

    main = Main(base_address)
    studies = main.etap.studies.runULF(revision_name, config_name, study_case, presentation, output_report)
    result_dict=json.loads(studies)
    P_value=result_dict["ReportPath"]
    print(P_value)
    # a='<string xmlns="http://schemas.microsoft.com/2003/10/Serialization/">'
    # b='</string>'
    # result=a+P_value+b
    # report_path=ET.fromstring(result).text
    # print(report_path)
    result_bus = export_report(P_value)
    print(result_bus)
    print("Done.")