#This function is aimed to output the essential parameter of etap project to an excel file
#Usage
#Step1: open etap project and start datahub
#Step2: run this file

import etap.api
import json
import etap
import pandas as pd

from configuration.configuration import base_address, revision_name, config_name, study_case, presentation, output_report, get_online_data,online_config_only,what_if_commands

from src.export_data import export_report
from src.export_result import export_time_series_power_flow
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

    def run_time_domain_load_flow(self, revision_name, config_name, study_case, presentation, output_report, get_online_data,online_config_only, what_if_commands):
        print("Run time domain load flow...")
        response = self.etap.studies.runTDLF(revision_name, config_name, study_case, presentation, output_report, get_online_data, online_config_only, what_if_commands)
        print("Save time domain load flow result...")
        paths = json.loads(response)
        print(paths)
        self.path_result = paths["ReportPath"]
    def export_report(self):
        result_bus = export_report(self.path_result)
        return result_bus
    def change_parameters(self, elementType, elementName, fieldName, value):
        print("Change parameters...")
        self.etap.projectdata.setelementprop(elementType, elementName, fieldName, value)
    def export_output(self):
        custom_buses = ['BUS_CNODE_JCT__1333', 'BUS_都天元居开闭所（自）_220', 'BUS_kV爱涛线腾亚环网柜_207']
        custom_loads = ["LOAD_南京原野制衣有限公司_939"]
        result_bus = export_time_series_power_flow(self.path_result,output_file="custom_results.txt",custom_buses=custom_buses,custom_loads=custom_loads)
        return result_bus

    # def run_power_flow(self):

if __name__ == "__main__":
    # configuration/configuration.py
    main = Main(base_address)
    main.run_time_domain_load_flow(revision_name, config_name, study_case, presentation, output_report, get_online_data, online_config_only, what_if_commands)
    result = main.export_output()

    print("Export report to Excel file")
    result.to_excel("result.xlsx", index=False)

    print("Done.")