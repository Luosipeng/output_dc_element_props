import etap.api
import json
import etap
import pandas as pd

from configuration.configuration import base_address, revision_name, config_name, study_case, presentation, output_report, get_online_data

from src.export_pfdata import export_pfreport
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
    def export_pfreport(self):
        result_bus = export_pfreport(self.path_result)
        return result_bus
    def change_parameters(self, elementType, elementName, fieldName, value):
        print("Change parameters...")
        self.etap.projectdata.setelementprop(elementType, elementName, fieldName, value)

if __name__ == "__main__":
    main = Main(base_address)
    # obtain the data
    # Prepare lists to store the results
    x_array = []
    bus_ID1_array = []
    bus_ID2_array = []
    volt_mag1_array = []
    volt_mag2_array = []
    volt_ang1_array = []
    volt_ang2_array = []
    #Create a loop to test ten cases
    start_value=40
    iterations=10
    for i in range(iterations):
        current_value = start_value + (i * 10)
        current_value_str = str(current_value)
        #change the parameter:input its element type, name, filed and value
        # main.change_parameters("XFORM2W", "T2", "AnsiPosXR", current_value_str)
        main.change_parameters("LUMPEDLOAD", "Lump1", "MVA", current_value_str)
        # main.change_parameters("STLOAD", "Load4", "KVA", current_value_str)
        # main.change_parameters("XLINE", "Line2", "Length", current_value_str)
        #Run the power flow analysis
        main.run_power_flow(revision_name, config_name, study_case, presentation, output_report, get_online_data)
        # main.run_unbalanced_power_flow(revision_name, config_name, study_case, presentation, output_report, get_online_data)
        # export result to excel file
        result = main.export_pfreport()
        # print(result.volt_mag.values[0])
        x_array.append(current_value)  # MVA values
        bus_ID1_array.append(result.bus_ID.values[0])
        bus_ID2_array.append(result.bus_ID.values[1])
        volt_mag1_array.append(result.volt_mag.values[0])
        volt_mag2_array.append(result.volt_mag.values[1])
        volt_ang1_array.append(result.volt_ang.values[0])
        volt_ang2_array.append(result.volt_ang.values[1])

    # Create a DataFrame to store the results
    df = pd.DataFrame({
        'MVA Value': x_array,
        'bus_ID1': bus_ID1_array,
        'bus_ID2': bus_ID2_array,
        'volt1_mag': volt_mag1_array,
        'volt2_mag': volt_mag2_array,
        'volt1_ang': volt_ang1_array,
        'volt2_ang': volt_ang2_array,
    })

    # Save the results to an Excel file
    print("Export report to Excel file")
    df.to_excel("result.xlsx", index=False)
    print("Done.")