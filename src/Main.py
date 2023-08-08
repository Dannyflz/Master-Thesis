import pypsa
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


def xlsx_to_csv(xlsx_file, output_folder, selected_sheets=None):
    # Einlesen der XLSX-Datei
    xls = pd.ExcelFile(xlsx_file)
    
    # Liste der Tabellenblätter
    sheet_names = xls.sheet_names

    # Wenn keine ausgewählten Blätter übergeben wurden, exportiere alle Blätter
    if selected_sheets is None:
        selected_sheets = sheet_names

    # Überprüfe, ob die ausgewählten Blätter in der XLSX-Datei vorhanden sind
    invalid_sheets = set(selected_sheets) - set(sheet_names)
    if invalid_sheets:
        raise ValueError(f"Ungültige Tabellenblätter ausgewählt: {', '.join(invalid_sheets)}")

    # Für jedes ausgewählte Tabellenblatt eine CSV-Datei erstellen
    for sheet_name in selected_sheets:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Convert the sheet_name to lowercase for the CSV filename
        sheet_name_lower = sheet_name.lower()
        
        csv_file = os.path.join(output_folder, f"{sheet_name_lower}.csv")
        df.to_csv(csv_file, index=False, mode="w")
        
        print(f"Erstelle {csv_file}.")

def Inputs_dataframes (xlsx_file, selected_sheets_techno = None):
    
    # Erstelle ein leeres Dictionary, um die DataFrames für jede Technologie zu speichern
    dataframes_dict = {}
    for sheet_name in selected_sheets_techno:
        df = pd.read_excel(xlsx_file, sheet_name=sheet_name)
        dataframes_dict[sheet_name] = df
    return dataframes_dict

def create_dataframes_from_csv(csv_folder, selected_sheets):
    # Create an empty dictionary to store the DataFrames
    dataframes_dict_nw = {}
    # For each CSV file, create a DataFrame and store it in the dictionary
    for sheet_name in selected_sheets:
        csv_file_path = os.path.join(csv_folder, f"{sheet_name}.csv")
        df = pd.read_csv(csv_file_path)
        # Capitalize the variable name
        sheet_name_capitalized = sheet_name.capitalize()
        dataframes_dict_nw[sheet_name_capitalized] = df

    return dataframes_dict_nw

def import_components_from_dataframes(dataframes_dict_nw, network):
    
    # Import "Buses" components
    if "Buses" in dataframes_dict_nw:
        buses_df = dataframes_dict_nw["Buses"]
        buses_df.set_index("name", inplace=True)  # Set "name" column as the index
        network.import_components_from_dataframe(buses_df, "Bus")
        
    # Import "Carriers" components
    if "Carriers" in dataframes_dict_nw:
        carriers_df = dataframes_dict_nw["Carriers"]
        carriers_df.set_index("name", inplace=True)  # Set "name" column as the index
        network.import_components_from_dataframe(carriers_df, "Carrier")
        
    # Import "Generators" components
    if "Generators" in dataframes_dict_nw:
        generators_df = dataframes_dict_nw["Generators"]
        generators_df.set_index("name", inplace=True)  # Set "name" column as the index
        network.import_components_from_dataframe(generators_df, "Generator")
    
    # Import "Links" components
    if "Links" in dataframes_dict_nw:
        links_df = dataframes_dict_nw["Links"]
        links_df.set_index("name", inplace=True)  # Set "name" column as the index
        network.import_components_from_dataframe(links_df, "Link")
    
    # Import "Loads" components
    if "Loads" in dataframes_dict_nw:
        loads_df = dataframes_dict_nw["Loads"]
        loads_df.set_index("name", inplace=True)  # Set "name" column as the index
        network.import_components_from_dataframe(loads_df, "Load")
    
    # Import "Stores" components
    if "Stores" in dataframes_dict_nw:
        stores_df = dataframes_dict_nw["Stores"]
        
        # Löschen der Spalten "name" und "bus"
        stores_df.drop(columns=["name", "bus"], inplace=True)
        
        # Umbenennen der Spalten "name.1" in "name" und "bus.1" in "bus"
        stores_df.rename(columns={"name.1": "name", "bus.1": "bus"}, inplace=True)
        
        # Setzen der Spalte "name" als Index
        stores_df.set_index("name", inplace=True)
        
        network.import_components_from_dataframe(stores_df, "Store")
        
    # Import "Generators-p_max_pu" components
    
    if "Generators-p_max_pu" in dataframes_dict_nw:
        p_max_pu_df = dataframes_dict_nw["Generators-p_max_pu"]
        p_max_pu_df.set_index("snapshots", inplace=True)
        network.set_snapshots(p_max_pu_df.index)
        network.import_series_from_dataframe(p_max_pu_df, "Generator", "p_max_pu")        
    return network   

def optimize_and_export_to_csv(network, export_folder):
    # Optimize the network
    network.optimize()

    # Export the optimized network to CSV files
    network.export_to_csv_folder(export_folder)

    # Return the optimized network
    return network
 
def export_components_to_excel(network, postpro_folder):
    component_list = ["buses", "generators", "loads", "links", "stores"]

    existing_files = [f for f in os.listdir(postpro_folder) if f.startswith("Results_export")]
    existing_indices = [int(f.split("_")[2].split(".")[0]) for f in existing_files if len(f.split("_")) == 3]
    if existing_indices:
        next_index = max(existing_indices) + 1
    else:
        next_index = 1
    
    output_filename = f"Results_export_{next_index}.xlsx"
    output_path = os.path.join(postpro_folder, output_filename)
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for component_name in component_list:
            component_df = getattr(network, component_name)
            component_df.to_excel(writer, sheet_name=component_name, index=False)
    
    # Export the time-dependent component DataFrames
        for component_type in ["generators_t"]:
            for attribute in ["p", "p_set", "marginal_cost", "efficiency", "stand_by_cost", "status"]:
                component_df_t = getattr(network, component_type)
                attribute_df = component_df_t[attribute]
                sheet_name = f"{component_type}.{attribute}"
                attribute_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        for component_type in ["buses_t"]:
            for attribute in ["v_mag_pu_set", "p", "marginal_price"]:
                component_df_t = getattr(network, component_type)
                attribute_df = component_df_t[attribute]
                sheet_name = f"{component_type}.{attribute}"
                attribute_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        for component_type in ["stores_t"]:
            for attribute in ["e_min_pu", "e_max_pu", "p_set", "marginal_cost", "standing_loss", "p", "e"]:
                component_df_t = getattr(network, component_type)
                attribute_df = component_df_t[attribute]
                sheet_name = f"{component_type}.{attribute}"
                attribute_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        for component_type in ["loads_t"]:
            for attribute in ["p_set","p"]:
                component_df_t = getattr(network, component_type)
                attribute_df = component_df_t[attribute]
                sheet_name = f"{component_type}.{attribute}"
                attribute_df.to_excel(writer, sheet_name=sheet_name, index=False)
                         
    print(f"Results_export_{next_index}.xlsx has been saved!")

def CST_cost_calc(dataframes_dict, p_opt):
    solarfield_power = p_opt
    CST_df = dataframes_dict['CST']
    
    solarfield_ref_cost = CST_df.loc[CST_df['name'] == 'solarfield_ref_cost', 'value'].values[0]
    solarfield_ref_capacity = CST_df.loc[CST_df['name'] == 'solarfield_reference_capacity', 'value'].values[0]
    solarfield_eos_factor = CST_df.loc[CST_df['name'] == 'solarfield_eos_factor', 'value'].values[0]

    solarfield_specific_cost = solarfield_ref_cost * (solarfield_power / solarfield_ref_capacity) ** solarfield_eos_factor
    
    print("Calculated specific cost", solarfield_specific_cost)
    return solarfield_specific_cost

def iterative_optimization(network):
    
    old_p_opt = network.generators.p_nom["CST"]
    
    for _ in range(25):
        capital_cost_CST = CST_cost_calc(dataframes_dict, old_p_opt)
        network.generators.capital_cost["CST"] = capital_cost_CST
        optimize_and_export_to_csv(network, export_folder)
        export_components_to_excel(network, postpro_folder)
    return network
        
    

if __name__ == "__main__":
    xlsx_file = r"D:\OneDrive - Fichtner GmbH & Co. KG\Masterarbeit\Repository\Master-Thesis\OPTIMIZER\data\Setup.xlsx" # Füge hier den Pfad ein, der die CSV Setup Daten enthält
    output_folder = r"D:\OneDrive - Fichtner GmbH & Co. KG\Masterarbeit\Repository\Master-Thesis\OPTIMIZER\data\CSV_data" # FÜge hier den Pfad ein, wo die CSV Dateien gespeichert werden
    export_folder = r"D:\OneDrive - Fichtner GmbH & Co. KG\Masterarbeit\Repository\Master-Thesis\OPTIMIZER\data\Output"
    postpro_folder = r"D:\OneDrive - Fichtner GmbH & Co. KG\Masterarbeit\Repository\Master-Thesis\OPTIMIZER\data\Post_processing"
    selected_sheets = ["buses", "carriers", "generators","stores", "links","loads", "generators-p_max_pu"]  # Füge hier die Namen der gewünschten Blätter hinzu
    selected_sheets_techno = ["CST", "TES"]
    csv_folder = r"D:\OneDrive - Fichtner GmbH & Co. KG\Masterarbeit\Repository\Master-Thesis\OPTIMIZER\data\CSV_data"  # Passe den Pfad zum Ordner mit den CSV-Dateien an
    csv_file_path_p_max_pu = r"D:\OneDrive - Fichtner GmbH & Co. KG\Masterarbeit\Repository\Master-Thesis\OPTIMIZER\data\CSV_data\generators-p_max_pu.csv"
    
    # Erstelle den Ausgabeordner, falls er nicht existiert
    os.makedirs(output_folder, exist_ok=True)
    
    # Erstelle CSV Dateien aus Setup.xlsx
    xlsx_to_csv(xlsx_file, output_folder, selected_sheets)
    
    # Erstelle DataFrames für die nicht integrierten Inputs
    dataframes_dict = Inputs_dataframes(xlsx_file, selected_sheets_techno)
    
    # Erstelle Dataframes für die Netzwerk Inputs
    dataframes_dict_nw = create_dataframes_from_csv(csv_folder, selected_sheets)
        
    # Erstelle Netzwerk
    network = pypsa.Network()   

    # Erstelle Komponenten im netzwerk aus Dataframes
    network = import_components_from_dataframes(dataframes_dict_nw, network) 
     
    optimize_and_export_to_csv(network, export_folder)
    export_components_to_excel(network, postpro_folder)
    iterative_optimization(network)
