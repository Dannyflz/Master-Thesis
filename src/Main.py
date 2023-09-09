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
    if "Constraints" in dataframes_dict_nw:
        constraints_df = dataframes_dict_nw["Constraints"]
        constraints_df.set_index("name", inplace=True)  # Set "name" column as the index
        network.import_components_from_dataframe(constraints_df, "GlobalConstraint")
    
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
    
    if "Generators-p_max_pu" in dataframes_dict_nw:
        p_max_pu_df = dataframes_dict_nw["Generators-p_max_pu"]
        p_max_pu_df.set_index("snapshots", inplace=True)
        network.set_snapshots(p_max_pu_df.index)
        network.import_series_from_dataframe(p_max_pu_df, "Generator", "p_max_pu")        
    return network   

def optimize(network):
    # Optimize the network
    network.optimize()
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

def calculate_technology_sums(network):
    technology_sums = {}
    generators_t_df = getattr(network, "generators_t")
    generators_t_p_df = generators_t_df["p"]
    
    for technology in generators_t_p_df.columns:
        if technology != "name":
            column_sum = generators_t_p_df[technology].sum()
            technology_sums[technology] = column_sum
    print (technology_sums)        
    return technology_sums

def calculate_co2_emissions(network):
    
    generators_t_df = getattr(network, "generators_t")
    generators_t_p_df = generators_t_df["p"]
    backup_sum = generators_t_p_df["Backup"].sum()
    
    carrier_df = getattr(network, "carriers")
    carrier_emissions = carrier_df["co2_emissions"]
    print(carrier_emissions)
    
    co2_backup = backup_sum / network.generators.efficiency["Backup"] * carrier_emissions["fossil heat"]
    return co2_backup       

def calc_initial_costs(network, dataframes_dict):
    # Hier werden Daten aus der Setup Datei eingelesen.
    
    
    Input_df = dataframes_dict['Inputs']
    initial_cst_solar_field = Input_df.loc[Input_df["name"] == "initial_cst_solar_field", "value"].values[0]
    initial_cst_steam_generator = Input_df.loc[Input_df["name"] == "initial_cst_steam_generator", "value"].values[0]
    initial_tes_charge_power = Input_df.loc[Input_df["name"] == "initial_tes_charge_power", "value"].values[0]
    initial_tes_capacity = Input_df.loc[Input_df["name"] == "initial_tes_capacity", "value"].values[0]
    initial_ppa_power = Input_df.loc[Input_df["name"] == "initial_ppa_power", "value"].values[0]
    initial_wind_power = Input_df.loc[Input_df["name"] == "initial_wind_power", "value"].values[0]
    initial_pv_power = Input_df.loc[Input_df["name"] == "initial_pv_power", "value"].values[0]
    initial_etes_heater_power = Input_df.loc[Input_df["name"] == "initial_etes_heater_power", "value"].values[0]
    initial_etes_capacity = Input_df.loc[Input_df["name"] == "initial_etes_capacity", "value"].values[0]
    initial_backup_power = Input_df.loc[Input_df["name"] == "initial_backup_power", "value"].values[0]
    
    cst_design_dni = Input_df.loc[Input_df["name"] == "cst_design_dni", "value"].values[0]
    cst_total_efficiency = Input_df.loc[Input_df["name"] == "cst_total_efficiency", "value"].values[0]
    cst_sf_capacity = Input_df.loc[Input_df["name"] == "cst_sf_capacity", "value"].values[0]
    cst_heat_output = Input_df.loc[Input_df["name"] == "cst_heat_output", "value"].values[0]
    cst_annual_heat = Input_df.loc[Input_df["name"] == "cst_annual_heat", "value"].values[0]
    cst_sf_cost_m2 = Input_df.loc[Input_df["name"] == "cst_sf_cost_m2", "value"].values[0]
    cst_site_prep = Input_df.loc[Input_df["name"] == "cst_site_prep", "value"].values[0]
    cst_HTF_cost = Input_df.loc[Input_df["name"] == "cst_HTF_cost", "value"].values[0]
    cst_steam_gen_cost = Input_df.loc[Input_df["name"] == "cst_steam_gen_cost", "value"].values[0]
    cst_EPC_contingencies = Input_df.loc[Input_df["name"] == "cst_EPC_contingencies", "value"].values[0]
    cst_owner_cost = Input_df.loc[Input_df["name"] == "cst_owner_cost", "value"].values[0]
    cst_OM_variable = Input_df.loc[Input_df["name"] == "cst_OM_variable", "value"].values[0]
    cst_OM_fixed = Input_df.loc[Input_df["name"] == "cst_OM_fixed", "value"].values[0]
    cst_aperture_area = Input_df.loc[Input_df["name"] == "cst_aperture_area", "value"].values[0]
    cst_solar_multiple = Input_df.loc[Input_df["name"] == "cst_solar_multiple", "value"].values[0]
    cst_OM_fixed = Input_df.loc[Input_df["name"] == "cst_OM_fixed", "value"].values[0]
    cst_OM_variable = Input_df.loc[Input_df["name"] == "cst_OM_variable", "value"].values[0]
    tes_capacity = Input_df.loc[Input_df["name"] == "tes_capacity", "value"].values[0]
    tes_capacity_cost = Input_df.loc[Input_df["name"] == "tes_capacity_cost", "value"].values[0]
    tes_power_in = Input_df.loc[Input_df["name"] == "tes_power_in", "value"].values[0]
    tes_power_out = Input_df.loc[Input_df["name"] == "tes_power_out", "value"].values[0]    
    tes_power_cost = Input_df.loc[Input_df["name"] == "tes_power_cost", "value"].values[0]
    tes_efficiency = Input_df.loc[Input_df["name"] == "tes_efficiency", "value"].values[0]
    tes_OM_fixed = Input_df.loc[Input_df["name"] == "tes_OM_fixed", "value"].values[0]
    tes_EPC = Input_df.loc[Input_df["name"] == "tes_EPC", "value"].values[0]
    tes_owners_cost = Input_df.loc[Input_df["name"] == "tes_owners_cost", "value"].values[0]
    pv_peak_capacity = Input_df.loc[Input_df["name"] == "pv_peak_capacity", "value"].values[0]
    pv_turnkey_cost_Input = Input_df.loc[Input_df["name"] == "pv_turnkey_cost_Input", "value"].values[0]
    pv_EPC_contingencies = Input_df.loc[Input_df["name"] == "pv_EPC_contingencies", "value"].values[0]
    pv_owner_cost = Input_df.loc[Input_df["name"] == "pv_owner_cost", "value"].values[0]
    pv_OM_variable = Input_df.loc[Input_df["name"] == "pv_OM_variable", "value"].values[0]
    pv_OM_fixed = Input_df.loc[Input_df["name"] == "pv_OM_fixed", "value"].values[0]
    wind_peak_capacity = Input_df.loc[Input_df["name"] == "wind_peak_capacity", "value"].values[0]
    wind_turnkey_cost = Input_df.loc[Input_df["name"] == "wind_turnkey_cost", "value"].values[0]
    wind_epc_cost = Input_df.loc[Input_df["name"] == "wind_epc_cost", "value"].values[0]
    wind_owner_cost = Input_df.loc[Input_df["name"] == "wind_owner_cost", "value"].values[0]
    wind_OM_variable = Input_df.loc[Input_df["name"] == "wind_OM_variable", "value"].values[0]
    wind_OM_fixed = Input_df.loc[Input_df["name"] == "wind_OM_fixed", "value"].values[0]
    etes_capacity = Input_df.loc[Input_df["name"] == "etes_capacity", "value"].values[0]
    etes_heater_power = Input_df.loc[Input_df["name"] == "etes_heater_power", "value"].values[0]
    etes_turnkey_cost = Input_df.loc[Input_df["name"] == "etes_turnkey_cost", "value"].values[0]
    etes_heater_cost = Input_df.loc[Input_df["name"] == "etes_cost_heater", "value"].values[0]
    etes_epc_cost = Input_df.loc[Input_df["name"] == "etes_epc_cost", "value"].values[0]
    etes_owner_cost = Input_df.loc[Input_df["name"] == "etes_owner_cost", "value"].values[0]
    etes_OM_variable = Input_df.loc[Input_df["name"] == "etes_OM_variable", "value"].values[0]
    etes_OM_fixed = Input_df.loc[Input_df["name"] == "etes_OM_fixed", "value"].values[0]
    etes_OM_fixed = Input_df.loc[Input_df["name"] == "etes_OM_fixed", "value"].values[0]
    backup_power_output = Input_df.loc[Input_df["name"] == "backup_power_output", "value"].values[0]
    backup_main_machinery_cost = Input_df.loc[Input_df["name"] == "backup_main_machinery_cost", "value"].values[0]
    backup_electric_IC_system_cost = Input_df.loc[Input_df["name"] == "backup_electric_IC_system_cost", "value"].values[0]
    backup_transport_installation_commissioning_cost = Input_df.loc[Input_df["name"] == "backup_transport_installation_commissioning_cost", "value"].values[0]
    backup_epc_cost = Input_df.loc[Input_df["name"] == "backup_epc_cost", "value"].values[0]
    backup_owners_cost = Input_df.loc[Input_df["name"] == "backup_owners_cost", "value"].values[0]
    backup_OM_fixed = Input_df.loc[Input_df["name"] == "backup_OM_fixed", "value"].values[0]
    backup_OM_variable = Input_df.loc[Input_df["name"] == "backup_OM_variable", "value"].values[0]
    backup_annual_degradation = Input_df.loc[Input_df["name"] == "backup_annual_degradation", "value"].values[0]
    backup_fuel_cost = Input_df.loc[Input_df["name"] == "backup_fuel_cost", "value"].values[0]
    backup_total_capex = Input_df.loc[Input_df["name"] == "backup_total_capex", "value"].values[0]
    backup_OM_fixed = Input_df.loc[Input_df["name"] == "backup_OM_fixed", "value"].values[0]
    backup_turnkey_cost = Input_df.loc[Input_df["name"] == "backup_turnkey_cost", "value"].values[0]
    
    emission_cost = Input_df.loc[Input_df["name"] == "emission_cost", "value"].values[0]
    emission_fossil_heat = Input_df.loc[Input_df["name"] == "emission_fossil_heat", "value"].values[0]
    eos_cst_sf_cost = Input_df.loc[Input_df['name'] == 'cst_sf_cost_m2', 'EoS'].values[0]
    eos_cst_site_prep_m2 = Input_df.loc[Input_df['name'] == 'cst_site_prep', 'EoS'].values[0]
    eos_cst_HTF_cost = Input_df.loc[Input_df['name'] == 'cst_HTF_cost', 'EoS'].values[0]
    eos_cst_steam_gen = Input_df.loc[Input_df['name'] == 'cst_steam_gen_cost', 'EoS'].values[0]
    eos_tes_capacity = Input_df.loc[Input_df['name'] == 'tes_capacity_cost', 'EoS'].values[0]
    eos_tes_power = Input_df.loc[Input_df['name'] == 'tes_power_cost', 'EoS'].values[0]
    eos_pv =  Input_df.loc[Input_df['name'] == 'pv_turnkey_cost_Input', 'EoS'].values[0]
    eos_wind =  Input_df.loc[Input_df['name'] == 'wind_turnkey_cost', 'EoS'].values[0]
    eos_etes = Input_df.loc[Input_df['name'] == 'etes_turnkey_cost', 'EoS'].values[0]
    eos_etes_power = Input_df.loc[Input_df['name'] == 'etes_power_cost', 'EoS'].values[0]
    eos_backup_mm = Input_df.loc[Input_df['name'] == 'backup_main_machinery_cost', 'EoS'].values[0]
    
    # Hier werden die neuen spezifischen Kosten berechnet anhand der Komponenten Eingabe und den EoS Faktoren.  
    cst_aperture_area = cst_sf_capacity * 1000000 / cst_design_dni / cst_total_efficiency
    new_aperture_area = initial_cst_solar_field * 1000000 / cst_design_dni / cst_total_efficiency
    new_site_prep = cst_site_prep * cst_aperture_area / new_aperture_area * (new_aperture_area / cst_aperture_area) ** eos_cst_site_prep_m2
    new_sf_cost_m2 = cst_sf_cost_m2 * cst_aperture_area / new_aperture_area * (new_aperture_area / cst_aperture_area) ** eos_cst_sf_cost
    new_HTF_cost = cst_HTF_cost * cst_sf_capacity / initial_cst_solar_field * (initial_cst_solar_field / cst_sf_capacity) ** eos_cst_HTF_cost
    new_steam_gen_cost = cst_steam_gen_cost * cst_heat_output / initial_cst_steam_generator * (initial_cst_steam_generator / cst_heat_output) ** eos_cst_steam_gen
    new_tes_capacity_cost = tes_capacity_cost * tes_capacity / initial_tes_capacity * (initial_tes_capacity / tes_capacity) ** eos_tes_capacity
    new_tes_power_cost = tes_power_cost * tes_power_in / initial_tes_charge_power * (initial_tes_charge_power / tes_power_in) ** eos_tes_power
    new_pv_turnkey_cost = pv_turnkey_cost_Input * pv_peak_capacity / initial_pv_power * (initial_pv_power / pv_peak_capacity) ** eos_pv
    new_wind_turnkey_cost = wind_turnkey_cost * wind_peak_capacity / initial_wind_power * (initial_wind_power / wind_peak_capacity) ** eos_wind
    new_etes_capacity_cost = etes_turnkey_cost * etes_capacity / initial_etes_capacity * (initial_etes_capacity / etes_capacity) ** eos_etes
    new_etes_power_cost = etes_heater_cost * etes_heater_power / initial_etes_heater_power * (initial_etes_heater_power / etes_heater_power) ** eos_etes_power
    new_backup_MM_cost = backup_main_machinery_cost * backup_power_output / initial_backup_power * (initial_backup_power / backup_power_output) ** eos_backup_mm
    
    # Hier werden alle Kosten der CST (Ohne Steam Generation) Anlage berechnet
    cst_site_prep_capex = new_site_prep * new_aperture_area
    cst_sf_cost_capex = new_aperture_area * new_sf_cost_m2
    cst_HTF_cost_capex = new_HTF_cost * initial_cst_solar_field
    cst_EPC_capex = (cst_site_prep_capex + cst_sf_cost_capex + cst_HTF_cost_capex) * cst_EPC_contingencies #heat gen und backup heater werden in ihrer Größe auch optimiert und PyPSA rechnet ihre Kosten selbstständig hinzu
    cst_owner_capex = (cst_site_prep_capex + cst_sf_cost_capex + cst_HTF_cost_capex + cst_EPC_capex) * cst_owner_cost
    cst_total_capex = cst_site_prep_capex + cst_sf_cost_capex + cst_HTF_cost_capex + cst_EPC_capex + cst_owner_capex
    cst_OM_fixed_capex = cst_total_capex * cst_OM_fixed
    cst_turnkey_capex = (cst_total_capex + cst_OM_fixed_capex) / initial_cst_solar_field
    print("CST Turnkey Capex are =" + cst_turnkey_capex + "$/MW")
    print("CST Turnkey Opex are =" + cst_OM_variable + "$/MWh")
    
    # Hier wird nur der Steam Generation Anteil berechnet. 
    # Da dieses Bauteil der übertragbaren Wärmeleistung entspricht, wird es seperat optimiert und modeliert.
    cst_steam_gen_capex = initial_cst_steam_generator * new_steam_gen_cost
    cst_steam_gen_EPC_capex = cst_steam_gen_capex * cst_EPC_contingencies
    cst_steam_gen_owners_capex = cst_steam_gen_EPC_capex * cst_owner_cost
    cst_steam_gen_total_capex = cst_steam_gen_capex + cst_steam_gen_EPC_capex + cst_steam_gen_owners_capex
    cst_steam_gen_OM_fixed_capex = cst_steam_gen_total_capex * cst_OM_fixed
    cst_steam_gen_turnkey_capex = (cst_steam_gen_total_capex + cst_steam_gen_OM_fixed_capex) / initial_cst_steam_generator
    print("CST Steamgen Turnkey Capex are =" + cst_steam_gen_turnkey_capex + "$/MW")
        
    # Hier werden die Kosten des TES berechnet 
    # Es gibt die Möglichkeit die Kosten nach Kapazität und Power aufzuteilen.
    
    tes_capacity_capex = initial_tes_capacity * new_tes_capacity_cost
    tes_capacity_epc_capex = (tes_capacity_capex ) * tes_EPC
    tes_capacity_owners_cost_capex = (tes_capacity_capex + tes_capacity_epc_capex) * tes_owners_cost
    tes_capacity_total_capex = (tes_capacity_capex +  tes_capacity_epc_capex + tes_capacity_owners_cost_capex)
    tes_capacity_OM_fixed_capex = tes_capacity_total_capex * tes_OM_fixed
    tes_capacity_turnkey = (tes_capacity_OM_fixed_capex + tes_capacity_total_capex) / initial_tes_capacity
    print("TES Capacity Turnkey Capex are =" + tes_capacity_turnkey + "$/MWh")
    
    tes_power_capex = initial_tes_charge_power * new_tes_power_cost
    tes_power_epc_capex = (tes_power_capex) * tes_EPC
    tes_power_owners_cost_capex = (tes_power_capex + tes_power_epc_capex) * tes_EPC
    tes_power_total_capex = tes_power_capex + tes_power_epc_capex + tes_power_owners_cost_capex
    tes_power_OM_fixed_capex = tes_power_total_capex * tes_OM_fixed
    tes_power_turnkey = (tes_power_OM_fixed_capex + tes_power_total_capex) / initial_tes_charge_power
    print("TES Power Turnkey Capex are =" + tes_power_turnkey + "$/MW")
    
    
    # Hier werden die Turnkey Cost für PV berechnet
    pv_capacity_capex = initial_pv_power * new_pv_turnkey_cost
    pv_epc_capex = pv_capacity_capex * pv_EPC_contingencies
    pv_owners_cost_capex = (pv_capacity_capex + pv_epc_capex) * pv_owner_cost
    pv_total_capex = pv_capacity_capex + pv_epc_capex + pv_owners_cost_capex
    pv_OM_fixed_capex = pv_total_capex * pv_OM_fixed
    pv_turnkey_capex = (pv_OM_fixed_capex + pv_total_capex) / initial_pv_power
    print("PV Turnkey Capex are =" + pv_turnkey_capex + "$/MW")
    
    # Hier werden die Turnkey Cost für WIND berechnet
    wind_capacity_capex = initial_wind_power * new_wind_turnkey_cost
    wind_epc_capex = wind_capacity_capex * wind_epc_cost
    wind_owners_cost_capex = (wind_capacity_capex + wind_epc_capex) * wind_owner_cost
    wind_total_capex = wind_capacity_capex + wind_epc_capex + wind_owners_cost_capex
    wind_OM_fixed_capex = wind_total_capex * wind_OM_fixed
    wind_turnkey_capex = (wind_OM_fixed_capex + wind_total_capex) / initial_wind_power
    print("Wind Turnkey Capex are =" + wind_turnkey_capex + "$/MW")
    
    # Hier werden die Kosten des etes berechnet 
    # Es gibt die Möglichkeit die Kosten nach Kapazität und Power aufzuteilen.
    etes_capacity_capex = initial_etes_capacity * new_etes_capacity_cost
    etes_capacity_epc_capex = (etes_capacity_capex ) * etes_epc_cost
    etes_capacity_owners_cost_capex = (etes_capacity_capex + etes_capacity_epc_capex) * etes_owner_cost
    etes_capacity_total_capex = (etes_capacity_capex +  etes_capacity_epc_capex + etes_capacity_owners_cost_capex)
    etes_capacity_OM_fixed_capex = etes_capacity_total_capex * etes_OM_fixed
    etes_capacity_turnkey = (etes_capacity_OM_fixed_capex + etes_capacity_total_capex) / initial_etes_capacity
    print("etes Capacity Turnkey Capex are =" + etes_capacity_turnkey + "$/MWh")
    
    etes_power_capex = initial_etes_heater_power * new_etes_power_cost
    etes_power_epc_capex = (etes_power_capex) * etes_epc_cost
    etes_power_owners_cost_capex = (etes_power_capex + etes_power_epc_capex) * etes_owner_cost
    etes_power_total_capex = etes_power_capex + etes_power_epc_capex + etes_power_owners_cost_capex
    etes_power_OM_fixed_capex = etes_power_total_capex * etes_OM_fixed
    etes_power_turnkey = (etes_power_OM_fixed_capex + etes_power_total_capex) / initial_etes_heater_power
    print("etes Power Turnkey Capex are =" + etes_power_turnkey + "$/MW")
    
    
    
    backup_MM_capex = initial_backup_power * new_backup_MM_cost
    backup_electric_system_capex = backup_MM_capex * backup_electric_IC_system_cost
    backup_t_i_c_capex = backup_MM_capex * backup_transport_installation_commissioning_cost
    backup_epc_capex = (backup_MM_capex + backup_electric_system_capex + backup_t_i_c_capex) * backup_epc_cost
    backup_owners_cost_capex = (backup_MM_capex + backup_electric_system_capex + backup_t_i_c_capex + backup_epc_capex) * backup_owners_cost
    backup_total_capex = (backup_MM_capex + backup_electric_system_capex + backup_t_i_c_capex + backup_epc_capex + backup_owners_cost_capex)
    backup_OM_fixed_capex = backup_total_capex * backup_OM_fixed
    backup_turnkey = (backup_total_capex + backup_OM_fixed_capex) / initial_backup_power
    print("Backup Power Turnkey Capex are =" + backup_turnkey + "$/MW")
    
    # OPEX Calculation
    backup_marginal_cost = backup_fuel_cost + emission_cost * emission_fossil_heat
    
    
    
    network.generators.capital_cost["CST"] = calc_turnkey_cost
    network.generators.marginal_cost["CST"] = calc_turnkey_opex
    network.links.capital_cost["Heat Sink CST"] = calc_steam_gen_cost
    network.stores.capital_cost["TES"] = calc_tes_turnkey

    print("The initial cost for CST Capital Costs are", network.generators.capital_cost["CST"], "\n"
          "The initial cost for CST Marginal Costs are", network.generators.marginal_cost["CST"], "\n"
          "The initial cost for Heat sink are", network.links.capital_cost["Heat Sink CST"], "\n"
          "The initial cost for TES are", network.stores.capital_cost["TES"])   
    return network

def calc_iterative_cost(dataframes_dict, network, technology_sums):
    # Hier werden Daten aus der Setup Datei eingelesen.
    new_sf_capacity = network.generators.p_nom_opt["CST"]
    new_heat_output = network.links.p_nom_opt["Heat Sink CST"]
    new_backup_heater_capacity = network.generators.p_nom_opt["Backup"]
    new_tes_capacity = network.stores.e_nom_opt["TES"]
    new_annual_generation = technology_sums["CST"] 
    
    CST_df = dataframes_dict['CST']
    TES_df = dataframes_dict["TES"]
    
    # Alle Größen der CST Modellierung
    ref_design_dni = CST_df.loc[CST_df['name'] == 'ref_design_dni', 'value'].values[0]
    ref_total_efficiency = CST_df.loc[CST_df['name'] == 'ref_total_efficiency', 'value'].values[0]
    ref_sf_capacity = CST_df.loc[CST_df['name'] == 'ref_sf_capacity', 'value'].values[0]
    ref_heat_output = CST_df.loc[CST_df['name'] == 'ref_heat_output', 'value'].values[0]
    ref_backup_heater_capacity = CST_df.loc[CST_df['name'] == 'ref_backup_heater_capacity', 'value'].values[0]
    ref_annual_heat = CST_df.loc[CST_df['name'] == 'ref_annual_heat', 'value'].values[0]
    ref_sf_cost_m2 = CST_df.loc[CST_df['name'] == 'ref_sf_cost_m2', 'value'].values[0]
    ref_site_prep = CST_df.loc[CST_df['name'] == 'ref_site_prep', 'value'].values[0]
    ref_HTF_cost = CST_df.loc[CST_df['name'] == 'ref_HTF_cost', 'value'].values[0]
    ref_steam_gen_cost = CST_df.loc[CST_df['name'] == 'ref_steam_gen_cost', 'value'].values[0]
    ref_backup_heater_cost = CST_df.loc[CST_df['name'] == 'ref_backup_heater_cost', 'value'].values[0]
    ref_EPC_contingencies = CST_df.loc[CST_df['name'] == 'ref_EPC_contingencies', 'value'].values[0]
    ref_owner_cost = CST_df.loc[CST_df['name'] == 'ref_owner_cost', 'value'].values[0]
    ref_OM_variable = CST_df.loc[CST_df['name'] == 'ref_OM_variable', 'value'].values[0]
    ref_OM_fixed = CST_df.loc[CST_df['name'] == 'ref_OM_fixed', 'value'].values[0]
    ref_annual_degradation = CST_df.loc[CST_df['name'] == 'ref_annual_degradation', 'value'].values[0]
    ref_backup_fuel = CST_df.loc[CST_df['name'] == 'ref_backup_fuel', 'value'].values[0]
    ref_governmental_subsidy = CST_df.loc[CST_df['name'] == 'ref_governmental_subsidy', 'value'].values[0]
    ref_aperture_area = CST_df.loc[CST_df['name'] == 'ref_aperture_area', 'value'].values[0]
    ref_solar_multiple = CST_df.loc[CST_df['name'] == 'ref_solar_multiple', 'value'].values[0]
    
    ref_tes_capacity = TES_df.loc[TES_df['name'] == 'ref_tes_capacity', 'value'].values[0]
    ref_tes_capital_cost = TES_df.loc[TES_df['name'] == 'ref_tes_capital_cost', 'value'].values[0]
    ref_tes_T_in = TES_df.loc[TES_df['name'] == 'ref_tes_T_in', 'value'].values[0]
    ref_tes_T_out = TES_df.loc[TES_df['name'] == 'ref_tes_T_out', 'value'].values[0]
    ref_tes_T_heat_consumer = TES_df.loc[TES_df['name'] == 'ref_tes_T_heat_consumer', 'value'].values[0]
    ref_tes_EPC = TES_df.loc[TES_df['name'] == 'ref_tes_EPC', 'value'].values[0]
    ref_tes_owners_cost = TES_df.loc[TES_df['name'] == 'ref_tes_owners_cost', 'value'].values[0]
    
    eos_site_prep_m2 = CST_df.loc[CST_df['name'] == 'ref_site_prep', 'EoS'].values[0]
    eos_HTF_cost = CST_df.loc[CST_df['name'] == 'ref_HTF_cost', 'EoS'].values[0]
    eos_steam_gen = CST_df.loc[CST_df['name'] == 'ref_steam_gen_cost', 'EoS'].values[0]
    eos_backup_heater = CST_df.loc[CST_df['name'] == 'ref_backup_heater_cost', 'EoS'].values[0]

    eos_tes = TES_df.loc[TES_df['name'] == 'ref_tes_capital_cost', 'EoS'].values[0]
    
    # Hier werden die Größen neu berechnet anhand der Komponenten Eingabe und den EoS Faktoren.  
    calc_aperture_area = new_sf_capacity * 1000000 / ref_design_dni / ref_total_efficiency
    calc_solar_multiple = new_sf_capacity / new_heat_output
    new_site_prep = ref_site_prep * ref_aperture_area / calc_aperture_area * (calc_aperture_area / ref_aperture_area) ** eos_site_prep_m2
    new_sf_cost_m2 = ref_sf_cost_m2 * ref_aperture_area / calc_aperture_area * (calc_aperture_area / ref_aperture_area) ** eos_site_prep_m2
    new_HTF_cost = 1000 * ref_HTF_cost * ref_sf_capacity / new_sf_capacity * (new_sf_capacity / ref_sf_capacity) ** eos_HTF_cost
    new_steam_gen_cost = 1000 * ref_steam_gen_cost * ref_heat_output / new_heat_output * (new_heat_output / ref_heat_output) ** eos_steam_gen
    new_tes_cost = 1000 * ref_tes_capital_cost * ref_tes_capacity / new_tes_capacity * (new_tes_capacity / ref_tes_capacity) ** eos_tes
    
    # Hier werden die Kapital Kosten der CST Anlage berechnet
    calc_site_prep_cost = new_site_prep * calc_aperture_area
    calc_sf_cost = calc_aperture_area * new_sf_cost_m2
    calc_HTF_cost = new_HTF_cost * new_sf_capacity
    calc_steam_gen_cost = new_heat_output * new_steam_gen_cost
    calc_EPC_cost = (calc_site_prep_cost + calc_sf_cost + calc_HTF_cost) * ref_EPC_contingencies #heat gen und backup heater werden in ihrer Größe auch optimiert und PyPSA rechnet ihre Kosten selbstständig hinzu
    calc_owner_cost = calc_EPC_cost * ref_owner_cost
    calc_total_capex = calc_owner_cost + calc_HTF_cost + calc_EPC_cost + calc_steam_gen_cost + calc_sf_cost + calc_site_prep_cost
    calc_turnkey_cost = calc_total_capex / new_sf_capacity
    
    # Hier werden die Betriebskosten der CST Anlage berechnet
    calc_OM_variable = ref_OM_variable
    calc_OM_fixed = ref_OM_fixed * calc_total_capex / new_annual_generation / 1000
    calc_annual_degradation = new_annual_generation * ref_annual_degradation / 1000
    calc_turnkey_opex = calc_OM_fixed + calc_OM_variable + calc_annual_degradation
    
    # Hier werden die Kapital Kosten des Heat Outputs berechnet
    calc_steam_gen_cost = new_steam_gen_cost * new_heat_output
    
    # Hier werden die Kapital Kosten des TES berechnet
    calc_capex_tes = new_tes_capacity * new_tes_cost
    calc_tes_EPC = calc_capex_tes * ref_tes_EPC
    calc_tes_owners_cost = (calc_capex_tes + calc_tes_EPC) * ref_tes_owners_cost
    calc_total_capex_tes = calc_capex_tes + calc_tes_EPC + calc_tes_owners_cost
    calc_tes_turnkey = calc_total_capex_tes / new_tes_capacity
    
    network.generators.capital_cost["CST"] = calc_turnkey_cost
    network.generators.marginal_cost["CST"] = calc_turnkey_opex
    network.links.capital_cost["Heat Sink CST"] = calc_steam_gen_cost
    network.stores.capital_cost["TES"] = calc_tes_turnkey

    print("The initial cost for CST Capital Costs are", network.generators.capital_cost["CST"], "\n"
          "The initial cost for CST Marginal Costs are", network.generators.marginal_cost["CST"], "\n"
          "The initial cost for Heat sink are", network.links.capital_cost["Heat Sink CST"], "\n"
          "The initial cost for TES are", network.stores.capital_cost["TES"])

    return network

def iterative_optimization(network, postpro_folder):
    # Initialize the constant for the emission limit
    initial_constant = 1
    
    # Initialize an empty dictionary to store collected data
    collected_data = {}
    co2_backup = calculate_co2_emissions(network)
    network.add("GlobalConstraint",
                "emission_limit",
                carrier_attribute = "co2_emissions",
                sense = "<=",
                constant = co2_backup * initial_constant,    
                )
    
    # Begin the iterative process
    for iteration in range(10):
        print(f"Iteration {iteration + 1}")
        
        # Optimize the network
        technology_sums = calculate_technology_sums(network)
        initial_constant -= 0.1
        network.global_constraints.constant["emission_limit"] = co2_backup * initial_constant
        network = calc_iterative_cost(dataframes_dict, network, technology_sums)
        network = optimize(network)
        export_components_to_excel(network, postpro_folder)

        # Collect and store relevant data
        iteration_data = {
            "CST_p_nom": network.generators.p_nom["CST"],
            "CST_p_nom_opt": network.generators.p_nom_opt["CST"],
            "CST_Capital_Cost": network.generators.capital_cost["CST"],
            "CST_Marginal_Cost": network.generators.marginal_cost["CST"],
            "PV Capital Cost": network.generators.capital_cost["PV"],
            "PV Marginal Cost": network.generators.marginal_cost["PV"],
            "PV p_nom": network.generators.p_nom["PV"],
            "PV p_nom_opt": network.generators.p_nom_opt["PV"],
            "WIND Capital Cost": network.generators.capital_cost["Wind"],
            "WIND Marginal Cost": network.generators.marginal_cost["Wind"],
            "WIND p_nom": network.generators.p_nom["Wind"],
            "WIND p_nom_opt": network.generators.p_nom_opt["Wind"],
            "PPA Capital Cost": network.generators.capital_cost["PPA"],
            "PPA Marginal Cost": network.generators.marginal_cost["PPA"],
            "PPA p_nom": network.generators.p_nom["PPA"],
            "PPA p_nom_opt": network.generators.p_nom_opt["PPA"],
            "BackUp Boiler Capital Cost": network.generators.capital_cost["Backup"],
            "BackUp Boiler Marginal Cost": network.generators.marginal_cost["Backup"],
            "BackUp Boiler p_nom": network.generators.p_nom["Backup"],
            "BackUp Boiler p_nom_opt": network.generators.p_nom_opt["Backup"],
            "CO2 Emissions Fossil Heat": network.carriers.co2_emissions["fossil heat"],
            "CO2 Limit": network.global_constraints.constant["CO2_Limit"],
            "CO2 Total Emissions" : co2_backup,       
            "Heat_Sink_p_nom_opt": network.links.p_nom_opt["Heat Sink CST"],
            "Heat Sink Capital Cost": network.links.capital_cost["Heat Sink CST"],
            "Heat Sink Marginal Cost": network.links.capital_cost["Heat Sink CST"],
            "Heat Sink Efficiency": network.links.efficiency["Heat Sink CST"],
            "TES e_nom": network.stores.e_nom["TES"],
            "TES e_nom_opt": network.stores.e_nom_opt["TES"],
            "TES Capital Cost": network.stores.capital_cost["TES"],
            "TES Marginal Cost": network.stores.marginal_cost["TES"],
            "TES e_initial": network.stores.e_initial["TES"],
            "Annual Energy CST": technology_sums["CST"],
            "Annual Energy PV": technology_sums["PV"],
            "Annual Energy WIND": technology_sums["Wind"],
            "Annual Energy PPA": technology_sums["PPA"],
            "Annual Energy Backup": technology_sums["Backup"],
            "ETES_Link p_nom": network.links.p_nom["ETES Thermal Power"],
            "ETES_Link p_nom_opt": network.links.p_nom_opt["ETES Thermal Power"],
            "ETES_Link Capital Cost": network.links.capital_cost["ETES Thermal Power"],
            "ETES_Link Marginal Cost": network.links.marginal_cost["ETES Thermal Power"],
            "ETES e_nom": network.stores.e_nom["ETES"],
            "ETES e_nom_opt": network.stores.e_nom_opt["ETES"],
            "ETES Capital Cost": network.stores.capital_cost["ETES"],
            "ETES Marginal Cost": network.stores.marginal_cost["ETES"]
        }
               
        collected_data[iteration + 1] = iteration_data

    # Convert collected data into a pandas DataFrame
    collected_data_df = pd.DataFrame.from_dict(collected_data, orient='index')

    # Create a DataFrame for the keys
    keys_df = pd.DataFrame({"Keys": collected_data_df.columns})
    
    # Transpose the collected data DataFrame to match keys with values
    collected_data_df = collected_data_df.transpose()

    # Concatenate the keys DataFrame and collected data DataFrame
    collected_data_df = pd.concat([keys_df, collected_data_df], axis=1)
    
    # Save collected data to a CSV file
    existing_files = [f for f in os.listdir(postpro_folder) if f.startswith("Iteration_export")]
    existing_indices = [int(f.split("_")[2].split(".")[0]) for f in existing_files if len(f.split("_")) == 3]
    if existing_indices:
        next_index = max(existing_indices) + 1
    else:
        next_index = 1

    output_filename = f"Iteration_export_{next_index}.csv"
    output_path = os.path.join(postpro_folder, output_filename)    
    collected_data_df.to_csv(output_path, index=False)
    return network
        
if __name__ == "__main__":
    xlsx_file = r"D:\OneDrive - Fichtner GmbH & Co. KG\Masterarbeit\Repository\Master-Thesis\OPTIMIZER\data\Setup.xlsx" # Füge hier den Pfad ein, der die CSV Setup Daten enthält
    output_folder = r"D:\OneDrive - Fichtner GmbH & Co. KG\Masterarbeit\Repository\Master-Thesis\OPTIMIZER\data\CSV_data" # FÜge hier den Pfad ein, wo die CSV Dateien gespeichert werden
    export_folder = r"D:\OneDrive - Fichtner GmbH & Co. KG\Masterarbeit\Repository\Master-Thesis\OPTIMIZER\data\Output"
    postpro_folder = r"D:\OneDrive - Fichtner GmbH & Co. KG\Masterarbeit\Repository\Master-Thesis\OPTIMIZER\data\Post_processing"
    selected_sheets = ["buses", "carriers", "generators","stores", "links","loads", "constraints", "generators-p_max_pu"]  # Füge hier die Namen der gewünschten Blätter hinzu
    selected_sheets_techno = ["Inputs"]
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
    
    # Initialize Costs
    network = calc_initial_costs(network, dataframes_dict)
    
    # Initial Optimization to produce first Solution for iterative Process
    network = optimize(network)
        
    export_components_to_excel(network, postpro_folder)
    
    co2_backup = calculate_co2_emissions(network)
    #Beginn Iterative Process to converge Solutions
    iterative_optimization(network, postpro_folder)
