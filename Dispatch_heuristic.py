import datetime
import openpyxl
import pickle
import csv

# Dateipfade und -namen
file_path = 'D:/Eigene Dateien/Dokumente/02_Universität/RWTH/03_Masterarbeit/Zeitreihen/'
file_name_PV = "PV_125.xlsx"
file_name_CST = "CST_SM_1,1.xlsx"
file_name_PPA = "WIND_100.xlsx"
all_file_names = [file_name_CST, file_name_PV, file_name_PPA]

# Erstelle leeres Dictionary für die Generationen
generation_data = {}

# Generiere Daten für jede Generationstechnologie
for file_name in all_file_names:
    time_series = file_path + file_name
    workbook = openpyxl.load_workbook(time_series)
    worksheet = workbook.active

    if 'CST' in file_name:
        generation_data['CST'] = []
        for row in range(3, 8762):
            value = worksheet['B' + str(row)].value
            generation_data['CST'].append(value)
    elif 'PV' in file_name:
        generation_data['PV'] = []
        for row in range(3, 8762):
            value = worksheet['B' + str(row)].value
            generation_data['PV'].append(value)
    elif 'PPA' in file_name or 'WIND' in file_name:
        generation_data['PPA'] = []
        for row in range(3, 8762):
            value = worksheet['B' + str(row)].value
            generation_data['PPA'].append(value)
            
generation_PV = generation_data['PV'][:]
generation_PPA = generation_data['PPA'][:]
generation_CST = generation_data['CST'][:]

# Definition der Bedarfswerte für jede Stunde
hourly_demand = 500
demand_data = [hourly_demand for _ in range(0, 8759)]

# Definition des thermischen Speichers
TES_thermal_storage = 0
storage_capacity = 3000
# Priorität der Technologien
technology_priority = ['PPA', 'PV', 'CST']

def dispatch(demand_data, generation_data, storage_capacity, technology_priority):
    dispatched_CST = []
    dispatched_PV = []
    dispatched_PPA = []
    CST_dump = []
    PV_dump = []
    PPA_dump = []
    store_TES = []
    store_ETES = []
    TES_storage = [0]
    ETES_storage = [0]
    energy_purchase = []

    for demand in demand_data:
        generation_dict = {tech: generation_data[tech].pop(0) for tech in technology_priority}
        generation_sorted = sorted(generation_dict.items(), key=lambda x: technology_priority.index(x[0]))
        unmet_demand = demand
        total_generation = sum(generation_dict.values())
        stored_energy = 0

        for technology in technology_priority:
            generation = next(gen for tech, gen in generation_sorted if tech == technology)
            if total_generation >= demand:
                if generation >= unmet_demand:
                    dispatched_energy = unmet_demand
                    stored_energy = generation - unmet_demand
                    unmet_demand = 0
                    
                else:
                    dispatched_energy = generation
                    unmet_demand -= generation
                    stored_energy = 0
            else:
                dispatched_energy = generation
                total_generation -= generation
                unmet_demand -= generation
                stored_energy = 0 

            if technology == 'CST':
                if unmet_demand > 0:
                    dispatched_CST.append(dispatched_energy)
                    CST_dump.append(0)
                    store_TES = 0
                else:
                    dispatched_CST.append(dispatched_energy)
                    if TES_storage[-1] + stored_energy <= storage_capacity:
                        store_TES += stored_energy
                        CST_dump.append(0)
                    else:
                        store_TES += storage_capacity - TES_storage[-1]
                        CST_dump.append(generation - storage_capacity + TES_storage[-1])

            elif technology == 'PV':
                if unmet_demand > 0:
                    dispatched_PV.append(dispatched_energy)
                    PV_dump.append(0)
                    store_ETES = 0                
                else:
                    dispatched_PV.append(dispatched_energy)
                    if ETES_storage[-1] + stored_energy <= storage_capacity:
                        store_ETES += stored_energy
                        PV_dump.append(0)
                    else:
                        store_ETES += storage_capacity - ETES_storage[-1]
                        PV_dump.append(generation - storage_capacity + ETES_storage[-1])

            elif technology == 'PPA':
                if unmet_demand > 0:
                    dispatched_PPA.append(dispatched_energy)
                    PPA_dump.append(0)
                    store_ETES = 0
                else:
                    dispatched_PPA.append(dispatched_energy)
                    if ETES_storage[-1] + stored_energy <= storage_capacity:
                        store_ETES += stored_energy
                        PPA_dump.append(0)
                    else:
                        store_ETES += storage_capacity - ETES_storage[-1]
                        PPA_dump.append(generation - storage_capacity + ETES_storage[-1])
            
        if unmet_demand <= 0:
            ETES_storage.append(ETES_storage[-1] + store_ETES)
            TES_storage.append(TES_storage[-1] + store_TES)
            energy_purchase.append(unmet_demand)
            store_TES = 0
            store_ETES = 0
       
        if unmet_demand > 0:
            if  TES_storage[-1] >= unmet_demand:
                unstore_TES = unmet_demand
                unmet_demand = 0 
                TES_storage.append(TES_storage[-1] - unstore_TES)
                ETES_storage.append(ETES_storage[-1])
                energy_purchase.append(0)
                continue
            else: 
                unstore_TES = TES_storage[-1]
                unmet_demand -= unstore_TES
                TES_storage.append(0)
                    
            if  ETES_storage[-1] >= unmet_demand:
                unstore_ETES = unmet_demand
                ETES_storage.append(ETES_storage[-1] - unstore_ETES)
                energy_purchase.append(0)
            else: 
                unstore_ETES = ETES_storage[-1]
                ETES_storage.append((ETES_storage[-1] - unstore_ETES))
                unmet_demand -= unstore_ETES
                energy_purchase.append(unmet_demand)

    return (
        dispatched_CST, dispatched_PV, dispatched_PPA, CST_dump, PV_dump, PPA_dump, TES_storage, ETES_storage, energy_purchase
    )
    


# Aufruf der Dispatch-Funktion
dispatched_CST, dispatched_PV, dispatched_PPA, CST_dump, PV_dump, PPA_dump, TES_storage, ETES_storage, energy_purchase = dispatch(demand_data, generation_data, storage_capacity, technology_priority)

del TES_storage[0]
del ETES_storage[0]

# Ergebnisse speichern
results = {
    'dispatched_CST': dispatched_CST,
    'dispatched_PV': dispatched_PV,
    'dispatched_PPA': dispatched_PPA,
    'CST_dump': CST_dump,
    'PV_dump': PV_dump,
    'PPA_dump': PPA_dump,  
    'TES_storage': TES_storage,
    'ETES_storage': ETES_storage,
    'energy_purchase': energy_purchase,
    'CST_generation': generation_CST,
    'PV_generation': generation_PV,
    'PPA_generation': generation_PPA,
    'Demand': demand_data
}
csv_file = 'results.csv'
with open(csv_file, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(results.keys())
    writer.writerows(zip(*results.values()))
    

with open('dispatch_results.pkl', 'wb') as f:
    pickle.dump(results, f)



