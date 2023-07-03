import pickle
import matplotlib.pyplot as plt

# Laden der Listen aus einer Datei
def load_data(filename):
    with open(filename, 'rb') as file:
        data = pickle.load(file)
    return data

loaded_data = load_data('D:/Eigene Dateien/Programming/01_Projekte/Master Thesis/dispatch_results.pkl')
for key, value in loaded_data.items():
    if key == 'dispatched_CST':
        dispatched_CST = value
    elif key == 'dispatched_PV':
        dispatched_PV = value
    elif key == 'dispatched_PPA':
        dispatched_PPA = value
    elif key == 'energy_purchase':
        energy_purchase = value
    elif key == 'CST_dump':
        CST_dump = value
    elif key == 'PV_dump':
        PV_dump = value
    elif key == 'PPA_dump':
        PPA_dump = value
    elif key == 'TES_storage':
        TES_storage = value
    elif key == 'ETES_storage':
        ETES_storage = value
    elif key == 'CST_generation':
        CST_generation = value
    elif key == 'PV_generation':
        PV_generation = value
    elif key == 'PPA_generation':
        PPA_generation = value
    elif key == 'Demand':
        Demand = value

def plot_time_series(dispatched_CST, dispatched_PV, dispatched_PPA, energy_purchase, CST_dump, PV_dump, PPA_dump, TES_storage, ETES_storage, CST_generation, PV_generation, PPA_generation, Demand):
    # Die Range der X Werte "0, 8760"
    x_range = (0, 8760)

    # Plotten der Zeitreihen separat
    fig, axes = plt.subplots(5, 3, figsize=(12, 15))

    # Plot 1: Dispatched CST Generation
    axes[0, 0].plot(dispatched_CST)
    axes[0, 0].set_xlabel('Time')
    axes[0, 0].set_ylabel('Energy')
    axes[0, 0].set_title('Dispatched CST Generation')
    if x_range:
        axes[0, 0].set_xlim(*x_range)
        axes[0, 0].set_ylim(0, max(dispatched_CST))

    # Plot 2: Dispatched PV Generation
    axes[0, 1].plot(dispatched_PV)
    axes[0, 1].set_xlabel('Time')
    axes[0, 1].set_ylabel('Energy')
    axes[0, 1].set_title('Dispatched PV Generation')
    if x_range:
        axes[0, 1].set_xlim(*x_range)
        axes[0, 1].set_ylim(0, max(dispatched_PV))

    # Plot 3: Dispatched PPA Generation
    axes[0, 2].plot(dispatched_PPA)
    axes[0, 2].set_xlabel('Time')
    axes[0, 2].set_ylabel('Energy')
    axes[0, 2].set_title('Dispatched PPA Generation')
    if x_range:
        axes[0, 2].set_xlim(*x_range)
        axes[0, 2].set_ylim(0, max(dispatched_PPA))

    # Plot 4: Energy Purchase
    axes[1, 0].plot(energy_purchase)
    axes[1, 0].set_xlabel('Time')
    axes[1, 0].set_ylabel('Energy')
    axes[1, 0].set_title('Energy Purchase')
    if x_range:
        axes[1, 0].set_xlim(*x_range)
        axes[1, 0].set_ylim(0, max(energy_purchase))

    # Plot 5: CST Dump
    axes[1, 1].plot(CST_dump)
    axes[1, 1].set_xlabel('Time')
    axes[1, 1].set_ylabel('Energy')
    axes[1, 1].set_title('CST Dump')
    if x_range:
        axes[1, 1].set_xlim(*x_range)
        axes[1, 1].set_ylim(0, max(CST_dump))

    # Plot 6: PV Dump
    axes[1, 2].plot(PV_dump)
    axes[1, 2].set_xlabel('Time')
    axes[1, 2].set_ylabel('Energy')
    axes[1, 2].set_title('PV Dump')
    if x_range:
        axes[1, 2].set_xlim(*x_range)
        axes[1, 2].set_ylim(0, max(PV_dump))

    # Plot 7: PPA Dump
    axes[2, 0].plot(PPA_dump)
    axes[2, 0].set_xlabel('Time')
    axes[2, 0].set_ylabel('Energy')
    axes[2, 0].set_title('PPA Dump')
    if x_range:
        axes[2, 0].set_xlim(*x_range)
        axes[2, 0].set_ylim(0, max(PPA_dump))


    # Plot 10: Storage
    axes[3, 0].plot(TES_storage)
    axes[3, 0].set_xlabel('Time')
    axes[3, 0].set_ylabel('Energy')
    axes[3, 0].set_title('TES_Storage')
    if x_range:
        axes[3, 0].set_xlim(*x_range)
        axes[3, 0].set_ylim(0, max(TES_storage))

    # Plot 11: ETES Storage
    axes[3, 1].plot(ETES_storage)
    axes[3, 1].set_xlabel('Time')
    axes[3, 1].set_ylabel('Energy')
    axes[3, 1].set_title('ETES Storage')
    if x_range:
        axes[3, 1].set_xlim(*x_range)
        axes[3, 1].set_ylim(0, max(ETES_storage))

    # Plot 12: CST Generation
    axes[3, 2].plot(CST_generation)
    axes[3, 2].set_xlabel('Time')
    axes[3, 2].set_ylabel('Energy')
    axes[3, 2].set_title('CST Generation')
    if x_range:
        axes[3, 2].set_xlim(*x_range)
        axes[3, 2].set_ylim(0, max(CST_generation))

    # Plot 13: PV Generation
    axes[4, 0].plot(PV_generation)
    axes[4, 0].set_xlabel('Time')
    axes[4, 0].set_ylabel('Energy')
    axes[4, 0].set_title('PV Generation')
    if x_range:
        axes[4, 0].set_xlim(*x_range)
        axes[4, 0].set_ylim(0, max(PV_generation))

    # Plot 14: PPA Generation
    axes[4, 1].plot(PPA_generation)
    axes[4, 1].set_xlabel('Time')
    axes[4, 1].set_ylabel('Energy')
    axes[4, 1].set_title('PPA Generation')
    if x_range:
        axes[4, 1].set_xlim(*x_range)
        axes[4, 1].set_ylim(0, max(PPA_generation))

    # Anpassen des Layouts und Anzeigen der Plots
    fig.tight_layout()
    plt.show()


plot_time_series(dispatched_CST, dispatched_PV, dispatched_PPA, energy_purchase, CST_dump, PV_dump, PPA_dump,TES_storage, ETES_storage, CST_generation, PV_generation, PPA_generation, Demand)
