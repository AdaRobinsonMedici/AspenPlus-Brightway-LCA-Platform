from aspen_utils import *
import pandas as pd
from aspen_processtools import *
from brightway_LCA import run_LCA

def main():

    # Preparation of simulation runs
    SimRunIndex = 1  # Start index of run
    SimRunNumber = 1  # Number of runs to perform

    # Open Aspen Plus file
    Aspen_Instance = Aspen_Plus_Interface()
    Aspen_Sim = Aspen_Plus_ProcessTools()
    Aspen_Instance.load_bkp(rf"Aspen_Plus_File/Post_combustion_solvent_based_MEA.bkp", 0, 1)
    time.sleep(2)

    ####################################################################################################################
    # SIMULATION START AND LOOPING
    ####################################################################################################################

    for run in range(SimRunIndex, SimRunNumber+1):

        print(f"------------------------------------- Run {SimRunNumber}--------------------------------------------\n")

        print("PROCESS SIMULATION IN ASPEN PLUS:")

        # Specify process parameter for variation
        print("1. Specify process parameter for variation")
        process_parameter_list = ["Runs", "FLUEGAS CO2 conc"]
        FluegasCO2conc = Aspen_Sim.get_fluegasCO2(SimRunIndex-1, process_parameter_list, "Sensitivity_GWP_Fluegas_CO2")

        # Adjust design specification for new value of process parameter
        print("2. Adjust design specification for simulation and start initial run with new values. Simulation running...")
        Aspen_Sim.design_spec(Aspen_Instance, FluegasCO2conc)

        print("3. Close MEA mass balance (by decreasing difference in mass flows in tear stroms below tolerance. Simulation running...")
        Aspen_Sim.check_MEAbalance_tearstreams(Aspen_Instance, MEA_tol=6)

        print("4. Close water mass balance (by decreasing difference in mass flows in tear stroms below tolerance. Simulation running...")
        Aspen_Sim.check_waterbalance_tearstreams(Aspen_Instance, H2O_tol=6)

        print("5. Check simulation status of final results")
        # Through adjusting the design specification and checking the component mass balance of MEA and H2O of tear streams
        # LEANMEA and LEANMEAC the simulation has to be run several times. Thus, at the end, it is checked whether the simulation
        # was successful or if there occurred a warning/ error.
        Aspen_Sim.check_simulation_status(Aspen_Instance, "Final simulation")

        print("6. Collect and export results to an Excel file for inspection (optional)")
        data = [
            Aspen_Sim.retrieve_ProcessParam(Aspen_Instance),
            Aspen_Sim.retrieve_Foreground(Aspen_Instance),
            Aspen_Sim.retrieve_NaturalRes(Aspen_Instance, 4.2, 5),
            Aspen_Sim.retrieve_EnergyFlow(Aspen_Instance),
            Aspen_Sim.retrieve_Infrastruct(Aspen_Instance, k=2000, d=2000, diameter=1)
        ]

        print("LCA CALCULATION IN BRIGHTWAY")

        # The LCA calculation is defined in just one function. Although this function involves many lines of code,
        # there are many repetitive tasks. It can be structured as follows:
        # 1. Setting up LCA by selecting project, databanks and LCIA method
        # 2. Receiving results from Aspen Plus process simulation and structure them according to the LCI
        # 3. Optional: Exporting Aspen Plus process simulation results to Excel for inspection
        # 4. Assign variables to certain activities of the biosphere3 and ecoinvent database for later use in populating
        #    the custom MEA_carbon_capture database
        # 5. Defining activities and exchanges
        # 6. LCA calculation and post-processing (e.g. extracting LCA results to an Excel file)

        run_LCA(data, run-1)

        print("Calculations finished.")

    Aspen_Instance.close_bkp()

    # KillAspen()

if __name__ == "__main__":
    main()
    