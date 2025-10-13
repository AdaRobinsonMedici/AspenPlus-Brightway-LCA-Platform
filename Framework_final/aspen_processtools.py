from aspen_utils import *
import pandas as pd
from openpyxl import load_workbook

########################################################################################################################
# INPUT OF PROCESS PARAMETERS FOR EACH RUN
########################################################################################################################

# Molar mass of components in g/mol
Mol_MEA = 61
Mol_H2O = 18
Mol_CO2 = 44

MassConc_MEA = 0.3 # wt%

FluegasH2Oconc = 0.1125 # vol%
FluegasO2conc = 0.0381 # vol%

########################################################################################################################
# LOOPING SETTINGS
########################################################################################################################

# Checking convergence: Aspen gives these number as an output regarding the success of a simulation run. They are based
# on hexadecimal codes.
sim_success = 9345  # Hexadecimal: 0x00002481
sim_warning = 9348  # Hexadecimal: 0x00002484
sim_error = 9376    # Hexadecimal: 0x000024A0

# Preparation of simulation runs
SimRunIndex = 0     # Start index of run
SimRunNumber = 1    # Number of runs to perform

########################################################################################################################
# FUNCTIONS FOR ACCESSING ASPEN PLUS AND LOOPING
########################################################################################################################

class Aspen_Plus_ProcessTools:

    def get_fluegasCO2(self, SimRunIndex: int, column_names: list, process_param_filename ):
        df_InputParameters_names = column_names
        df_InputParameters = pd.read_excel(rf"{process_param_filename}.xlsx", names=df_InputParameters_names,
                                           skiprows=SimRunIndex, nrows=1)
        FluegasCO2conc = df_InputParameters.at[0, column_names[-1]] / 100

        return FluegasCO2conc

    def design_spec(self, Aspen_Plus, FluegasCO2conc):

        # Set fluegas CO2 conc. to evaluate new target value for design specification
        Aspen_Plus.Application.Tree.FindNode(r"\Data\Streams\FLUEGAS\Input\FLOW\MIXED\CO2").Value = FluegasCO2conc

        # Adjust N2 conc accordingly (concentration of H2O and O2 are kept constant.)
        FluegasN2conc = 1 - FluegasCO2conc - FluegasO2conc - FluegasH2Oconc
        Aspen_Plus.Application.Tree.FindNode(r"\Data\Streams\FLUEGAS\Input\FLOW\MIXED\N2").Value = FluegasN2conc

        # Run process simulation with adjusted flue gas composition
        Aspen_Plus.run_simulation()
        Aspen_Plus.check_run_completion()

        # Get new target value for design specification (CO2 removal efficiency is fixed at 95%)
        DS_TargetValue = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEGAS\Output\MASSFLOW\MIXED\CO2").Value * 0.95 * 3600

        # Set new target value in design specification
        Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Flowsheeting Options\Design-Spec\DS-1\Input\EXPR2").Value = DS_TargetValue

        # Run process simulation again with adjusted design specification
        Aspen_Plus.run_simulation()
        Aspen_Plus.check_run_completion()

    def check_MEAbalance_tearstreams(self, Aspen_Plus, MEA_tol):

        # MEA_tol in kg/h for defining tolerance for difference of water flows in tear streams LEANMEA and LEANMEAC
        print("  Tolerance for closing water balance: ", MEA_tol,
              " kg/h difference between tear streams LEANMEA and LEANMEAC")

        # Get loadings of stream LEANMEA
        LEANMEA_CO2Loading = Aspen_Plus.Application.Tree.FindNode(r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\CO2").Value
        LEANMEA_H2OLoading = Aspen_Plus.Application.Tree.FindNode(r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\H2O").Value

        # Get MEA mass flow rate in tear streams LEANMEA and LEANMEAC
        LEANMEA_MEA = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\LEANMEA\Output\STR_MAIN\MASSFLOW\MIXED\MEA").Value * 3600

        LEANMEAC_MEA = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\LEAMEAC\Output\STR_MAIN\MASSFLOW\MIXED\MEA").Value * 3600

        # Calculate difference of MEA mass flow
        delta_MEA = LEANMEA_MEA - LEANMEAC_MEA
        print("  Delta MEA: ", round(delta_MEA, 2))

        MEAbalance_counter = 0
        while delta_MEA > MEA_tol and MEAbalance_counter < 25 or delta_MEA < -MEA_tol and MEAbalance_counter < 25:
            MEAbalance_counter += 1
            print(f"  Check {MEAbalance_counter}:")

            # Adjust CO2 loading in LEANMEA to decrease/increase MEA mass flow
            print("  CO2 Loading: ", round(LEANMEA_CO2Loading, 2))
            if delta_MEA > 50:
                LEANMEA_CO2Loading = LEANMEA_CO2Loading + 0.01
                Aspen_Plus.Application.Tree.FindNode(
                    r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\CO2").Value = LEANMEA_CO2Loading
            elif 50 >= delta_MEA > 20:
                LEANMEA_CO2Loading = LEANMEA_CO2Loading + 0.005
                Aspen_Plus.Application.Tree.FindNode(
                    r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\CO2").Value = LEANMEA_CO2Loading
            elif 20 >= delta_MEA > 3:
                LEANMEA_CO2Loading = LEANMEA_CO2Loading + 0.002
                Aspen_Plus.Application.Tree.FindNode(
                    r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\CO2").Value = LEANMEA_CO2Loading
            elif -3 >= delta_MEA > -20:
                LEANMEA_CO2Loading = LEANMEA_CO2Loading - 0.0002
                Aspen_Plus.Application.Tree.FindNode(
                    r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\CO2").Value = LEANMEA_CO2Loading
            elif -20 > delta_MEA > -50:
                LEANMEA_CO2Loading = LEANMEA_CO2Loading - 0.005
                Aspen_Plus.Application.Tree.FindNode(
                    r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\CO2").Value = LEANMEA_CO2Loading
            else:
                LEANMEA_CO2Loading = LEANMEA_CO2Loading - 0.01
                Aspen_Plus.Application.Tree.FindNode(
                    r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\CO2").Value = LEANMEA_CO2Loading

            # Adjust H2O loading in LEANMEA based on LEANMEA_CO2Loading to achieve 30 wt% for all runs
            print("  H2O Loading: ", round(LEANMEA_H2OLoading, 2))
            LEANMEA_H2OLoading = (1 / Mol_H2O) * (Mol_MEA / MassConc_MEA - Mol_MEA - LEANMEA_CO2Loading * Mol_CO2)
            Aspen_Plus.Application.Tree.FindNode(
                r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\H2O").Value = LEANMEA_H2OLoading

            # Run simulation with modified values
            Aspen_Plus.run_simulation()
            Aspen_Plus.check_run_completion()

            # Get new values for control
            LEANMEA_CO2Loading_new = Aspen_Plus.Application.Tree.FindNode(
                r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\CO2").Value
            LEANMEA_H2OLoading_new = Aspen_Plus.Application.Tree.FindNode(
                r"\Data\Streams\LEANMEA\Input\FLOW\MIXED\H2O").Value
            print("  New CO2 loading: ", round(LEANMEA_CO2Loading_new, 2))
            print("  New H2O loading: ", round(LEANMEA_H2OLoading_new, 2))

            LEANMEA_MEA = Aspen_Plus.Application.Tree.FindNode(
                r"\Data\Streams\LEANMEA\Output\STR_MAIN\MASSFLOW\MIXED\MEA").Value * 3600
            LEANMEAC_MEA = Aspen_Plus.Application.Tree.FindNode(
                r"\Data\Streams\LEAMEAC\Output\STR_MAIN\MASSFLOW\MIXED\MEA").Value * 3600
            delta_MEA = LEANMEA_MEA - LEANMEAC_MEA
            print("  New delta MEA: ", round(delta_MEA, 2))

        print("  Closing MEA mass balance finished successfully.")

    def check_waterbalance_tearstreams(self, Aspen_Plus, H2O_tol):

        # H2O_tol in kg/h for defining tolerance for difference of water flows in tear streams LEANMEA and LEANMEAC
        print("  Tolerance for closing water balance: ", H2O_tol,
              " kg/h difference between tear streams LEANMEA and LEANMEAC")

        LEANMEA_H2O = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\LEANMEA\Output\STR_MAIN\MASSFLOW\MIXED\H2O").Value * 3600
        print("  LEANMEA water: ", round(LEANMEA_H2O, 2))

        LEANMEAC_H2O = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\LEAMEAC\Output\STR_MAIN\MASSFLOW\MIXED\H2O").Value * 3600
        print("  LEANMEAC water: ", round(LEANMEAC_H2O, 2))

        delta_H2O = LEANMEA_H2O - LEANMEAC_H2O
        print("  Delta water: ", round(delta_H2O, 2))

        H2Obalance_counter = 0
        while delta_H2O > H2O_tol and H2Obalance_counter < 25 or delta_H2O < -H2O_tol and H2Obalance_counter < 25:
            H2Obalance_counter += 1
            print(f"  Check {H2Obalance_counter}:")

            WATOUT = Aspen_Plus.Application.Tree.FindNode(r"\Data\Blocks\SPLIT\Input\BASIS_FLOW\WATOUT").Value
            MAKEUP = Aspen_Plus.Application.Tree.FindNode(r"\Data\Streams\MAKEUP\Input\FLOW\MIXED\H2O").Value

            if WATOUT > 0 and delta_H2O > 1:
                diff1 = WATOUT - delta_H2O
                if diff1 > 0:
                    Aspen_Plus.Application.Tree.FindNode(
                        r"\Data\Blocks\SPLIT\Input\BASIS_FLOW\WATOUT").Value = WATOUT - delta_H2O
                else:
                    Aspen_Plus.Application.Tree.FindNode(r"\Data\Blocks\SPLIT\Input\BASIS_FLOW\WATOUT").Value = 0
                    Aspen_Plus.Application.Tree.FindNode(
                        r"\Data\Streams\MAKEUP\Input\FLOW\MIXED\H2O").Value = MAKEUP + delta_H2O - WATOUT

            if MAKEUP > 0.01 and delta_H2O > 1:
                Aspen_Plus.Application.Tree.FindNode(
                    r"\Data\Streams\MAKEUP\Input\FLOW\MIXED\H2O").Value = MAKEUP + delta_H2O

            if WATOUT > 0 and delta_H2O < -1:
                Aspen_Plus.Application.Tree.FindNode(
                    r"\Data\Blocks\SPLIT\Input\BASIS_FLOW\WATOUT").Value = WATOUT - delta_H2O

            if MAKEUP > 0.01 and delta_H2O < -1:
                diff2 = MAKEUP + delta_H2O
                if diff2 > 0:
                    Aspen_Plus.Application.Tree.FindNode(
                        r"\Data\Streams\MAKEUP\Input\FLOW\MIXED\H2O").Value = MAKEUP + delta_H2O
                else:
                    Aspen_Plus.Application.Tree.FindNode(r"\Data\Streams\MAKEUP\Input\FLOW\MIXED\H2O").Value = 0
                    Aspen_Plus.Application.Tree.FindNode(
                        r"\Data\Blocks\SPLIT\Input\BASIS_FLOW\WATOUT").Value = WATOUT - delta_H2O - MAKEUP

            Aspen_Plus.run_simulation()
            Aspen_Plus.check_run_completion()

            LEANMEA_H2O = Aspen_Plus.Application.Tree.FindNode(
                r"\Data\Streams\LEANMEA\Output\STR_MAIN\MASSFLOW\MIXED\H2O").Value * 3600
            LEANMEAC_H2O = Aspen_Plus.Application.Tree.FindNode(
                r"\Data\Streams\LEAMEAC\Output\STR_MAIN\MASSFLOW\MIXED\H2O").Value * 3600
            delta_H2O = LEANMEA_H2O - LEANMEAC_H2O

            # Show results
            LEANMEA_H2O = Aspen_Plus.Application.Tree.FindNode(
                r"\Data\Streams\LEANMEA\Output\STR_MAIN\MASSFLOW\MIXED\H2O").Value * 3600
            print("  LEANMEA water: ", round(LEANMEA_H2O, 2))
            LEANMEAC_H2O = Aspen_Plus.Application.Tree.FindNode(
                r"\Data\Streams\LEAMEAC\Output\STR_MAIN\MASSFLOW\MIXED\H2O").Value * 3600
            print("  LEANMEAC water: ", round(LEANMEAC_H2O, 2))
            delta_H2O = LEANMEA_H2O - LEANMEAC_H2O
            print("  Delta water: ", round(delta_H2O, 2))

        print("  Closing water mass balance finished successfully.")


    def check_simulation_status(self, Aspen_Plus, run_type="non-defined simulation type"):
        """
        Checks the status of an Aspen Plus simulation and returns a list of status flags.

        Parameters:
        ----------
        run_type : str, optional
            Tag for the simulation run, e.g., 'designspec' or 'final'. Default is 'designspec'.

        Returns:
        ---------
        list : [success, warning_HX, other_warning, error_COND, other_error]
            List of boolean flags indicating different status outcomes.
        """

        # Initialize status flags for Aspen Plus simulation run
        success = False
        warning_HX = False
        other_warning = False
        error_COND = False
        other_error = False

        # Get the main simulation status
        run_status = Aspen_Plus.Application.Tree.FindNode(r"\Data").AttributeValue(12)

        if run_status == sim_success:
            print(f"  No errors encountered. {run_type.capitalize()} simulation finished.")
            success = True

        elif run_status == sim_warning:
            print(f"  Warnings encountered during {run_type} run. Checking where the warnings occur...")
            status_HX = Aspen_Plus.Application.Tree.FindNode(r"\Data\Blocks\HX").AttributeValue(12)

            if status_HX == sim_warning:
                print(f"  -> Warning in HX during {run_type} run.")
                warning_HX = True
            else:
                print(f"  -> Warning elsewhere during {run_type} run.")
                other_warning = True

        else:
            print(f"  Errors encountered during {run_type} run. Checking if the condenser is the issue.")
            other_error = True
            status_COND = Aspen_Plus.Application.Tree.FindNode(r"\Data\Blocks\COND").AttributeValue(12)

            if status_COND == sim_error:
                print("  -> Error at condenser. Trying to switch convergence solver.")
                error_COND = True
            else:
                print("  -> Some other error occurred.")

        return [success, warning_HX, other_warning, error_COND, other_error]

    def change_solver(self, Aspen_Plus, solver:str):

        Aspen_Plus.Application.Tree.FindNode(r"\Data\Convergence\Conv-Options\Input\TEAR_METHOD").Value = solver

        Aspen_Plus.run_simulation()
        Aspen_Plus.check_run_completion()

        solver_results = self.check_simulation_status(f"{solver}")

        if solver_results[-1] == False:
            print("  " + solver + " solved the error.")
        else:
            print("  Errors encountered." + solver + " could not solve the error.")

        return solver_results

    def retrieve_ProcessParam(self, Aspen_Plus):
        LEANMEA_MassFlow = Aspen_Plus.Application.Tree.FindNode(r"\Data\Streams\LEANMEA\Output\MASSFLMX\MIXED").Value
        LEANMEA_MassFlow = round(3600 * LEANMEA_MassFlow, 5)

        # Dataframe for process parameters to export data to Excel
        df_ProcessParameters = pd.DataFrame([[LEANMEA_MassFlow]]).T
        return df_ProcessParameters

    def retrieve_Foreground(self, Aspen_Plus):
        CO2OUT_CO2CaptureRate = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\CO2-OUT\Output\STR_MAIN\MASSFLOW\MIXED\CO2").Value
        CO2OUT_CO2CaptureRate = round(3600 * CO2OUT_CO2CaptureRate, 7)

        # FLUEGAS total flow rate in kg/h
        FLUEGAS_TotalMassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEGAS\Output\MASSFLMX\MIXED").Value
        FLUEGAS_TotalMassFlow = round(3600 * FLUEGAS_TotalMassFlow, 5)

        # FLUEGAS CO2 flow rate in kg/h
        FLUEGAS_CO2MassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEGAS\Output\STR_MAIN\MASSFLOW\MIXED\CO2").Value
        FLUEGAS_CO2MassFlow = round(3600 * FLUEGAS_CO2MassFlow, 5)

        # FLUEGAS N2 flow rate in kg/h
        FLUEGAS_N2MassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEGAS\Output\STR_MAIN\MASSFLOW\MIXED\N2").Value
        FLUEGAS_N2MassFlow = round(3600 * FLUEGAS_N2MassFlow, 5)

        # FLUEGAS O2 flow rate in kg/h
        FLUEGAS_O2MassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEGAS\Output\STR_MAIN\MASSFLOW\MIXED\O2").Value
        FLUEGAS_O2MassFlow = round(3600 * FLUEGAS_O2MassFlow, 5)

        # FLUEGAS H2O flow rate in kg/h
        FLUEGAS_H2OMassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEGAS\Output\STR_MAIN\MASSFLOW\MIXED\H2O").Value
        FLUEGAS_H2OMassFlow = round(3600 * FLUEGAS_H2OMassFlow, 5)

        # MAKEUP H2O flow rate in kg/h
        MAKEUP_H2OMassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\MAKEUP\Output\STR_MAIN\MASSFLOW\MIXED\H2O").Value
        MAKEUP_H2OMassFlow = round(3600 * MAKEUP_H2OMassFlow, 5)

        if MAKEUP_H2OMassFlow < 0.01:
            MAKEUP_H2OMassFlow = 0

        # MAKEUP MEA flow rate in kg/h
        MAKEUP_MEAMassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\MAKEUP\Output\STR_MAIN\MASSFLOW\MIXED\MEA").Value
        MAKEUP_MEAMassFlow = round(3600 * MAKEUP_MEAMassFlow, 5)

        # FLUEOFF total flow rate in kg/h
        FLUEOFF_TotalMassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEOFF\Output\MASSFLMX\MIXED").Value
        FLUEOFF_TotalMassFlow = round(3600 * FLUEOFF_TotalMassFlow, 5)

        # FLUEOFF CO2 flow rate in kg/h
        FLUEOFF_CO2MassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEOFF\Output\STR_MAIN\MASSFLOW\MIXED\CO2").Value
        FLUEOFF_CO2MassFlow = round(3600 * FLUEOFF_CO2MassFlow, 5)

        # FLUEOFF N2 flow rate in kg/h
        FLUEOFF_N2MassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEOFF\Output\STR_MAIN\MASSFLOW\MIXED\N2").Value
        FLUEOFF_N2MassFlow = round(3600 * FLUEOFF_N2MassFlow, 5)

        # FLUEOFF H2O flow rate in kg/h
        FLUEOFF_H2OMassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEOFF\Output\STR_MAIN\MASSFLOW\MIXED\H2O").Value
        FLUEOFF_H2OMassFlow = round(3600 * FLUEOFF_H2OMassFlow, 5)
        # FLUEOFF H2O is required in m³/h
        # Not able to extract partial density of H2O in FLUEOFF out of Aspen. Thus, a mean value is assumed.
        FLUEOFF_H2OPartialDensity = 0.705
        FLUEOFF_H2OVolumeFlow = round(FLUEOFF_H2OMassFlow / FLUEOFF_H2OPartialDensity, 5)

        # FLUEOFF MEA flow rate in kg/h
        FLUEOFF_MEAMassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\FLUEOFF\Output\STR_MAIN\MASSFLOW\MIXED\MEA").Value
        FLUEOFF_MEAMassFlow = round(3600 * FLUEOFF_MEAMassFlow, 5)

        # WASHWAT total flow rate in kg/h
        WASHWAT_TotalMassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\WASHWAT\Output\MASSFLMX\MIXED").Value
        WASHWAT_TotalMassFlow = round(3600 * WASHWAT_TotalMassFlow, 5)

        # WATOUT flow rate in kg/h
        WATOUT_Massflow = Aspen_Plus.Application.Tree.FindNode(r"\Data\Streams\WATOUT\Output\MASSFLMX\MIXED").Value
        WATOUT_Massflow = round(3600 * WATOUT_Massflow, 5)

        # WATOUT MEA flow rate in kg/h
        WATOUT_MEAMassFlow = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\WATOUT\Output\STR_MAIN\MASSFLOW\MIXED\MEA").Value
        WATOUT_MEAMassFlow = round(3600 * WATOUT_MEAMassFlow, 5)

        # Dataframe for material flow to export data to Excel
        array_MaterialFlow = np.array(
            [[CO2OUT_CO2CaptureRate, FLUEGAS_TotalMassFlow, FLUEGAS_CO2MassFlow, FLUEGAS_N2MassFlow,
                FLUEGAS_O2MassFlow, FLUEGAS_H2OMassFlow, MAKEUP_H2OMassFlow, MAKEUP_MEAMassFlow, FLUEOFF_TotalMassFlow,
                FLUEOFF_CO2MassFlow, FLUEOFF_N2MassFlow, FLUEOFF_H2OVolumeFlow, FLUEOFF_MEAMassFlow,
                WASHWAT_TotalMassFlow,
                WATOUT_Massflow, WATOUT_MEAMassFlow]])
        df_MaterialFlow = pd.DataFrame(array_MaterialFlow).T
        return df_MaterialFlow

    def retrieve_EnergyFlow(self, Aspen_Plus):

        # Reboiler duty in kW
        STRIPPER_ReboilerDuty = Aspen_Plus.Application.Tree.FindNode(r"\Data\Blocks\REBOILER\Output\QCALC").Value
        STRIPPER_ReboilerDuty = round(1e-03 * STRIPPER_ReboilerDuty, 5)
        CO2OUT_CO2CaptureRate = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Streams\CO2-OUT\Output\STR_MAIN\MASSFLOW\MIXED\CO2").Value
        CO2OUT_CO2CaptureRate = round(3600 * CO2OUT_CO2CaptureRate, 7)

        try:
            STRIPPER_SpecReboilerDuty = round((STRIPPER_ReboilerDuty / (CO2OUT_CO2CaptureRate / 3600)), 5)
        except ZeroDivisionError:
            return False

        # Electric power for pump compression in kW
        PUMP_Elec = round(1e-03 * Aspen_Plus.Application.Tree.FindNode(r"\Data\Blocks\PUMP\Output\ELEC_POWER").Value,
                            5)

        # Total electricity
        Total_Elec = round(6 * PUMP_Elec, 5)

        # Dataframe for energy flow to export data to Excel
        array_EnergyFlow = np.array([[STRIPPER_SpecReboilerDuty, Total_Elec, PUMP_Elec]])
        df_EnergyFlow = pd.DataFrame(array_EnergyFlow).T
        return df_EnergyFlow

    def retrieve_NaturalRes(self, Aspen_Plus, cp, dT):
        # Assumptions
        # cp - Specific heat capacity of water in kJ/kgK
        # dT - Temperature difference between cooling water inlet and outlet temperature in K

        # COOLER Cooling duty in kW
        COOLER_CoolingDuty = Aspen_Plus.Application.Tree.FindNode(r"\Data\Blocks\COOLER\Output\QCALC").Value
        COOLER_CoolingDuty = round(-1e-03 * COOLER_CoolingDuty, 5)
        # COOLER Cooling water in kg/h
        COOLER_CoolingWater = round((COOLER_CoolingDuty / (cp * dT)) * 3600, 2)

        # COND Cooling duty in kW
        COND_CoolingDuty = Aspen_Plus.Application.Tree.FindNode(r"\Data\Blocks\COND\Output\QCALC").Value
        COND_CoolingDuty = round(-1e-03 * COND_CoolingDuty, 5)
        # COND Cooling water in kg/h
        COND_CoolingWater = round((COND_CoolingDuty / (cp * dT)) * 3600, 5)

        # Dataframe for natural resources to export data to Excel
        array_NaturalResources = np.array([[COOLER_CoolingWater, COND_CoolingWater]])
        df_NaturalResources = pd.DataFrame(array_NaturalResources).T
        return df_NaturalResources

    def retrieve_Infrastruct(self, Aspen_Plus, k, d, diameter=0.5):

        # Assumptions
        # Literature assumes the column material between 50000 and 100000 kg for a standard size column
        # As the process parameters of a pilot plant are used the column will be sized smaller than a standard size column
        # Thus, 50000 kg is assumed to be the maximum value
        # As only the height of the columns is varied, a linear relation is assumed between the column height and the required material
        # Minimum stage (STRIPPER): 13, Maximum stage (ABSORBER): 24
        # Assumption: Column with 24 stages: 50000 kg, column with 13 stages: 28000 kg
        # Linear equation: e. g. k = 2000, d = 2000, e. g. column with 18 m height: 18 * 2000 + 2000 = 38000 kg

        # k - slope of linear curve
        # d - vertical distance of linear curve
        # diameter - Diameter of column

        ##### Calculation of required column material ######

        # In this example the packing type is always Mellapak. However, you could change it in the simulation and
        # adjust it here accordingly
        PackingType = Aspen_Plus.Application.Tree.FindNode(
            r"\Data\Blocks\ABSORBER\Input\CA_PACKTYPE\OPT-R\P-1").Value

        # Number of stages absorber and its height (1 stage equal 1 m height)
        ABSORBER_NStages = Aspen_Plus.Application.Tree.FindNode(r"\Data\Blocks\ABSORBER\Input\NSTAGE").Value
        ABSORBER_Height = ABSORBER_NStages

        # Number of stages stripper and its height (1 stage equal 1 m height)
        STRIPPER_NStages = Aspen_Plus.Application.Tree.FindNode(r"\Data\Blocks\STRIPPER\Input\NSTAGE").Value
        STRIPPER_Height = STRIPPER_NStages

        ABSORBER_ColumnMaterial = ABSORBER_Height * k + d
        STRIPPER_ColumnMaterial = STRIPPER_Height * k + d

        ##### Calculation of required packing material ######

        # Define void fractions for each packing type (Values from Aspen)
        void_Mellapak = 0.987
        void_Flexipac = 0.91
        void_BX = 0.9

        Density_Steel = 7850  # Mean density of steel in kg/dm³

        # Volume of columns in m³
        ABSORBER_TotalVol = (diameter ** 2 * np.pi / 4) * ABSORBER_Height
        STRIPPER_TotalVol = (diameter ** 2 * np.pi / 4) * STRIPPER_Height

        # Check which packing is used and calculate volume of packing
        ABSORBER_PackingVol = 0
        STRIPPER_PackingVol = 0

        if PackingType == "MELLAPAK":
            PackingType = "Mellapak 250Y"
            ABSORBER_PackingVol = (1 - void_Mellapak) * ABSORBER_TotalVol
            STRIPPER_PackingVol = (1 - void_Mellapak) * STRIPPER_TotalVol

        elif PackingType == "FLEXIPAC":
            PackingType = "Flexipac 1Y"
            ABSORBER_PackingVol = (1 - void_Flexipac) * ABSORBER_TotalVol
            STRIPPER_PackingVol = (1 - void_Flexipac) * STRIPPER_TotalVol

        else:
            PackingType = "BX"
            ABSORBER_PackingVol = (1 - void_BX) * ABSORBER_TotalVol
            STRIPPER_PackingVol = (1 - void_BX) * STRIPPER_TotalVol

        # Calculate packing material
        ABSORBER_PackingMaterial = round(ABSORBER_PackingVol * Density_Steel, 5)
        STRIPPER_PackingMaterial = round(STRIPPER_PackingVol * Density_Steel, 5)

        # Dataframe for infrastructure to export data to Excel
        array_Infrastructure = np.array(
            [[ABSORBER_ColumnMaterial, ABSORBER_PackingMaterial, STRIPPER_ColumnMaterial,
                STRIPPER_PackingMaterial]])
        df_Infrastructure = pd.DataFrame(array_Infrastructure).T
        return df_Infrastructure



