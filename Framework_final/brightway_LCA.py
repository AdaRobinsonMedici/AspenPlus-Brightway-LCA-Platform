
from brightway_utilis import *
import bw2data as bd
import bw2io as bi
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

def run_LCA(sim_data, i=0, export_excel=False):

    #######################################################################################################################
    # LCA SETUP
    #######################################################################################################################

    print("1. Start calculating LCA results by Brightway")

    selected_project = "project_E2DT"
    bd.projects.set_current(name=selected_project)
    print("Current project: ", bd.projects.current)

    # Assign names to databases created in advance in initial setup
    ecoinvent_version_cutoff = "ecoinvent_cutoff_391"
    ecoinvent_version_biosphere = "biosphere3"
    db_biosphere3 = bd.Database(ecoinvent_version_biosphere)
    db_ecoinvent = bd.Database(ecoinvent_version_cutoff)

    # Function for clearing LCA results (if you want to start from scratch)
    # clear_LCA_excel_output("LCA_results_rawfile")

    # Choose LCIA method
    method_input = "ReCiPe 2016 v1.03 midpoint (E) global warming potential"

    # For each simulation run: Delete old database MEA_Carbon_Capture (It is recreated below and populated with new
    # Aspen Plus simulation results
    try:
        del bd.databases["MEA_Carbon_Capture"]
    except KeyError:
        pass   

    # Create new database MEA Carbon Capture for new run
    my_database = bd.Database("MEA_Carbon_Capture")
    my_database.register()

    print("a) Databases loaded: ")
    print(list(bd.databases))

    #######################################################################################################################
    # VALUES FROM PYTHON INTERFACE FOR EXCHANGES OF ACTIVITIES OF MEA CARBON CAPTURE DATABASE
    #######################################################################################################################

    print("b) Importing LCI for this run from Aspen Plus results.")

    # --- Get Process Parameters ---
    df_ProcessParameters = pd.DataFrame(sim_data[0].iloc[:, i].astype(float).tolist()).T
    df_ProcessParameters.columns = ["Solvent flow rate"]

    # Plant normalisation:
    total_hours = 216000
    df_ProcessParameters["Norm Solvent flow rate"] = (
        df_ProcessParameters["Solvent flow rate"] / (sim_data[1].iloc[0, i] * total_hours)
    )
    #print(df_ProcessParameters, "\n")

    # --- Get Material Flows ---
    df_MaterialFlow = pd.DataFrame(sim_data[1].iloc[:, i].astype(float).tolist()).T
    df_MaterialFlow.columns = [
        "CO2 captured", "FLUEGAS total", "FLUEGAS CO2", "FLUEGAS N2", "FLUEGAS O2",
        "FLUEGAS H2O", "MAKEUP Water", "MAKEUP Amine", "FLUEOFF total", "FLUEOFF CO2",
        "FLUEOFF N2", "FLUEOFF H2O", "FLUEOFF MEA", "WASHWAT in", "WATOUT out water", "WATOUT out MEA"
    ]
    #print(df_MaterialFlow, "\n")

    # --- Get Natural Resources ---
    df_NaturalResources =  pd.DataFrame(sim_data[2].iloc[:, i].astype(float).tolist()).T
    df_NaturalResources.columns = ["Cooling water COOLER", "Cooling water COND"]
    #print(df_NaturalResources, "\n")

    # --- Get Energy Flow (Background) ---
    df_EnergyFlow = pd.DataFrame(sim_data[3].iloc[:, i].astype(float).tolist()).T
    df_EnergyFlow.columns = ["Reboiler steam in MJ/ton CO2", "Total electr.", "Pump electr."]
    
    df_EnergyFlow["Reboiler steam in kWh"] = (
        df_EnergyFlow.iloc[0, 0] * df_MaterialFlow.iloc[0, 0] * 0.001 / 3.6
    )
    df_EnergyFlow["Total energy"] = (
        df_EnergyFlow.iloc[0, 1] / (df_MaterialFlow.iloc[0, 0] * 0.001 / 3600)
    )
    #print(df_EnergyFlow, "\n")

    # --- Get Equipment ---
    df_equipment = pd.DataFrame(np.array([6.0, 1.0, 1.0, 1.0, 2.0])).T
    df_equipment.columns = ["Pumps", "Reboiler", "Compressor", "Condenser", "Cooler"]
    #print(df_equipment, "\n")

    # --- Get Infrastructure ---
    df_Infrastructure = pd.DataFrame(sim_data[4].iloc[:, i].astype(float).tolist()).T
    df_Infrastructure.columns = [
        "Absorber Column", "Absorber Packing", "Desorber Column", "Desorber Packing"
    ]
    df_Infrastructure["Norm Carbon capture plant"] = 1 / (
        df_MaterialFlow.iloc[0, 0] * total_hours
    )
    #print(df_Infrastructure, "\n")

    ######################################################################################################################
    # EXPORTING INVENTORY DATA TO EXCEL
    ######################################################################################################################

    if export_excel == True:
        StartRow_Categories=[]
        StartRow_Categories.append(search_row(column="B", search_value="Process Parameters") + 5)
        StartRow_Categories.append(search_row(column="B", search_value="Material Flow (Foreground System)") + 7)
        StartRow_Categories.append(search_row(column="B", search_value="Natural Resources (Background System)") + 6)
        StartRow_Categories.append(search_row(column="B", search_value="Energy Flow (Background)") + 6)
        StartRow_Categories.append(search_row(column="B", search_value="Infrastructure") + 19)

        Rows_Categories = [x + i for x in StartRow_Categories]

        with pd.ExcelWriter(r"LCI_raw.xlsx", mode="a", engine="openpyxl",
                            if_sheet_exists="overlay") as writer:
            df_ProcessParameters.to_excel(writer, header=False, sheet_name="Scenario_1_base case", index=False,
                                            startrow=Rows_Categories[0], startcol=2)
            df_MaterialFlow.to_excel(writer, header=False, sheet_name="Scenario_1_base case", index=False,
                                        startrow=Rows_Categories[1], startcol=2)
            df_NaturalResources.to_excel(writer, header=False, sheet_name="Scenario_1_base case", index=False,
                                            startrow=Rows_Categories[2], startcol=2)
            df_EnergyFlow.to_excel(writer, header=False, sheet_name="Scenario_1_base case", index=False,
                                    startrow=Rows_Categories[3], startcol=2)
            df_Infrastructure.to_excel(writer, header=False, sheet_name="Scenario_1_base case", index=False,
                                        startrow=Rows_Categories[4], startcol=3)

    ########################################################################################################################
    # GET INPUTS OF DATABASES BIOSPHERE3 AND ECOINVENT TO DEFINE EXCHANGES OF ACTIVITIES
    ########################################################################################################################

    print("d) Get data from ecoinvent and biosphere3 database to define exchanges.")

    # Data (activities) from database biosphere 3 for Carbon dioxide, Nitrogen, Oxygen, Water, Monoethanolamine
    biosphere3_act_CarbonDioxide_airurban = db_biosphere3.get(code="aa7cac3a-3625-41d4-bc54-33e2cf11ec46")
    biosphere3_act_Nitrogen_air = db_biosphere3.get(code="e5ea66ee-28e2-4e9b-9a25-4414551d821c")
    biosphere3_act_Water_air = db_biosphere3.get(code="075e433b-4be4-448e-9510-9a5029c1ce94")
    biosphere3_act_Monoethanolamine_air = db_biosphere3.get(code="9f550c00-0f0f-45e7-a29c-a450752d4c7a")

    # Data (activities) from database ecoinvent:

    # Option 1: How to look for activities in a database
    # print(db_ecoinvent.search("steel, low-alloyed", filter={'location': 'GLO', 'unit': 'kilogram',} )[0].as_dict()['code'])
    # print(inspect.getsource(db_ecoinvent.search))
    # Filter location, keyword and second keyword, unit, product or market
    # ecoinvent_act_low_alloyed = db_ecoinvent.search("market for steel, low-alloyed")[0]
    # ecoinvent_act_unalloyed =  db_ecoinvent.search("market for steel, unalloyed")[0]

    # Option 2: How to look for activities in a database
    # print(db_ecoinvent.search("market for steel"))
    # print(db_ecoinvent.search("market for steel, unalloyed")[0].as_dict()['code'])

    # Market for steel, low-alloyed and unalloy for packings (low-alloyed) and columns (unalloyed)
    ecoinvent_act_low_alloyed = db_ecoinvent.get(code="a0bd454544d5b0919de48edec5c458e7")
    ecoinvent_act_unalloyed = db_ecoinvent.get(code="d872e0d78319cb13e12b96de83e19dd7")

    # Reboiler: heat production, natural gas, at industrial furnace >100kW
    ecoinvent_act_heatprod = db_ecoinvent.get(code="b74c8625e0adf3b76b6a1aafed10312b")

    # Market for monoethanolamine
    ecoinvent_act_marketMEA = db_ecoinvent.get(code="1d95a9865e5ecb049eb64f6823865078")

    # Market for tap water
    ecoinvent_act_markettapwater = db_ecoinvent.get(code="4fe1f2dae4830593ee2d608cb2d5ff2c")

    # Market for waste water, average
    ecoinvent_act_marketwastewater = db_ecoinvent.get(code="ef9e2b6a0815008d49a77388e7c5b0e8")

    # Market group for electricity, medium voltage
    ecoinvent_act_marketelectr = db_ecoinvent.get(code="476b68dc744251f288055e5ce8264feb")

    # Market for water, deionised
    ecoinvent_act_marketwaterdeion = db_ecoinvent.get(code="95d26245bc411de42cbca001406f6131")

    # Water production, deionsed
    ecoinvent_act_waterproddeion = db_ecoinvent.get(code="395d665bd9b8ce065b590bd96fd74ca9")

    # Market for air compressor, screw-type compressor, 300kW
    ecoinvent_act_compr = db_ecoinvent.get(code="fdabbf27abc2d302066ebce8affec56a")

    # Market for borehole heat exchanger, 150m
    ecoinvent_act_heatex = db_ecoinvent.get(code="5d68a7fd42b989d214867478b999a04c")

    # Market for gas boiler
    ecoinvent_act_gasboil = db_ecoinvent.get(code="976694b4c0b31641b846a89d2fbf6c4b")

    # Market for pump, 40W
    ecoinvent_act_pump = db_ecoinvent.get(code="1358cde069b0d7df0f29529d19c6f900")

    ######################################################################################################################
    # DEFINE ACTIVITIES OF DATABASE MEA CARBON CAPTURE
    ######################################################################################################################

    print("c) Creating MEA carbon capture database: activities.")

    # Activity: Absorber
    data_abs = {
        'name': 'Absorber',
        'code': '1',
        'location': 'GLO',
        'reference product': 'Absorber with Packing',
        'type': 'process',
        'unit': 'unit',
    }
    act_abs = my_database.new_activity(**data_abs)
    act_abs.save()

    # Activity: Desorber
    data_des = {
        'name': 'Desorber',
        'code': '2',
        'location': 'GLO',
        'reference product': "Desorber with Packing",
        'type': "process",
        'unit': 'unit',
    }
    act_des = my_database.new_activity(**data_des)
    act_des.save()

    # Activity: MEA production
    data_MEAprod = {
        'name': 'MEA production',
        'code': '4',
        'location': 'GLO',
        'reference product': "MEA production",
        'type': "process",
        'unit': 'kg',
    }
    act_MEAprod = my_database.new_activity(**data_MEAprod)
    act_MEAprod.save()

    # Activity:Infrastructure Absorption/ Desorption
    data_InfraAbsDes = {
        'name': 'inf_absorption_desorption',
        'code': '5',
        'location': 'GLO',
        'reference product': 'Carbon Capture Plant',
        'type': 'process',
        'unit': 'unit',
    }
    act_InfraAbsDes = my_database.new_activity(**data_InfraAbsDes)
    act_InfraAbsDes.save()

    # Activity: Production Absorption/ Desorption
    data_ProAbsDes = {
        'name': 'pro_absorption_desorption',
        'code': '6',
        'location': 'GLO',
        'reference product': 'CO2 produced',
        'type': 'process',
        'unit': 'kg/h',
    }
    act_ProAbsDes = my_database.new_activity(**data_ProAbsDes)
    act_ProAbsDes.save()

    ########################################################################################################################
    # CREATE EXCHANGES FOR EACH ACTIVITY AND SET VALUES FOR CURRENT RUN in MEA CARBON CAPTURE DATABASE
    ########################################################################################################################

    ############################################# Exchanges for activity Absorber ##########################################

    print("e) Creating MEA carbon capture database: exchanges.")

    # Exchange 1: Absorber
    act_abs.new_exchange(
        name = "Absorber",
        input = act_abs,
        amount = 1,
        unit = "unit",
        type = "production",
    ).save()

    # Exchange 2: market for steel, low-alloyed
    act_abs.new_exchange(
        name = "market for steel, low-alloyed",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_low_alloyed['code']}"),
        amount = float(df_Infrastructure.at[0, "Absorber Packing"]),
        unit = "kilogram/kilogram CO2",
        type = "technosphere",
    ).save()

    # Exchange 3: market for steel, unalloyed
    act_abs.new_exchange(
        name = "market for steel, unalloyed",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_unalloyed['code']}"),
        amount = float(df_Infrastructure.at[0, "Absorber Column"]),
        unit = "kilogram/kilogram CO2",
        type = "technosphere",
    ).save()

    ########################################### Exchanges for activity Desorber ############################################

    # Exchange 1: Desorber
    act_des.new_exchange(
        name = "Desorber",
        input = act_des,
        amount = 1,
        unit = "unit",
        type = "production",
    ).save()

    # Exchange 2: market for steel, low-alloyed
    act_des.new_exchange(
        name = "market for steel, low-alloyed",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_low_alloyed['code']}"),
        amount = float(df_Infrastructure.at[0, "Desorber Packing"]),
        unit = "kilogram/kilogram CO2",
        type = "technosphere",
    ).save()

    # Exchange 3: market for steel, unalloyed
    act_des.new_exchange(
        name = "market for steel, unalloyed",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_unalloyed['code']}"),
        amount = float(df_Infrastructure.at[0, "Desorber Column"]),
        unit = "kilogram/kilogram CO2",
        type = "technosphere",
    ).save()

    ######################################## Exchanges for activity MEA production #########################################

    # Exchange 1: MEA production
    act_MEAprod.new_exchange(
        name = "MEA production",
        input = act_MEAprod,
        amount = 1,
        unit = "kilogram",
        type = "production",
    ).save()

    # Exchange 2: market for monoethanolamine
    act_MEAprod.new_exchange(
        name = "market for monoethanolamine",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_marketMEA['code']}"),
        amount = 0.3,
        unit = "kilogram",
        type = "technosphere",
    ).save()

    # Exchange 3: market for water, deionised
    act_MEAprod.new_exchange(
        name = "market for water, deionised",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_marketwaterdeion['code']}"),
        amount = 0.7,
        unit = "kilogram",
        type = "technosphere",
    ).save()

    ######################################## Exchanges for activity InfraAbsDes #########################################

    # Exchange 1: InfraAbsDes
    act_InfraAbsDes.new_exchange(
        name = "InfraAbsDes",
        input = act_InfraAbsDes,
        amount = 1,
        unit = "unit",
        type = "production",
    ).save()

    # Exchange 2: Absorber
    act_InfraAbsDes.new_exchange(
        name = "Absorber",
        input = act_abs,
        amount = 1,
        unit = "unit",
        type = "technosphere",
    ).save()

    # Exchange 3: Desorber
    act_InfraAbsDes.new_exchange(
        name = "Desorber",
        input = act_des,
        amount = 1,
        unit = "unit",
        type = "technosphere",
    ).save()

    # Exchange 4: Market for air compressor, screw-type compressor, 300kW
    act_InfraAbsDes.new_exchange(
        name = "market for air compressor, screw-type compressor, 300kW",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_compr['code']}"),
        amount = 1,
        unit = "unit",
        type = "technosphere",
    ).save()

    # Exchange 5: Market for borehole heat exchanger, 150m
    act_InfraAbsDes.new_exchange(
        name = "Market for borehole heat exchanger, 150m",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_heatex['code']}"),
        amount = 3,
        unit = "unit",
        type = "technosphere",
    ).save()

    # Exchange 6: Market for gas boiler
    act_InfraAbsDes.new_exchange(
        name = "Market for gas boiler",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_gasboil['code']}"),
        amount = 1,
        unit = "unit",
        type = "technosphere",
    ).save()

    # Exchange 7: Market for pump, 40W
    act_InfraAbsDes.new_exchange(
        name = "Market for pump, 40W",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_pump['code']}"),
        amount = 6,
        unit = "unit",
        type = "technosphere",
    ).save()

    ######################################## Exchanges for activity ProAbsDes #########################################

    # Exchange 1: Carbon dioxide, fossil
    act_ProAbsDes.new_exchange(
        name = "Carbon dioxide, fossil",
        input = (ecoinvent_version_biosphere,f"{biosphere3_act_CarbonDioxide_airurban['code']}"),
        amount = df_MaterialFlow.at[0, "FLUEOFF CO2"],
        unit = "kilogram",
        type = "biosphere",
    ).save()

    # Exchange 2: Monoethanolamine
    act_ProAbsDes.new_exchange(
        name = "Monoethanolamine",
        input = (ecoinvent_version_biosphere,f"{biosphere3_act_Monoethanolamine_air['code']}"),
        amount = df_MaterialFlow.at[0, "FLUEOFF MEA"],
        unit = "kilogram",
        type = "biosphere",
    ).save()

    # Exchange 3: Nitrogen
    act_ProAbsDes.new_exchange(
        name = "Nitrogen",
        input = (ecoinvent_version_biosphere,f"{biosphere3_act_Nitrogen_air['code']}"),
        amount = df_MaterialFlow.at[0, "FLUEOFF N2"],
        unit = "kilogram",
        type = "biosphere",
    ).save()

    # Exchange 4: Water
    act_ProAbsDes.new_exchange(
        name = "Water",
        input = (ecoinvent_version_biosphere,f"{biosphere3_act_Water_air['code']}"),
        amount = df_MaterialFlow.at[0, "FLUEOFF H2O"],
        unit = "cubicmeter",
        type = "biosphere",
    ).save()

    # Exchange 5: ProAbsDes
    act_ProAbsDes.new_exchange(
        name = "pro_absorption_desorption",
        input = act_ProAbsDes,
        amount = df_MaterialFlow.at[0, "CO2 captured"],
        unit = "kilogram",
        type = "production",
    ).save()

    # Exchange 6: MEA production
    act_ProAbsDes.new_exchange(
        name = "MEA production",
        input = act_MEAprod,
        amount = df_ProcessParameters.at[0, "Norm Solvent flow rate"],
        #amount=df_ProcessParameters.at[0, "Solvent flow rate"],
        unit = "kilogram",
        type = "technosphere",
    ).save()

    # Exchange 7: Reboiler - heat production, natural gas, at industrial furnace >100kW
    act_ProAbsDes.new_exchange(
        name = "heat production, natural gas, at industrial furnace >100kW",
        input = (ecoinvent_version_cutoff,f"{ecoinvent_act_heatprod['code']}"),
        amount = df_EnergyFlow.at[0, "Reboiler steam in MJ/ton CO2"] * (df_MaterialFlow.at[0,"CO2 captured"]/1000),
        unit = "megajoule",
        type = "technosphere",
    ).save()

    # Exchange 8: InfraAbsDes - Carbon capture plant
    act_ProAbsDes.new_exchange(
        name = "Infrastructure Absorber/ Desorber",
        input = act_InfraAbsDes,
        amount = df_Infrastructure.at[0, "Norm Carbon capture plant"],
        #amount = 1,
        unit = "unit",
        type = "technosphere",
    ).save()

    # Exchange 9: Market for monoethanolamine
    act_ProAbsDes.new_exchange(
        name = "Market for monoethanolamine",
        input =(ecoinvent_version_cutoff,f"{ecoinvent_act_marketMEA['code']}"),
        amount = df_MaterialFlow.at[0, "WATOUT out MEA"],
        unit = "kilogram",
        type = "technosphere",
    ).save()

    # Exchange 10: Market for tap water
    act_ProAbsDes.new_exchange(
        name = "Market for tap water",
        input =(ecoinvent_version_cutoff,f"{ecoinvent_act_markettapwater['code']}"),
        amount = df_NaturalResources.at[0, "Cooling water COOLER"] + df_NaturalResources.at[0, "Cooling water COND"],
        unit = "kilogram",
        type = "technosphere",
    ).save()

    # Exchange 11: Market for wastewater, average
    act_ProAbsDes.new_exchange(
        name = "Market for tap wastewater, average",
        input =(ecoinvent_version_cutoff,f"{ecoinvent_act_marketwastewater['code']}"),
        amount = df_MaterialFlow.at[0, "WATOUT out water"],
        unit = "cubicmeter",
        type = "technosphere",
    ).save()

    # Exchange 12: Market group for electricity, medium voltage
    act_ProAbsDes.new_exchange(
        name = "Market group for electricity, medium voltage",
        input =(ecoinvent_version_cutoff,f"{ecoinvent_act_marketelectr['code']}"),
        amount = df_EnergyFlow.at[0, "Total electr."],
        unit = "kilowatt hour",
        type = "technosphere",
    ).save()

    # Exchange 13: water production, deionised
    act_ProAbsDes.new_exchange(
        name = "Water production, deionised",
        input =(ecoinvent_version_cutoff,f"{ecoinvent_act_waterproddeion['code']}"),
        amount = df_MaterialFlow.at[0, "WASHWAT in"],
        unit = "kilogram",
        type = "technosphere",
    ).save()

    ######################################################################################################################
    # 3. CALCULATION OF LCA RESULTS
    ######################################################################################################################

    print("3. Calculating LCA results")
    #Calculating LCA score with activities and given method 
    lca_res=calculate_LCA(act_ProAbsDes, method_input)

    ###################################################################################################################
    # 4. POST-PROCESSING
    ###################################################################################################################

    print("4. Post-processing")

    # Dataframe for LCA results
    df_LCAresults = pd.DataFrame([[lca_res]])

    # Extract LCA results to excel file
    extract_res(df_LCAresults,i)

    print(f"> Simulation {i+1} finished.")
    print("-----------------------\n")

    return lca_res

if __name__=="__main__":

    # Exemplary data for testing brightway without Aspen Plus Simulation running each time
    data1=[
    pd.DataFrame(np.array([[19300]])).T,
    pd.DataFrame(np.array([
        [889.705289, 6093.5031, 936.35607, 4466.55226, 259.38827, 
         431.20635, 0.0, 0.0, 5379.03648, 45.19897, 4466.45466, 
         862.40675, 0.00819, 200.0, 16.55464, 0.14295]
    ])).T,
    pd.DataFrame(np.array([[82367.01, 45010.44686]])).T,
    pd.DataFrame(np.array([[3708.06766, 5.3988, 0.8998]])).T,
    pd.DataFrame(np.array([[50000, 2327.55259, 38000, 1745.66444]])).T
]
    data2=[ 
        pd.DataFrame(np.array([[19300], [19200]])).T, 
        pd.DataFrame(np.array([ [889.705289, 6093.5031, 936.35607, 4466.55226, 259.38827, 431.20635, 0.0, 0.0, 5379.03648, 45.19897, 4466.45466, 862.40675, 0.00819, 200.0, 16.55464, 0.14295], 
                               [800.0, 6000.0, 900.0, 4400.0, 260.0, 430.0, 0.0, 0.0, 5300.0, 45.0, 4400.0, 860.0, 0.01, 210.0, 17.0, 0.15] ])).T, 
        pd.DataFrame(np.array([[82367.01, 45010.44686], [82000, 44000]])).T, 
        pd.DataFrame(np.array([[3708.06766, 5.3988, 0.8998], [3600, 5.5, 1.0]])).T, 
        pd.DataFrame(np.array([[50000, 2327.55259, 38000, 1745.66444], [49000, 2300, 37000, 1700]])).T ]
    run_LCA(data2, 1, True)
    # print(data1)
    # print(data2[1].iloc[0, 1])

    