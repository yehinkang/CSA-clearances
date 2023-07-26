#importing the following packages
import os, math, json, copy, re
from datetime import datetime
import pandas as pd
import numpy as np
import sys
from commons import report_xlsx_general
import time
#The following can be toggled in order to see an unabridged numpy array
np.set_printoptions(threshold=sys.maxsize)

"""
This code will function by taking inputs in and using them to reference a spreadsheet.
From this spreadsheet results can then be computed and exported to another table.
This was done as tables can be easily updated by humans should there be changes to a standard in the future.
I'm unfamiliar with flask and as such there is nothing in the way of web integration thus far.
"""

#This is where the reference file is located
ref_table_filepath = r"C:\Users\yekang\OneDrive - POWER Engineers, Inc\Desktop\Python Stuff\Clearance Calcs\AEUC_CSA_clearances.xlsx"

#Indicated the location of the excel file and loads it into python, this ensures the file is only read once
ref_table= pd.ExcelFile(ref_table_filepath)

#Let's set up a universal input dictionary
inputs = {
    'p2p_voltage': 138,
    'Max_Overvoltage': 0.05,

    'Buffer_Neut':1.2,
    'Buffer_Live':1.2,

    'Design_buffer_2_obstacles':1.2,
    'Design_Buffer_Same_Structure':0.6,

    'Clearance_Rounding':0.01,
    
    #There will need to be a dropdown menu made on the webpage with matching locations made on the webpage
    'Location': "Alberta Camrose",

    #Add these as check boxes in the webpage
    'grounded': False,
    'sheathed': False,

    #Latitude_Northing
        'Northing_deg':53.02,
        'Northing_mins':0,
        'Northing_seconds':0,

    #Longitude Westing
        'Westing_deg':-112.83,
        'Westing_mins':0,
        'Westing_seconds':0,

    # Loc_Elevation (Calculated fromt he coordinates)

    'Custom_Elevation':1200,

    #Table 17 Horizontal Conductor Seperations
        'Span_Length':150,
        'Final_Unloaded_Sag_15C':1,

    #Crossing_or_Underbuild
        'Is_main_wire_upper':False,
        'lower wire':True,
        'XING_P2P_Voltage':79,
        'XING_Max_Overvoltage':0.05,

}

#The following function takes the decimal points in the dictionary and turns that into an integer to use in np.round in all functions
def decimal_points(num):
    num_str = str(num)
    if '.' in num_str:
        return len(num_str.split('.')[1])
    else:
        return 0
Clearance_Rounding = {key: inputs[key] for key in ['Clearance_Rounding']}
Nums_after_decimal = Clearance_Rounding['Clearance_Rounding']
Numpy_round_integer = decimal_points(Nums_after_decimal)

#The following function will take an dictionary and export it into excel using pandas(useful for troubleshooting dictionary outputs)
def save2xl(x):
    #Insert desired filepath here, can rename later
    Excel_export = pd.DataFrame(x)
    folder = "C:/Users/yekang/OneDrive - POWER Engineers, Inc/Desktop/Python Stuff/Clearance Calcs/Samples/"
    start = datetime.now().strftime("Clearance_Calc %Y-%d-%m-%H-%M-%S")
    fname = folder + start + ".xlsx"
    print("saving to", fname)
    Excel_export.to_excel(fname)

#The following function will take in Longitude and Latitude and return the closest location on which there is information
#Currently Using the old tables but will have to be updated
def location_lookup(Northing_deg, Northing_mins, Northing_seconds, Westing_deg, Westing_mins, Westing_seconds):
    #Taking the inputs and turning them into decimal Latitude and Longitude
    Latitude = Northing_deg + (Northing_mins/60) + (Northing_seconds/3600)
    Longitude = Westing_deg + (Westing_mins/60) + (Westing_seconds/3600)

    #Stitching together Latitude and longitude inputs in decimal form so that they match entries in the reference array
    Input_Coordinates= np.array([Latitude, Longitude])

    #Opening the sheets within the file
    Lookup = pd.read_excel(ref_table, "CSA C22.3 No.1 Table D2")

    #Importing the columns from the excel sheet
    Latitude_lookup = Lookup['Latitude']
    Longitude_lookup = Lookup['Longitude']
    Place_name = Lookup['For Lookup']
    Elevation = Lookup['Elevation, m']
    Snow_Depth = Lookup['Mean annual maximum snow depth, m']

    #Combining the Latitude and Longitude columns into a 2D data frame
    Combined_Coord = pd.concat([Latitude_lookup,Longitude_lookup], axis = 1)

    #Transforming the dataframe into a numpy array
    Combined_Coord_array = Combined_Coord.to_numpy()

    #Finding the closest match to the inputted coordinates in the array
    Closest = Combined_Coord_array[np.linalg.norm(Combined_Coord_array-Input_Coordinates, axis=1).argmin()]

    #Find the position in the array where the coordinate is
    position = np.where((Combined_Coord_array == Closest).all(axis=1))[0][0]

    #Finding the value in the same position in the other datasets
    Closest_place_name = Place_name.iloc[position]
    Max_snow_depth = Snow_Depth.iloc[position]

    Custom_Elevation = {key: inputs[key] for key in ['Custom_Elevation']}
    if Custom_Elevation['Custom_Elevation'] == 0:
        Closest_elevation = Elevation.iloc[position]
    else:
        Closest_elevation = Custom_Elevation['Custom_Elevation']

    #The following will return the values as calculated by the function to the user.
    return Closest_place_name, Closest_elevation, Max_snow_depth

#The following runs the location_lookup function with the dictionary inputs for use in code going forwards
loc_lookup_input = {key: inputs[key] for key in ['Northing_deg', 'Northing_mins', 'Northing_seconds', 'Westing_deg', 'Westing_mins', 'Westing_seconds']}
Place, Altitude, Snow_Depth = location_lookup(**loc_lookup_input)

#The following function will take inputs and create output arrays that can be inserted into a table
#This will take voltage in kilovolts and everthing else in meters
def AEUC_table5_clearance(p2p_voltage, Buffer_Neut, Buffer_Live):

    voltage = p2p_voltage / np.sqrt(3)
    #Start by opening the sheet for AEUC Table 5 within the reference table file
    AEUC_5 = pd.read_excel(ref_table, "AEUC Table 5")

    #this column is for guy wires, comms.etc
    AEUC_neut_clearance = AEUC_5['Neutral']

    #This next little bit will determine which column should be referenced
    if  0 < voltage <= 0.75: 
        #this column is for 120-660V
        AEUC_clearance = AEUC_5['0 to 0.75']
        voltage_range = '120-600V'
        column = "II"
    elif voltage <=25: 
        #this column is for 4, 13 & 25 kV
        AEUC_clearance = AEUC_5['0.75 to 22']
        voltage_range = '4, 13 & 25 kV'
        column = "III"
    elif voltage <=50: 
        #this column is for 35, 69 & 72 kV
        AEUC_clearance = AEUC_5['22 to 50']
        voltage_range = '35, 69 & 72 kV'
        column = "IV"
    elif voltage <=90: 
        #this column is for 138 & 144 kV
        AEUC_clearance = AEUC_5['50 to 90']
        voltage_range = '138 & 144 kV'
        column = "V"
    elif voltage <=150: 
        #this column is for 240 kV
        AEUC_clearance = AEUC_5['120 to 150']
        voltage_range = '240 kV'
        column = "VI"
    elif voltage <=318: 
        #this column is for 500 kV
        AEUC_clearance = AEUC_5['318']
        voltage_range = '500 kV'
        column = "VII"
    else: 
        #this column is for 1000 kV Pole-Pole
        AEUC_clearance = AEUC_5['+/- 500 kVDC']
        voltage_range = '1000 kVDC Pole-to-Pole'
        column = "VIII"

    #Adders will have to be multiplied by a series so that it only applies to the correct rows
    """ 
    The following list shows what each position in the Array correspons to when going L -> R
    1. Over walkways or land normally accessible only to pedestrians, snowmobiles, and all terrain vehicles not exceeding 3.6m
    2. Over rights of way of underground pipelines operating at a pressure of over 700 kilopascals; equipment not exceeding 4.15m 
    3. Over land likely to be travelled by road vehicles (including roadways, streets, lanes, alleys, driveways, and entrances); equipment not exceeding 4.15m"
    4. Over land likely to be travelled by road vehicles (including highways, roadways, streets, lanes, alleys, driveways, and entrances); equipment not exceeding 5.3m
    5. Over land likely to be travelled by agricultural or other equipment; equipment not exceeding 5.3m
    6. Above top of rails at railway crossings, equipment not exceeding 7.2m 
    """
    Repave_Adder = np.array([0,0,0.225,0.225,0,0.3], dtype=float)
    Snow_Adder = Snow_Depth * np.array([1,1,0,0,1,0], dtype=float)

    #Need to find altitude
    if Altitude > 1000:
        Altitude_Adder = (Altitude - 1000) / 100 * 0.01 * AEUC_clearance
    else:
        Altitude_Adder = np.array([0,0,0,0,0,0], dtype=float)

    #The buffers are now being added in so that they add to every entry in the table
    Buffer_live_array = Buffer_Live * np.ones(len(AEUC_clearance))
    Buffer_neut_array = Buffer_Neut * np.ones(len(AEUC_clearance))

    #Now the following columns are calculating base AEUC clearances and clerances with buffers
    AEUC_total_clearance_neut = AEUC_neut_clearance + Repave_Adder + Snow_Adder
    Design_clearance_neut = AEUC_total_clearance_neut + Buffer_neut_array

    AEUC_total_clearance = AEUC_clearance + Repave_Adder + Snow_Adder + Altitude_Adder
    Design_clearance = AEUC_total_clearance + Buffer_live_array

    #The following is to round all the values before they are sent to the table
    AEUC_neut_clearance = np.round(AEUC_neut_clearance, Numpy_round_integer)
    Repave_Adder = np.round(Repave_Adder, Numpy_round_integer)
    Snow_Adder = np.round(Snow_Adder, Numpy_round_integer)
    AEUC_total_clearance_neut = np.round(AEUC_total_clearance_neut, Numpy_round_integer)
    Design_clearance_neut = np.round(Design_clearance_neut, Numpy_round_integer)
    AEUC_clearance = np.round(AEUC_clearance, Numpy_round_integer)
    Altitude_Adder = np.round(Altitude_Adder, Numpy_round_integer)
    AEUC_total_clearance = np.round(AEUC_total_clearance, Numpy_round_integer)
    Design_clearance = np.round(Design_clearance, Numpy_round_integer)

    #Creating a dictionary to turn into a dataframe
    data = {
        'AEUC base neutral': AEUC_neut_clearance,
        'Re-pave Adder': Repave_Adder,
        'Snow Adder': Snow_Adder,
        'AEUC total neutral': AEUC_total_clearance_neut,
        'Design clearance neutral': Design_clearance_neut,

        'AEUC base clearance': AEUC_clearance,
        'Altitude Adder': Altitude_Adder,
        #Using the same repave and snow adders as above for conductors
        'AEUC total': AEUC_total_clearance,
        'Design clearance': Design_clearance,

        'Voltage range': voltage_range,
        'Column #': column,
        }
    
    #The following is a pandas dataframe and it is using the function from above to export the dataframe into an excel file
    #save2xl(data)
    return data
#Running the AEUC table 5 function to get a dictionary that can be used in a webapp and used in the export to excel code later on
AEUC5_input = {key: inputs[key] for key in ['p2p_voltage', 'Buffer_Neut', 'Buffer_Live']}
AEUC_Table5_data  = AEUC_table5_clearance(**AEUC5_input)

#The following function is for AEUC table 7
def AEUC_table7_clearance(p2p_voltage, grounded, sheathed, Design_buffer_2_obstacles):

    voltage = p2p_voltage / np.sqrt(3)
    #Start by opening the sheet for AEUC Table 5 within the reference table file
    AEUC_7 = pd.read_excel(ref_table, "AEUC Table 7")
    #Since .loc is being used to find the row a column must be chosen to use 
    AEUC_7.set_index("Wire or Conductor", inplace = True)
    
    #Specify the row that is being used for this
    AEUC_neut_clearance_basic = AEUC_7.loc['Guys communications cables, and drop wires'].values
    AEUC_neut_clearance = Design_buffer_2_obstacles * np.ones(len(AEUC_neut_clearance_basic))

    #Figuring out the clearance to use with the conductor voltage
    if 0 < voltage <= 0.75:
        if sheathed == True:
            #Specify the row that is being used for this
            AEUC_clearance = AEUC_7.loc['0 - 750 V Enclosed in effectively grounded metallic sheath'].values
            AEUC_voltage_adder = 0
        if grounded == True:
            #Specify the row that is being used for this
            AEUC_clearance = AEUC_7.loc['0 - 750 V Insulated or grounded'].values
            AEUC_voltage_adder = 0
        else:
            #Specify the row that is being used for this
            AEUC_clearance = AEUC_7.loc['0 - 750 V Neither insulated or grounded or enclosed in effectively grounded metallic sheath'].values
            AEUC_voltage_adder = 0
    elif 0.75 < voltage <= 22:
        if sheathed == False:
            #Specify the row that is being used for this
            AEUC_clearance = AEUC_7.loc['0.75 - 22kV Not enclosed in  effectively grounded metallic sheath'].values
            AEUC_voltage_adder = 0
        else:
            #Specify the row that is being used for this
            AEUC_clearance = AEUC_7.loc['0.75 - 22kV Enclosed in effectively grounded metallic sheath'].values
            AEUC_voltage_adder = 0
    else:
        #Specify the row that is being used for this
            AEUC_clearance = AEUC_7.loc['Over 22kV'].values
            AEUC_voltage_adder = (voltage - 22) * 0.01

    #The following is to round all the values before they are sent to the table
    AEUC_neut_clearance_basic = np.round(AEUC_neut_clearance_basic, Numpy_round_integer)
    AEUC_clearance = np.round(AEUC_clearance, Numpy_round_integer)
    AEUC_voltage_adder = np.round(AEUC_voltage_adder, Numpy_round_integer)
    Design_buffer_2_obstacles = np.round(Design_buffer_2_obstacles, Numpy_round_integer)

    #Indexing certain elements in the array BA = basic, VA = Voltage Adder, AT = AEUC Total, DC = Design Clearance
    H2building_BA = np.array([AEUC_neut_clearance_basic[0], AEUC_clearance[0]])
    H2building_VA = np.array([0, AEUC_voltage_adder])
    H2building_AT = np.array([AEUC_neut_clearance_basic[0], (AEUC_clearance[0] + AEUC_voltage_adder)])
    H2building_DC = np.array([AEUC_neut_clearance[0], (AEUC_clearance[0] + AEUC_voltage_adder + Design_buffer_2_obstacles)])

    V2building_BA = np.array([AEUC_neut_clearance_basic[1], AEUC_clearance[1]])
    V2building_VA = H2building_VA
    V2building_AT = np.array([AEUC_neut_clearance_basic[1], (AEUC_clearance[1] + AEUC_voltage_adder)])
    V2building_DC = np.array([AEUC_neut_clearance[1], (AEUC_clearance[1] + AEUC_voltage_adder + Design_buffer_2_obstacles)])


    H2object_BA = np.array([AEUC_neut_clearance_basic[2], AEUC_clearance[2]])
    H2object_VA = H2building_VA
    H2object_AT = np.array([AEUC_neut_clearance_basic[2], (AEUC_clearance[2] + AEUC_voltage_adder)])
    H2object_DC = np.array([AEUC_neut_clearance[2], (AEUC_clearance[2] + AEUC_voltage_adder + Design_buffer_2_obstacles)])

    V2object_BA = np.array([AEUC_neut_clearance_basic[3], AEUC_clearance[3]])
    V2object_VA = H2building_VA
    V2object_AT = np.array([AEUC_neut_clearance_basic[3], (AEUC_clearance[3] + AEUC_voltage_adder)])
    V2object_DC = np.array([AEUC_neut_clearance[3], (AEUC_clearance[3] + AEUC_voltage_adder + Design_buffer_2_obstacles)])

    Categories = np.array(["Guys, communication cables, and drop wires", "Supply conductors"])

    #Creating a dictionary to turn into a dataframe
    data = {
        'Wire or Conductor': Categories,
        'Building Horizontal to Surface Basic': H2building_BA,
        'Building Horizontal to Surface Voltage Adder': H2building_VA,
        'Building Horizontal to Surface AEUC Total': H2building_AT,
        'Building Horizontal to Surface Design Clearance': H2building_DC,

        'Building Vertical to Surface Basic': V2building_BA,
        'Building Vertical to Surface Voltage Adder': V2building_VA,
        'Building Vertical to Surface AEUC Total': V2building_AT,
        'Building Vertical to Surface Clearance': V2building_DC,

        'Obstacle Horizontal to Surface Basic': H2object_BA,
        'Obstacle Horizontal to Surface Voltage Adder': H2object_VA,
        'Obstacle Horizontal to Surface AEUC Total': H2object_AT,
        'Obstacle Horizontal to Surface Clearance': H2object_DC,

        'Obstacle Vertical to Surface Basic': V2object_BA,
        'Obstacle Vertical to Surface Voltage Adder': V2object_VA,
        'Obstacle Vertical to Surface AEUC Total': V2object_AT,
        'Obstacle Vertical to Surface Clearance': V2object_DC,
        }
    
    #The following is a pandas dataframe and it is using the function from above to export the dataframe into an excel file
    #save2xl(data)
    return data
AEUC7_input = {key: inputs[key] for key in ['p2p_voltage', "grounded", "sheathed", 'Design_buffer_2_obstacles']}
AEUC_Table7_data = AEUC_table7_clearance(**AEUC7_input)

#The following function is for CSA Table 2
def CSA_Table_2_clearance(p2p_voltage, Buffer_Neut, Buffer_Live):

    voltage = p2p_voltage / np.sqrt(3)
    #Start by opening the sheet for CSA Table 2 within the reference table file
    CSA_2 = pd.read_excel(ref_table, "CSA table 2")

    #Specify the row that is being used for this
    CSA_2_neut_clearance = CSA_2['Guys, messengers, communication, span & lightning protection wires; communication cables']

    if  0 < voltage <= 0.75: 
        #this column is for 120-660V
        CSA_2_clearance = CSA_2['0-0.75']
        voltage_range = '0-0.75'
        column = 'III'
    elif voltage <=22: 
        #this column is for 4, 13 & 25 kV
        CSA_2_clearance = CSA_2['0.75-22']
        voltage_range = '> 0.75 ≤ 22'
        column = 'IV'
    elif voltage <=50: 
        #this column is for 35, 69 & 72 kV
        CSA_2_clearance = CSA_2['22-50']
        voltage_range = '> 22 ≤ 50'
        column = 'V'
    elif voltage <=90: 
        #this column is for 138 & 144 kV
        CSA_2_clearance = CSA_2['50-90']
        voltage_range = '> 50 ≤ 90'
        column = 'VI'
    elif voltage <=120: 
        #this column is for 240 kV
        CSA_2_clearance = CSA_2['90-120']
        voltage_range = '> 90 ≤ 120'
        column = 'VII'
    elif voltage <=150: 
        #this column is for 500 kV
        CSA_2_clearance = CSA_2['120-150']
        voltage_range = '> 120 ≤ 150'
        column = 'VIII'
    elif voltage <=190: 
        #this column is for 500 kV
        CSA_2_clearance = CSA_2['150-190']
        voltage_range = '> 150 ≤ 190'
        column = 'IX'
    elif voltage == 360: 
        #this column is for 360 kV
        CSA_2_clearance = CSA_2['219']
        voltage_range = '219 (360)'
        column = 'X'
    elif voltage == 500: 
        #this column is for 500 kV
        CSA_2_clearance = CSA_2['318']
        voltage_range = '318 (500)'
        column = 'XI'
    else: 
        #this column is for 735 kV
        CSA_2_clearance = CSA_2['442']
        voltage_range = '442 (735)'
        column = 'XII'

    Repave_Adder = np.array([0.225,0,0,0.225,0,0.3], dtype=float)

    #Need to find mean annual snow depth
    Snow_Adder = Snow_Depth * np.array([0,1,1,1,1,0], dtype=float)

    #Need to find altitude
    if Altitude > 1000:
        Altitude_Adder = (Altitude - 1000)/100 * 0.01 * CSA_2_clearance
    else:
        Altitude_Adder = np.array([0,0,0,0,0,0], dtype=float)

    Neutral_CSA_Total = Snow_Adder + Repave_Adder + CSA_2_neut_clearance
    CSA_Total = Snow_Adder + Repave_Adder + Altitude_Adder + CSA_2_clearance

    Neutral_DC = Buffer_Neut + Neutral_CSA_Total
    CSA_DC = Buffer_Live + CSA_Total

    CSA_2_neut_clearance = np.round(CSA_2_neut_clearance, Numpy_round_integer)
    Repave_Adder = np.round(Repave_Adder, Numpy_round_integer)
    Snow_Adder = np.round(Snow_Adder, Numpy_round_integer)
    Neutral_CSA_Total = np.round(Neutral_CSA_Total, Numpy_round_integer)
    Neutral_DC = np.round(Neutral_DC, Numpy_round_integer)
    CSA_2_clearance = np.round(CSA_2_clearance, Numpy_round_integer)
    Altitude_Adder = np.round(Altitude_Adder, Numpy_round_integer)
    CSA_Total = np.round(CSA_Total, Numpy_round_integer)
    CSA_DC = np.round(CSA_DC, Numpy_round_integer)

    data = {
        'Neutral Basic (m)': CSA_2_neut_clearance,
        'Re-pave Adder (m)': Repave_Adder,
        'Snow Adder (m)': Snow_Adder,
        'Neutral CSA Total (m)': Neutral_CSA_Total,
        'Neutral Design Clearance (m)': Neutral_DC,

        'Basic (m)': CSA_2_clearance,
        'Altitude Adder (m)': Altitude_Adder,
        #Same Re-pave and Snow adder used as in above
        'CSA Total (m)': CSA_Total,
        'Design Clearance (m)': CSA_DC,

        #This will only be used for the excel doc
        'Voltage range': voltage_range,
        'Column #': column,
        }
    
    return data
CSA2_input = {key: inputs[key] for key in ['p2p_voltage', "Buffer_Neut", "Buffer_Live"]}
CSA_Table2_data = CSA_Table_2_clearance(**CSA2_input)

#The following function is for CSA Table 3
def CSA_Table_3_clearance(p2p_voltage, Buffer_Neut, Buffer_Live):

    voltage = p2p_voltage / np.sqrt(3)
    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_3 = pd.read_excel(ref_table, "CSA table 3")

    #Neutral clearance row being found on spreadsheet
    CSA_3_neut_clearance = CSA_3['Guys; messengers; communication, span, and lightning protection wires; communication cables']

    #SFinding the correct clearance value based on conductor voltage
    if  0 < voltage <= 22: 
        CSA_3_clearance = CSA_3['0-22']
        voltage_range = '0-22'
    elif voltage <=50: 
        CSA_3_clearance = CSA_3['>22<50']
        voltage_range = '> 22 ≤ 50'
    elif voltage <=90: 
        CSA_3_clearance = CSA_3['>50<90']
        voltage_range = '> 50 ≤ 90'
    elif voltage <=150: 
        CSA_3_clearance = CSA_3['>90<150']
        voltage_range = '> 90 ≤ 150'
    else: 
        CSA_3_clearance = CSA_3['>150 + 0.01m/kV over 150kV'] + np.ones(8)*0.1*(voltage-150)
        voltage_range = '> 150'

    #Need to find altitude
    if Altitude > 1000:
        arr = (Altitude - 1000)
        divisor = 100
        rounded_arr = np.floor(arr / divisor) * divisor
        Altitude_Adder = rounded_arr/100 * 0.01 * CSA_3_clearance 
    else:
        Altitude_Adder = np.array[(0,0,0,0,0,0)]

    #Calculating the total clearance with adders
    CSA_Total = Altitude_Adder + CSA_3_clearance
    #DC = Design Clearance
    Neutral_DC = Buffer_Neut + CSA_3_neut_clearance
    CSA_DC = Buffer_Live + CSA_Total

    #Rounding the
    CSA_3_neut_clearance = np.round(CSA_3_neut_clearance, Numpy_round_integer)
    Neutral_DC = np.round(Neutral_DC, Numpy_round_integer)
    CSA_3_clearance = np.round(CSA_3_clearance, Numpy_round_integer)
    Altitude_Adder = np.round(Altitude_Adder, Numpy_round_integer)
    CSA_Total = np.round(CSA_Total, Numpy_round_integer)
    CSA_DC = np.round(CSA_DC, Numpy_round_integer)

    data = {
        'Neutral Basic (m)': CSA_3_neut_clearance,
        'Neutral Design Clearance (m)': Neutral_DC,

        'Basic (m)': CSA_3_clearance,
        'Altitude Adder (m)': Altitude_Adder,
        'CSA Total (m)': CSA_Total,
        'Design Clearance (m)': CSA_DC,

        'Voltage range': voltage_range,
        }

    return data
CSA3_input = {key: inputs[key] for key in ['p2p_voltage', "Buffer_Neut", "Buffer_Live"]}
CSA_Table3_data = CSA_Table_3_clearance(**CSA3_input)

#The following function is for CSA Table 5
def CSA_Table_5_clearance(p2p_voltage, Buffer_Live):

    voltage = p2p_voltage / np.sqrt(3)
    #Start by opening the sheet for CSA Table 5 within the reference table file
    CSA_5 = pd.read_excel(ref_table, "CSA table 5")

    #Finding correct clearance value from the voltage
    if  0 < voltage <= 0.75: 
        CSA_5_clearance = CSA_5['0–750V']
        Voltage_range = '0–750V'
    elif voltage <=22: 
        CSA_5_clearance = CSA_5['>0.75kV<22kV']
        Voltage_range = '> 0.75 ≤ 22'
    elif voltage <=50: 
        CSA_5_clearance = CSA_5['>22kV<50kV']
        Voltage_range = '> 22 ≤ 50'
    else: 
        CSA_5_clearance = CSA_5['>50kV<90kV']
        Voltage_range = '> 50 ≤ 90'

    #DC = Design Clearance
    DC = CSA_5_clearance + Buffer_Live

    #Rounding the clearance values
    CSA_5_clearance = np.round(CSA_5_clearance, Numpy_round_integer)
    DC = np.round(DC, Numpy_round_integer)

    #Creating a dictionary to store all the outpt data
    data = {
        'Basic (m)': CSA_5_clearance,
        'Design Clearance (m)': DC,

        'Voltage range': Voltage_range
        }

    return data
CSA5_input = {key: inputs[key] for key in ['p2p_voltage', "Buffer_Live"]}
CSA_Table5_data = CSA_Table_5_clearance(**CSA5_input)

#The following function is for CSA Table 6
def CSA_Table_6_clearance(p2p_voltage, Buffer_Neut, Buffer_Live):

    voltage = p2p_voltage / np.sqrt(3)
    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_6 = pd.read_excel(ref_table, "CSA table 6")
    #Since .loc is being used to find the row a column must be chosen to use 
    CSA_6.set_index("Wire closest to tracks", inplace = True)

    CSA_6_neut_clearance = CSA_6.loc['Guys; messengers; communication, span, and lightning protection wires; communication cables']
    CSA_6_sub750_clearance = CSA_6.loc['Open supply-line conductors and service conductors of 0—750 V and effectively grounded continuous metallic sheathed cables of all voltages']

    if  0 < voltage <= 0.75: 
        CSA_6_clearance = CSA_6.loc ['AC > 0.75 < 22 kV']
        Voltage_range = '> 0.75 ≤ 22 kV'
    elif voltage <=22: 
        CSA_6_clearance = CSA_6.loc ['AC > 22 < 50 kV']
        Voltage_range = '> 22 ≤ 50 kV'
    elif voltage <=50: 
        CSA_6_clearance = CSA_6.loc ['AC > 50 < 90 kV']
        Voltage_range = '> 50 ≤ 90 kV'
    elif voltage <=50: 
        CSA_6_clearance = CSA_6.loc ['AC > 90 < 120 kV']
        Voltage_range = '> 90 ≤ 120 kV'
    elif voltage <=50: 
        CSA_6_clearance = CSA_6.loc ['AC > 120 < 150 kV']
        Voltage_range = '> 120 ≤ 150 kV'
    else: 
        CSA_6_clearance = CSA_6.loc ['AC Supply conductors > 150 kV +0.01m/kV over 150kV']
        Voltage_range = '> 150 kV'

    #DC = Design Clearance
    DC_neut = CSA_6_neut_clearance + Buffer_Neut
    DC_sub750 = CSA_6_sub750_clearance + Buffer_Live
    DC = CSA_6_clearance + Buffer_Live

    CSA_6_clearance = np.round(CSA_6_clearance, Numpy_round_integer)
    CSA_6_neut_clearance = np.round(CSA_6_neut_clearance, Numpy_round_integer)
    CSA_6_sub750_clearance = np.round(CSA_6_sub750_clearance, Numpy_round_integer)
    DC = np.round(DC, Numpy_round_integer)
    DC_neut = np.round(DC_neut, Numpy_round_integer)
    DC_sub750 = np.round(DC_sub750, Numpy_round_integer)

    Basic_Main_Track = np.array([CSA_6_neut_clearance[0], CSA_6_sub750_clearance[0], CSA_6_clearance[0]])
    DC_Main_Track = np.array([DC_neut[0], DC_sub750[0], DC[0]])
    Basic_Siding = np.array([CSA_6_neut_clearance[1], CSA_6_sub750_clearance[1], CSA_6_clearance[1]])
    DC_Siding = np.array([DC_neut[1], DC_sub750[1], DC[1]])

    data = {
        'Main Basic (m)': Basic_Main_Track,
        'Main Design Clearance (m)': DC_Main_Track,
        'Siding Basic (m)': Basic_Siding,
        'Siding Design Clearance (m)': DC_Siding,

        'Voltage range': Voltage_range
        }

    return data
CSA6_input = {key: inputs[key] for key in ['p2p_voltage', "Buffer_Neut", "Buffer_Live"]}
CSA_Table6_data = CSA_Table_6_clearance(**CSA6_input)

#The following function is for CSA Table 7
def CSA_Table_7_clearance(Buffer_Live):

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_7 = pd.read_excel(ref_table, "CSA table 7")

    CSA_7clearance = CSA_7['Minimum horizontal separation, m']

    #Creating an array to add a design buffer to the basic clearance
    Buffer_Live_array = np.ones(len(CSA_7)) * Buffer_Live
    CSA6_DC = CSA_7clearance + Buffer_Live_array

    #Rounding
    CSA_7clearance = np.round(CSA_7clearance, Numpy_round_integer)
    CSA6_DC = np.round(CSA6_DC, Numpy_round_integer)

    data = {
        'Basic (m)': CSA_7clearance,
        'Design Clearance (m)': CSA6_DC,
        }
    return data
CSA7_input = {key: inputs[key] for key in ["Buffer_Live"]}
CSA_Table7_data = CSA_Table_7_clearance(**CSA7_input)

#The following function will create worksheets from the data calculated by the functions above
def create_report_excel():

    #The next line defines the output filepath
    folder = "C:/Users/yekang/OneDrive - POWER Engineers, Inc/Desktop/Python Stuff/Clearance Calcs/Samples/"
    #This next line will create a string from the current time down to the second
    start = datetime.now().strftime("Clearance_Calc %Y-%d-%m-%H-%M-%S")
    #This next line will create a filename that uses the previous line to create a unique file name
    filename = folder + start + ".xlsx"
    print("saving to", filename)

    # define some colors for background fill of the worksheets
    color_bkg_header = 'E0E0E0'
    color_bkg_data_1 = 'CCE5FF'
    color_bkg_data_2 = 'CCFFFF'

#AEUC Table 5
#region
    #Creating the title blocks. etc
    AEUC5_cell00 = 'Over walkways or land normally accessible only to pedestrians, snowmobiles, and all terrain vehicles not exceeding 3.6m'
    AEUC5_cell10 = 'Over rights of way of underground pipelines operating at a pressure of over 700 kilopascals; equipment not exceeding 4.15m'
    AEUC5_cell20 = 'Over land likely to be travelled by road vehicles (including roadways, streets, lanes, alleys, driveways, and entrances); equipment not exceeding 4.15m'
    AEUC5_cell30 = 'Over land likely to be travelled by road vehicles (including highways, roadways, streets, lanes, alleys, driveways, and entrances); equipment not exceeding 5.3m'
    AEUC5_cell40 = 'Over land likely to be travelled by agricultural or other equipment; equipment not exceeding 5.3m'
    AEUC5_cell50 = 'Above top of rails at railway crossings, equipment not exceeding 7.2m'
    voltage = inputs['p2p_voltage']
    elevation = Altitude
    voltage_range = AEUC_Table5_data['Voltage range']
    col = 'Col. ' + AEUC_Table5_data['Column #']

    AEUC5_titles = np.array([AEUC5_cell00, AEUC5_cell10, AEUC5_cell20, AEUC5_cell30, AEUC5_cell40, AEUC5_cell50])
    
    AEUC_5 = [
    #The following is the title header before there is data
        ['AEUC table 5 \n Minimum Vertical Design Clearances above Ground or Rails \n (See Rule 10-002 (5) and (6), Appendix C and CSA C22-3 No. 1-15 Clause 5.3.1.1.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase), Site Elevation: '+ str(elevation) +' m'],
        [' '],
        [' '],
        [' ', 'Guys, Messengers, Span & Lightening Protection Wires and Communications Wires and Cables', ' ', ' ', ' ', ' ', 'Voltage of Open Supply Conductors and Service Conductors Voltage Line to Ground kV AC except where note (Voltages in Parentheses are AC Phase to Phase) ' + str(voltage_range)],
        [' '],
        [' ', 'Col. I', ' ', ' ', ' ', ' ' , col],
        [' ', 'Basic (m)', 'Re-pave Adder (m)', "Snow Adder (m)", "AEUC Total (m)", "Design Clearance (m)", 'Basic (m)', "Altitude Adder (m)", 'Re-pave Adder (m)', "Snow Adder (m)", "AEUC Total (m)", "Design Clearance (m)"],
              ]
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(6):
        row = [AEUC5_titles[i], AEUC_Table5_data['AEUC base neutral'][i], AEUC_Table5_data['Re-pave Adder'][i], AEUC_Table5_data['Snow Adder'][i], AEUC_Table5_data['AEUC total neutral'][i], AEUC_Table5_data['Design clearance neutral'][i], \
               AEUC_Table5_data['AEUC base clearance'][i], AEUC_Table5_data['Altitude Adder'][i], AEUC_Table5_data['Re-pave Adder'][i], AEUC_Table5_data['Snow Adder'][i], AEUC_Table5_data['AEUC total'][i], AEUC_Table5_data['Design clearance'][i]]
        AEUC_5.append(row)

    #This is retrieving the number of rows in each table array
    n_row_AEUC_table_5 = len(AEUC_5)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_AEUC_table_5 = len(AEUC_5[7])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_AEUC_5 = []
    for i in range(n_row_AEUC_table_5):
        if i < 3:
            list_range_color_AEUC_5.append((i + 1, 1, i + 1, n_column_AEUC_table_5, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_AEUC_5.append((i + 1, 1, i + 1, n_column_AEUC_table_5, color_bkg_data_1))
        else:
            list_range_color_AEUC_5.append((i + 1, 1, i + 1, n_column_AEUC_table_5, color_bkg_data_2))

    # define cell format
    cell_format_AEUC_5 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_AEUC_table_5, 'center'), (4, 2, 5, 6, 'center'), (4, 1, 6, 1, 'center'), (4, 7, 5, 12, 'center'), (6, 2, 6, 6, 'center'), (6, 7, 6, 12, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_AEUC_5,
        'range_border': [(1, 1, 1, n_column_AEUC_table_5), (2, 1, n_row_AEUC_table_5, n_column_AEUC_table_5)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 20) for i in range(n_column_AEUC_table_5)],
    }

    # define some footer notes
    footer_AEUC5 = ['Note: See high load corridors on map of Provincial Highways for vehicle hights of 9.0m and 12.8m.', 'AESO rules section 502.2 clause 17 (3) requires a minimum clearance of 12.2 m over agricultural land.', 'Basic clearances from AUEC code 2022 table 5.']

    # define the worksheet
    AEUC_Table_5 = {
        'ws_name': 'AEUC Table 5',
        'ws_content': AEUC_5,
        'cell_range_style': cell_format_AEUC_5,
        'footer': footer_AEUC5
    }
#endregion

#AEUC Table 7
#region
    AEUC7_cell00 = 'Guys, communication cables, and drop wires'
    AEUC7_cell10 = 'Supply conductors'
    voltage = inputs['p2p_voltage']
    elevation = Altitude
    voltage_range = AEUC_Table5_data['Voltage range']
    col = 'Col. ' + AEUC_Table5_data['Column #']

    AEUC7_titles = np.array([AEUC7_cell00, AEUC7_cell10])
    
    AEUC_7 = [
    #The following is the title header before there is data
        ['AEUC table 7 \n Minimum Design Clearances from Wires and Conductors Not Attached to Buildings, Signs, and Similar Plant \n  (See Rule 10-002 (8) and CSA C22-3 No. 1-10 Clauses 5.7.3.1 & 5.7.3.3.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Wire or Conductor', 'Buildings',' ',' ',' ',' ',' ',' ',' ', 'Signs, billboards, lamp and traffic sign standards, above ground pipelines, and similar plant',' ',' ',' ',' ',' ',' ',' '],
        [' ', 'Horizontal to surface', ' ', ' ', ' ', 'Vertical to surface', ' ', ' ', ' ', 'Horizontal to surface', ' ', ' ', ' ', 'Vertical to surface', ' ', ' ', ' '],
        [' ', 'Basic (m)', 'Voltage Adder (m)', "AEUC Total (m)", "Design Clearance (m)", 'Basic (m)', 'Voltage Adder (m)', "AEUC Total (m)", "Design Clearance (m)", 'Basic (m)', 'Voltage Adder (m)', "AEUC Total (m)", "Design Clearance (m)", 'Basic (m)', 'Voltage Adder (m)', "AEUC Total (m)", "Design Clearance (m)"],
              ]
    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(2):
        row = [AEUC7_titles[i], AEUC_Table7_data['Building Horizontal to Surface Basic'][i], AEUC_Table7_data['Building Horizontal to Surface Voltage Adder'][i], AEUC_Table7_data['Building Horizontal to Surface AEUC Total'][i], AEUC_Table7_data['Building Horizontal to Surface Design Clearance'][i],
               AEUC_Table7_data['Building Vertical to Surface Basic'][i], AEUC_Table7_data['Building Vertical to Surface Voltage Adder'][i], AEUC_Table7_data['Building Vertical to Surface AEUC Total'][i], AEUC_Table7_data['Building Vertical to Surface Clearance'][i],\
               AEUC_Table7_data['Obstacle Horizontal to Surface Basic'][i], AEUC_Table7_data['Obstacle Horizontal to Surface Voltage Adder'][i], AEUC_Table7_data['Obstacle Horizontal to Surface AEUC Total'][i], AEUC_Table7_data['Obstacle Horizontal to Surface Clearance'][i],\
               AEUC_Table7_data['Obstacle Vertical to Surface Basic'][i], AEUC_Table7_data['Obstacle Vertical to Surface Voltage Adder'][i], AEUC_Table7_data['Obstacle Vertical to Surface AEUC Total'][i], AEUC_Table7_data['Obstacle Vertical to Surface Clearance'][i]]
        AEUC_7.append(row)

    #This is retrieving the number of rows in each table array
    n_row_AEUC_table_7 = len(AEUC_7)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_AEUC_table_7 = len(AEUC_7[7])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_AEUC_7 = []
    for i in range(n_row_AEUC_table_7):
        if i < 3:
            list_range_color_AEUC_7.append((i + 1, 1, i + 1, n_column_AEUC_table_7, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_AEUC_7.append((i + 1, 1, i + 1, n_column_AEUC_table_7, color_bkg_data_1))
        else:
            list_range_color_AEUC_7.append((i + 1, 1, i + 1, n_column_AEUC_table_7, color_bkg_data_2))

    # define cell format
    cell_format_AEUC_7 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_AEUC_table_7, 'center'), (4, 1, 6, 1, 'center'), (4, 2, 4, 9, 'center'), (4, 10, 4, 17, 'center'), (5, 2, 5, 5, 'center'), (5, 6, 5, 9, 'center'), (5, 10, 5, 13, 'center'), (5, 14, 5, 17, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_AEUC_7,
        'range_border': [(1, 1, 1, n_column_AEUC_table_7), (2, 1, n_row_AEUC_table_7, n_column_AEUC_table_7)],
        'row_height': [(1, 30)],
        'column_width': [(1, 50)] + [(i + 1, 20) for i in range(1, n_column_AEUC_table_7)],
    }

    # define some footer notes
    footer_AEUC7 = ['Assumes that conductor is neither insulated nor grounded, not enclosed in effectively grounded metallic sheath.', 'Basic clearances from AUEC code 2022 table 5.']

    # define the worksheet
    AEUC_Table_7 = {
        'ws_name': 'AEUC Table 7',
        'ws_content': AEUC_7,
        'cell_range_style': cell_format_AEUC_7,
        'footer': footer_AEUC7
    }
#endregion

#CSA Table 2
#region
    #Creating the title blocks. etc
    CSA2_cell00 = 'Over land, or alongside land (within the maximum horizontal swing plus flashover distance) likely to be travelled by road vehicles (including roadways, streets, lanes, alleys, driveways and enterances); H = 4.15m'
    CSA2_cell10 = 'Over or land likely to be travelled by agricultural or other equipment, including access roads to farm fields or entrances to farmyards. Over the right-of-way of underground pipelines; H = 4.15m'
    CSA2_cell20 = 'Over land likely to be travelled by off-road vehicles, riders on horses or other large animals; H = 3.6m'
    CSA2_cell30 = 'Over land within the road right-of-way inaccessible to road vehicles; H = 2.9m'
    CSA2_cell40 = 'Over walkways or land normally accessible only to pedestrians, snowmobiles, and personal-use all-terrain vehicles; H = 2.4m'
    CSA2_cell50 = 'Above top of rails at railway crossings; H = 6.7m'
    voltage = inputs['p2p_voltage']
    elevation = Altitude
    voltage_range = CSA_Table2_data['Voltage range']
    col = 'Col. ' + CSA_Table2_data['Column #']

    CSA2_titles = np.array([CSA2_cell00, CSA2_cell10, CSA2_cell20, CSA2_cell30, CSA2_cell40, CSA2_cell50])
    
    CSA2 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 2 \n Minimum Vertical Design Clearances above Ground or Rails, ac \n (See clauses 5.3.1.1, 5.7.4.1 and A.5.3.1 and tables 9 and 11.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase), Site Elevation: '+ str(elevation) +' m'],
        [' '],
        [' '],
        [' ', 'Guys, messengers, communication, span & lightning protection wires; communication cables', ' ', ' ', ' ', ' ', 'Open Supply Conductors and Service Conductors, ac' + str(voltage_range) + ' kV'],
        [' '],
        [' ', 'Col. II', ' ', ' ', ' ', ' ' , col],
        [' ', 'Basic (m)', 'Re-pave Adder (m)', "Snow Adder (m)", "CSA Total (m)", "Design Clearance (m)", 'Basic (m)', "Altitude Adder (m)", 'Re-pave Adder (m)', "Snow Adder (m)", "CSA Total (m)", "Design Clearance (m)"],
              ]
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(6):
        row = [CSA2_titles[i], CSA_Table2_data['Neutral Basic (m)'][i], CSA_Table2_data['Re-pave Adder (m)'][i], CSA_Table2_data['Snow Adder (m)'][i], CSA_Table2_data['Neutral CSA Total (m)'][i], CSA_Table2_data['Neutral Design Clearance (m)'][i], \
               CSA_Table2_data['Basic (m)'][i], CSA_Table2_data['Altitude Adder (m)'][i], CSA_Table2_data['Re-pave Adder (m)'][i], CSA_Table2_data['Snow Adder (m)'][i], CSA_Table2_data['CSA Total (m)'][i], CSA_Table2_data['Design Clearance (m)'][i]]
        CSA2.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_2 = len(CSA2)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_2 = len(CSA2[7])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA2 = []
    for i in range(n_row_CSA_table_2):
        if i < 3:
            list_range_color_CSA2.append((i + 1, 1, i + 1, n_column_CSA_table_2, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA2.append((i + 1, 1, i + 1, n_column_CSA_table_2, color_bkg_data_1))
        else:
            list_range_color_CSA2.append((i + 1, 1, i + 1, n_column_CSA_table_2, color_bkg_data_2))

    # define cell format
    cell_format_CSA_2 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_2, 'center'), (4, 2, 5, 6, 'center'), (4, 1, 6, 1, 'center'), (4, 7, 5, 12, 'center'), (6, 2, 6, 6, 'center'), (6, 7, 6, 12, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA2,
        'range_border': [(1, 1, 1, n_column_CSA_table_2), (2, 1, n_row_CSA_table_2, n_column_CSA_table_2)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 20) for i in range(n_column_CSA_table_2)],
    }

    # define some footer notes
    footer_CSA2 = ['Re-pave addition to CSA C22.3 No.1 clause 5.3.1.1 (c) & (d).', 'Snow addition to CSA C22.3 No.1 clause 5.3.1.1 (e) table D1.', 'Altitude addition to CSA C22.3 No.1 clause 5.3.1.1 (b).',\
                   'For places that permit the combined vehicle and load height to exceed 4.15 m, the applicable clearance specified in rows I and II shall be increased by the amount by which the allowable combined vehicle and load height exceeds 4.15 m. E.g. High road corridors, railway operating yards, see clauses 5.3.1.2 and A.5.3.1.2.',\
                   'Where noted, minimum clearances shall be based on induced electrostatic steady-state currents. They shall be calculated considering but not limited to the following; line overvoltage, line and conductor configuration, conductor diameter, largest expected vehicle, and line-to-road crossing angle.']

    # define the worksheet
    CSA_Table_2 = {
        'ws_name': 'CSA Table 2',
        'ws_content': CSA2,
        'cell_range_style': cell_format_CSA_2,
        'footer': footer_CSA2
    }
#endregion

#CSA Table 3
#region
    #Creating the title blocks. etc
    CSA3_cell00 = 'Minor waterways'
    CSA3_cell10 = 'Shallow or fast-moving waterways capable of being used by canoes and paddle boats in isolated areas where motor boats are not expected. \n Creeks and streams: W = 3–50 m and D < 1 m \n Ponds: A < 8 ha and D < 1 m \n H = 4.0 m'
    CSA3_cell20 = 'Shallow or fast-moving waterways capable of being used by motorboats with antennas and unable to support masted vessels \n Creeks and streams: W = 3–50 m and D < 1 m \n Ponds: A < 8 ha and D < 1 m  \n H = 6.0 m'
    CSA3_cell30 = 'Small lakes and rivers used by masted vessels \n Rivers: W = 3–50 m and D > 1 m \n Ponds and lakes: A < 8 ha and D > 1 m \n H = 8.0 m'
    CSA3_cell40 = 'Small resort lakes, medium-sized rivers and reservoirs, rivers connecting lakes, and crossings adjacent to bridges and roads \n Rivers: W = 50–500 m  \n Lakes/reservoirs: 8 ha < A < 80 ha  \n H = 10.0 m'
    CSA3_cell50 = 'Large lakes, reservoirs, and main rivers in resort areas \n Rivers: W > 500 m  \n Lakes/reservoirs 80 ha < A < 800 ha  \n H = 12.0 m'
    CSA3_cell60 = 'Main lakes on main navigation routes and marinas \n A > 800 ha \n H = 14.0 m'
    voltage = inputs['p2p_voltage']
    elevation = Altitude
    voltage_range = CSA_Table3_data['Voltage range']

    CSA3_titles = np.array([CSA3_cell00, CSA3_cell10, CSA3_cell20, CSA3_cell30, CSA3_cell40, CSA3_cell50, CSA3_cell60])
    CSA3_Crossing_Class = np.arange(7)
    CSA3 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 3 \n Minimum Vertical Design Clearances above Waterways*, ac \n (See clause 5.3.3.2.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Crossing Class', 'Type of waterways crossed: \n A = water areas \n D = water depth \n W = water width \n H = reference vessel height‡',\
          'Guys, messengers, communication, span & lightning protection wires; communication cables', 'Open Supply Conductors and Service Conductors, ac ' + str(voltage_range) + ' kV', 'Minimum clearance above OHWM, m'],
        [' '],
        [' '],
        [' '],
        [' ', ' ', 'Basic (m)', "Design Clearance (m)", 'Basic (m)', "Altitude Adder (m)", "CSA Total (m)", "Design Clearance (m)"],
        [' '],
        [' '],
              ]
    
    #This is retrieving the number of columns in the row specified in the columns (Before appending)
    n_column_CSA_table_3 = len(CSA3) - 2

    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(7):
        row = [CSA3_Crossing_Class[i] ,CSA3_titles[i], CSA_Table3_data['Neutral Basic (m)'][i], CSA_Table3_data['Neutral Design Clearance (m)'][i], CSA_Table3_data['Basic (m)'][i], CSA_Table3_data['Altitude Adder (m)'][i], CSA_Table3_data['CSA Total (m)'][i], CSA_Table3_data['Design Clearance (m)'][i]]
        CSA3.append(row)

    row1 = ['7', 'Federally maintained commercial channels, rivers, harbours, or heritage canals','§','§','§','§','§','§']
    CSA3.append(row1)
    
    #This is retrieving the number of rows in each table array
    n_row_CSA_table_3 = len(CSA3)

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA3 = []
    for i in range(n_row_CSA_table_3):
        if i < 3:
            list_range_color_CSA3.append((i + 1, 1, i + 1, n_column_CSA_table_3, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA3.append((i + 1, 1, i + 1, n_column_CSA_table_3, color_bkg_data_1))
        else:
            list_range_color_CSA3.append((i + 1, 1, i + 1, n_column_CSA_table_3, color_bkg_data_2))

    # define cell format
    cell_format_CSA_3 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_3, 'center'), (4, 1, 10, 1, 'center'), (4, 2, 10, 2, 'center'), (4, 3, 7, 4, 'center'), (4, 5, 7, 8, 'center'),\
                        (8, 3, 10, 3, 'center'), (8, 4, 10, 4, 'center'), (8, 5, 10, 5, 'center'), (8, 6, 10, 6, 'center'), (8, 7, 10, 7, 'center'), (8, 8, 10, 8, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA3,
        'range_border': [(1, 1, 1, n_column_CSA_table_3), (2, 1, n_row_CSA_table_3, n_column_CSA_table_3)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 20) for i in range(n_column_CSA_table_3)],
    }

    # define some footer notes
    footer_CSA3 = ['* The clearance over a canal, river, or stream normally used to provide access for sailboats to a larger body of water shall be the same as that required for the larger body of water.',
                   '‡ Reference vessel height refers to the overall height of the vessel, including the heights of antennas or other attachments.', 
                   '§ In Canada, clearances are specified by the Transport Canada office responsible for the coastal regions, Great Lakes system, Red River–Lake Winnipeg system, Mackenzie River, and interior lakes of British Columbia. Where tide water has an effect on a body of water being crossed, the vertical design clearance shall be increased by an amount that takes into account peak tide.']

    # define the worksheet
    CSA_Table_3 = {
        'ws_name': 'CSA Table 3',
        'ws_content': CSA3,
        'cell_range_style': cell_format_CSA_3,
        'footer': footer_CSA3
    }
#endregion

#CSA Table 5
#region
    #Creating the title blocks. etc
    CSA5_cell00 = 'Live or exposed current-carrying parts of supply equipment (e.g., cable terminals, arresters, line switches) and ungrounded cases of supply equipment (e.g., transformers, regulators, capacitors)'
    CSA5_cell10 = 'Effectively grounded cases of supply equipment (e.g., transformers, regulators, capacitors)'
    CSA5_cell01 = 'Areas accessible to pedestrians only'
    CSA5_cell11 = 'Areas likely to be travelled by vehicles'
    CSA5_cell21 = 'Areas accessible to pedestrians only'
    CSA5_cell31 = 'Areas likely to be  travelled by vehicles'
    voltage = inputs['p2p_voltage']
    elevation = Altitude
    voltage_range = CSA_Table5_data['Voltage range']

    CSA5_titles = np.array([CSA5_cell00,' ', CSA5_cell10, ' '])
    CSA5_loc_titles = np.array([CSA5_cell01, CSA5_cell11, CSA5_cell21, CSA5_cell31])
    
    CSA5 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 5 \n Minimum Separations (heights) of Supply Equipment from Ground, ac \n (See clause 5.3.2.1 and A.5.3.2 and table 11.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Type of Equipment', 'Location of equipment', 'Minimum separation from ground, m'],
        [' ',' ', str(voltage_range) + ' kV'],
        [' ', ' ', 'Basic (m)', "Design Clearance (m)"],
              ]
    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(4):
        row = [CSA5_titles[i], CSA5_loc_titles[i], CSA_Table5_data['Basic (m)'][i], CSA_Table5_data['Design Clearance (m)'][i]]
        CSA5.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_5 = len(CSA5)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_5 = len(CSA5[7])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA5 = []
    for i in range(n_row_CSA_table_5):
        if i < 3:
            list_range_color_CSA5.append((i + 1, 1, i + 1, n_column_CSA_table_5, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA5.append((i + 1, 1, i + 1, n_column_CSA_table_5, color_bkg_data_1))
        else:
            list_range_color_CSA5.append((i + 1, 1, i + 1, n_column_CSA_table_5, color_bkg_data_2))

    # define cell format
    cell_format_CSA_5 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_5, 'center'), (7, 1, 8, 1, 'center'), (9, 1, 10, 1, 'center'), (4, 1, 6, 1, 'center'), (4, 2, 6, 2, 'center'), (4, 3, 4, 4, 'center'), (5, 3, 5, 4, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA5,
        'range_border': [(1, 1, 1, n_column_CSA_table_5), (2, 1, n_row_CSA_table_5, n_column_CSA_table_5)],
        'row_height': [(1, 50), (7, 60),(8, 60), (9, 50), (10, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_5)],
    }

    # define some footer notes
    footer_CSA5 = ['* The clearance over a canal, river, or stream normally used to provide access for sailboats to a larger body of water shall be the same as that required for the larger body of water.',
                   '‡ Reference vessel height refers to the overall height of the vessel, including the heights of antennas or other attachments.', 
                   '§ In Canada, clearances are specified by the Transport Canada office responsible for the coastal regions, Great Lakes system, Red River–Lake Winnipeg system, Mackenzie River, and interior lakes of British Columbia. Where tide water has an effect on a body of water being crossed, the vertical design clearance shall be increased by an amount that takes into account peak tide.']

    # define the worksheet
    CSA_Table_5 = {
        'ws_name': 'CSA Table 5',
        'ws_content': CSA5,
        'cell_range_style': cell_format_CSA_5,
        'footer': footer_CSA5
    }
#endregion

#CSA Table 6
#region
    #Creating the title blocks. etc
    CSA6_cell00 = 'Guys; messengers; communication, span, and lightning protection wires; communication cables'
    CSA6_cell10 = 'Open supply-line conductors and service conductors of 0–750 V and effectively grounded continuous metallic sheathed cables of all voltages'
    CSA6_cell01 = 'AC'

    CSA6_cell23 = 'Open supply-line conductors and cables other than those having an effectively grounded continuous metallic sheath ' + str(voltage_range) + 'kV'

    voltage = inputs['p2p_voltage']
    elevation = Altitude
    voltage_range = CSA_Table6_data['Voltage range']

    CSA6_titles = np.array([CSA6_cell00, CSA6_cell10, CSA6_cell01, ' '])
    CSA6_loc_titles = np.array([' ', ' ',CSA6_cell23 , ' '])
    
    CSA6 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 6 \n Minimum Horizontal Desgin Clearances between Wires and Railway Tracks, ac \n (See clauses 5.4.2 and 5.4.3.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Wire closest to tracks', ' ', ' ', ' ', ' ' , 'Minimum clearance, m'],
        [' ', ' ', ' ', ' ', ' ' , 'Main Tracks', ' ', 'Siding'],
        [' ', ' ', ' ', ' ', ' ' , 'Basic (m)', "Design Clearance (m)", 'Basic (m)', "Design Clearance (m)"],
              ]
    
    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_6 = len(CSA6[5])

    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(3):
        row = [CSA6_titles[i], CSA6_loc_titles[i], ' ', ' ', ' ', CSA_Table6_data['Main Basic (m)'][i], CSA_Table6_data['Main Design Clearance (m)'][i], CSA_Table6_data['Siding Basic (m)'][i], CSA_Table6_data['Siding Design Clearance (m)'][i]]
        CSA6.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_6 = len(CSA6)

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA6 = []
    for i in range(n_row_CSA_table_6):
        if i < 3:
            list_range_color_CSA6.append((i + 1, 1, i + 1, n_column_CSA_table_6, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA6.append((i + 1, 1, i + 1, n_column_CSA_table_6, color_bkg_data_1))
        else:
            list_range_color_CSA6.append((i + 1, 1, i + 1, n_column_CSA_table_6, color_bkg_data_2))

    # define cell format
    cell_format_CSA_6 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_6, 'center'), (4, 1, 6, 5, 'center'), (4, 6, 4, 9, 'center'), (5, 6, 5, 7, 'center'), (5, 8, 5, 9, 'center'), (7, 1, 7, 5, 'center'), (8, 1, 8, 5, 'center'), (9, 2, 9, 5, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA6,
        'range_border': [(1, 1, 1, n_column_CSA_table_6), (2, 1, n_row_CSA_table_6, n_column_CSA_table_6)],
        'row_height': [(1, 50), (8, 30), (9, 30)],
        'column_width': [(i + 1, 20) for i in range(n_column_CSA_table_6)],
    }

    # define some footer notes
    footer_CSA6 = ['Voltages are rms line-to-ground.',
                   'At a point of curvature of a railway track the horizontal clearances shall be increased with an increment of 25 mm added for each degree of curvature.'
                   'In addition, where the wire is closer to the low side of tracks that are at different elevations, a further increment of 2.5 mm shall be added for each millimetre of superelevation. \n (The total of these two increments shall not exceed 0.75 m, and the value of 0.75 m may be used in place of calculations.)'
                  ]

    # define the worksheet
    CSA_Table_6 = {
        'ws_name': 'CSA Table 6',
        'ws_content': CSA6,
        'cell_range_style': cell_format_CSA_6,
        'footer': footer_CSA6
    }
#endregion

#CSA Table 7
#region
    #Creating the title blocks. etc
    CSA7_cell00 = 'Main tracks (straight, level runs)'
    CSA7_cell10 = 'Sidings (straight, level runs)'

    voltage = inputs['p2p_voltage']
    elevation = Altitude

    CSA7_titles = np.array([CSA7_cell00, CSA7_cell10])
    
    CSA7 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 7 \n Minimum Horizontal Separations from Supporting Structures to Railway Tracks, ac \n (See clauses 5.5.2 and 5.5.3.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Tracks', ' ', ' ', 'Minimum horizontal separation, m'],
        [' ', ' ', ' ', 'Basic (m)', "Design Clearance (m)"],
              ]
    
    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_7 = len(CSA7[4])

    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(2):
        row = [CSA7_titles[i], ' ', ' ', CSA_Table7_data['Basic (m)'][i], CSA_Table7_data['Design Clearance (m)'][i]]
        CSA7.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_7 = len(CSA7)

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA7 = []
    for i in range(n_row_CSA_table_7):
        if i < 3:
            list_range_color_CSA7.append((i + 1, 1, i + 1, n_column_CSA_table_7, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA7.append((i + 1, 1, i + 1, n_column_CSA_table_7, color_bkg_data_1))
        else:
            list_range_color_CSA7.append((i + 1, 1, i + 1, n_column_CSA_table_7, color_bkg_data_2))

    # define cell format
    cell_format_CSA_7 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_7, 'center'), (4, 1, 5, 3, 'center'), (4, 4, 4, 5, 'center'), (6, 1, 6, 3, 'center'), (7, 1, 7, 3, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA7,
        'range_border': [(1, 1, 1, n_column_CSA_table_7), (2, 1, n_row_CSA_table_7, n_column_CSA_table_7)],
        'row_height': [(1, 50), (8, 30), (9, 30)],
        'column_width': [(i + 1, 20) for i in range(n_column_CSA_table_7)],
    }

    # define some footer notes
    footer_CSA7 = ['If there is a curve on the track refer to CSA C22.3 No.1 clause 5.4.3']

    # define the worksheet
    CSA_Table_7 = {
        'ws_name': 'CSA Table 7',
        'ws_content': CSA7,
        'cell_range_style': cell_format_CSA_7,
        'footer': footer_CSA7
    }
#endregion


    #This determines the workbook and the worksheets within the workbook
    workbook_content = [AEUC_Table_5, AEUC_Table_7, CSA_Table_2, CSA_Table_3, CSA_Table_5, CSA_Table_6, CSA_Table_7]

    #This will create the workbook with the filename specified at the top of this function
    report_xlsx_general.create_workbook(workbook_content=workbook_content, filename=filename)
    return()

if __name__ == '__main__':
    create_report_excel()
