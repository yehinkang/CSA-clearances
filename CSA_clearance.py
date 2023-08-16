""" 
BREAKDOWN OF HOW THE CSA CLEARANCES PYTHON SCRIPT WORKS:
This code takes inputs in a dictionary
This dictionary is then used to input the code into different functions (1 function per table)
The functions then call back to a spreadsheet referenceed locally using pandas, each sheet in this spreadsheet corresponds to a CSA table
From there the functions perform necessary operations and then spit out a dictionary that can then be exported to a webserver
These dictionaries are also used to put together an excel file using the report_xlsx_general script
This file is then saved locally to a selected folder
"""

#importing the following packages
import re
from datetime import datetime
import pandas as pd
import numpy as np
import sys

#The following is referenced locally commons is a file and report_xlsx_general is another python script
from commons import report_xlsx_general

#The following is being used to see full arrays being printed out in the terminal instead of [a, b, c, ..., z], this is mostly for troubleshooting
np.set_printoptions(threshold=sys.maxsize)

#This is the full filepath for the reference spreadsheet
ref_table_filepath = r"C:\Users\yekang\OneDrive - POWER Engineers, Inc\Desktop\Python Stuff\Clearance Calcs\CSA-clearances\AEUC_CSA_clearances.xlsx"
#This uses pandas to load in the excel file
ref_table= pd.ExcelFile(ref_table_filepath)

#This is the input dictionary that all functions will use
inputs = {
    'p2p_voltage': 2000,
    'Max_Overvoltage': 5,

    'Buffer_Neut':4,
    'Buffer_Live':3,

    'Design_buffer_2_obstacles':2,
    'Design_Buffer_Same_Structure':1,

    'Clearance_Rounding':0.01,
    
    #There will need to be a dropdown menu made on the webpage with matching locations made on the webpage
    'Location': "Alberta Camrose",

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
        'Is_main_wire_upper':True,
        'XING_P2P_Voltage':79,
        'Max_Overvoltage_XING': 5,

}

#The following function takes the decimal points in the dictionary and turns that into an integer to use in np.round in all functions
def decimal_points(num):
    num_str = str(num)
    if '.' in num_str:
        return len(num_str.split('.')[1])
    else:
        return 0
#Running the function so that it can be used in the np.round functions that are called on in the table functions.
Clearance_Rounding = {key: inputs[key] for key in ['Clearance_Rounding']}
Nums_after_decimal = Clearance_Rounding['Clearance_Rounding']
Numpy_round_integer = decimal_points(Nums_after_decimal)

#The following function will take a dictionary x and export it into excel using pandas(useful for troubleshooting dictionary outputs) (Can probably get rid of later)
def save2xl(x):
    #Pandas being used to export the dataframe
    Excel_export = pd.DataFrame(x)
    #desired filepath for sheet
    folder = "C:/Users/yekang/OneDrive - POWER Engineers, Inc/Desktop/Python Stuff/Clearance Calcs/Samples/"
    #This will be the name of the sheet, using current time down to the second to ensure that the file name is unique
    start = datetime.now().strftime("Clearance_Calc %Y-%d-%m-%H-%M-%S")
    fname = folder + start + ".xlsx"
    print("saving to", fname)
    Excel_export.to_excel(fname)

#The following function will take in Longitude and Latitude and return the closest location on which there is information
#This is set to work with either locations or longitude/latitude
#CSA_C22.3_No.1_table_D2 is being referenced
def location_lookup(Location, Northing_deg, Northing_mins, Northing_seconds, Westing_deg, Westing_mins, Westing_seconds):
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

    if Place_name.isin([Location]).any():

        Closest_place_name = Location
        position = Place_name[Place_name == Location].index

        Max_snow_depth = Snow_Depth.iloc[position].iloc[0]

        Custom_Elevation = {key: inputs[key] for key in ['Custom_Elevation']}
        if Custom_Elevation['Custom_Elevation'] == 0:
            Closest_elevation = Elevation.iloc[position]
        else:
            Closest_elevation = Custom_Elevation['Custom_Elevation']

    else:
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
loc_lookup_input = {key: inputs[key] for key in ['Location', 'Northing_deg', 'Northing_mins', 'Northing_seconds', 'Westing_deg', 'Westing_mins', 'Westing_seconds']}
Place, Altitude, Snow_Depth = location_lookup(**loc_lookup_input)

#ALL FUNCTION USE KILOVOLTS AND METERS AS INPUTS

#The following function is for AEUC table 5
def AEUC_table5_clearance(p2p_voltage, Buffer_Neut, Buffer_Live, Max_Overvoltage):

    #Getting Phase to Ground voltage
    voltage = p2p_voltage / np.sqrt(3)

    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for AEUC Table 5 within the reference table file
    AEUC_5 = pd.read_excel(ref_table, "AEUC Table 5")
    
    #this is selecting a column for guy wires, comms.etc
    AEUC_neut_clearance = AEUC_5['Neutral']

    #This next little bit will determine which column should be referenced for conductors
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
        AEUC_clearance = AEUC_5['318kV']
        voltage_range = '500 kV'
        column = "VII"
    else: 
        #this column is for 1000 kV Pole-Pole
        AEUC_clearance = AEUC_5['318kV']
        AEUC_clearance = np.zeros(len(AEUC_clearance))

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
    AEUC_neut_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', AEUC_neut_clearance)
    Repave_Adder = np.char.mod(f'%0.{Numpy_round_integer}f', Repave_Adder)
    Snow_Adder = np.char.mod(f'%0.{Numpy_round_integer}f', Snow_Adder)
    AEUC_total_clearance_neut = np.char.mod(f'%0.{Numpy_round_integer}f', AEUC_total_clearance_neut)
    Design_clearance_neut = np.char.mod(f'%0.{Numpy_round_integer}f', Design_clearance_neut)
    AEUC_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', AEUC_clearance)
    Altitude_Adder = np.char.mod(f'%0.{Numpy_round_integer}f', Altitude_Adder)
    AEUC_total_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', AEUC_total_clearance)
    Design_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', Design_clearance)

    #Adding conductor snow and repave adders so that when voltage is out of range neutral clearance still show up
    Cond_Snow_Adder = Snow_Adder
    Cond_Repave_Adder = Repave_Adder

    #Adding an error condition in case of too high of a voltage
    DataLength = len(AEUC_clearance)
    if voltage > 318:
        AEUC_clearance = np.full(DataLength, "ERROR Voltage too high")
        Altitude_Adder = AEUC_clearance
        Cond_Repave_Adder = AEUC_clearance
        Cond_Snow_Adder = AEUC_clearance
        AEUC_total_clearance = AEUC_clearance
        Design_clearance = AEUC_clearance

        voltage_range = "Voltage out of range"
        column = "ERROR"

    #Creating a dictionary to turn into a dataframe
    data = {
        'AEUC base neutral': AEUC_neut_clearance,
        'Re-pave Adder': Repave_Adder,
        'Snow Adder': Snow_Adder,
        'AEUC total neutral': AEUC_total_clearance_neut,
        'Design clearance neutral': Design_clearance_neut,

        'AEUC base clearance': AEUC_clearance,
        'Altitude Adder': Altitude_Adder,
        'Cond Re-pave Adder': Cond_Repave_Adder,
        'Cond Snow Adder': Cond_Snow_Adder,
        'AEUC total': AEUC_total_clearance,
        'Design clearance': Design_clearance,

        'Voltage range': voltage_range,
        'Column #': column,
        }
    return data
#Running the AEUC table 5 function to get a dictionary that can be used in a webapp and used in the export to excel code later on
AEUC5_input = {key: inputs[key] for key in ['p2p_voltage', 'Buffer_Neut', 'Buffer_Live', 'Max_Overvoltage']}
AEUC_Table5_data  = AEUC_table5_clearance(**AEUC5_input)

#The following function is for AEUC table 7
def AEUC_table7_clearance(p2p_voltage, Design_buffer_2_obstacles, Max_Overvoltage):

    #Getting phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for AEUC Table 7 within the reference table file
    AEUC_7 = pd.read_excel(ref_table, "AEUC Table 7")

    #This will look a bit different since we'll be finding data based on rows vs. columns.
    #Since .loc is being used to find the row a column must be chosen to use 
    AEUC_7.set_index("Wire or Conductor", inplace = True)
    
    #Specify the row that is being used for this
    AEUC_neut_clearance_basic = AEUC_7.loc['Guys communications cables, and drop wires'].values
    AEUC_neut_clearance = Design_buffer_2_obstacles * np.ones(len(AEUC_neut_clearance_basic))

    #Figuring out the clearance to use with the conductor voltage
    if 0 < voltage <= 0.75:
            #Specify the row that is being used for this
            AEUC_clearance = AEUC_7.loc['0 - 750 V Neither insulated or grounded or enclosed in effectively grounded metallic sheath'].values
            AEUC_voltage_adder = 0
    elif 0.75 < voltage <= 22:
            #Specify the row that is being used for this
            AEUC_clearance = AEUC_7.loc['0.75 - 22kV Not enclosed in  effectively grounded metallic sheath'].values
            AEUC_voltage_adder = 0
    else:
        #Specify the row that is being used for this
            AEUC_clearance = AEUC_7.loc['Over 22kV'].values
            AEUC_voltage_adder = (voltage - 22) * 0.01

    #Turning the rows into columns for easier display
    #Indexing certain elements in the array BA = basic, VA = Voltage Adder, AT = AEUC Total, Design_Clearance = Design Clearance
    H2building_BA = np.array([AEUC_neut_clearance_basic[0], AEUC_clearance[0]])
    H2building_VA = np.array([0, AEUC_voltage_adder])
    H2building_AT = np.array([AEUC_neut_clearance_basic[0], (AEUC_clearance[0] + AEUC_voltage_adder)])
    H2building_Design_Clearance = np.array([AEUC_neut_clearance[0], (AEUC_clearance[0] + AEUC_voltage_adder + Design_buffer_2_obstacles)])

    V2building_BA = np.array([AEUC_neut_clearance_basic[1], AEUC_clearance[1]])
    V2building_VA = H2building_VA
    V2building_AT = np.array([AEUC_neut_clearance_basic[1], (AEUC_clearance[1] + AEUC_voltage_adder)])
    V2building_Design_Clearance = np.array([AEUC_neut_clearance[1], (AEUC_clearance[1] + AEUC_voltage_adder + Design_buffer_2_obstacles)])

    H2object_BA = np.array([AEUC_neut_clearance_basic[2], AEUC_clearance[2]])
    H2object_VA = H2building_VA
    H2object_AT = np.array([AEUC_neut_clearance_basic[2], (AEUC_clearance[2] + AEUC_voltage_adder)])
    H2object_Design_Clearance = np.array([AEUC_neut_clearance[2], (AEUC_clearance[2] + AEUC_voltage_adder + Design_buffer_2_obstacles)])

    V2object_BA = np.array([AEUC_neut_clearance_basic[3], AEUC_clearance[3]])
    V2object_VA = H2building_VA
    V2object_AT = np.array([AEUC_neut_clearance_basic[3], (AEUC_clearance[3] + AEUC_voltage_adder)])
    V2object_Design_Clearance = np.array([AEUC_neut_clearance[3], (AEUC_clearance[3] + AEUC_voltage_adder + Design_buffer_2_obstacles)])

    #The following is needed to round and display the numbers to the amount of decimal points requested in the inputs
    H2building_BA = np.char.mod(f'%0.{Numpy_round_integer}f', H2building_BA)
    H2building_VA = np.char.mod(f'%0.{Numpy_round_integer}f', H2building_VA)
    H2building_AT = np.char.mod(f'%0.{Numpy_round_integer}f', H2building_AT)
    H2building_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', H2building_Design_Clearance)

    V2building_BA = np.char.mod(f'%0.{Numpy_round_integer}f', V2building_BA)
    V2building_VA = np.char.mod(f'%0.{Numpy_round_integer}f', V2building_VA)
    V2building_AT = np.char.mod(f'%0.{Numpy_round_integer}f', V2building_AT)
    V2building_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', V2building_Design_Clearance)

    H2object_BA = np.char.mod(f'%0.{Numpy_round_integer}f', H2object_BA)
    H2object_VA = np.char.mod(f'%0.{Numpy_round_integer}f', H2object_VA)
    H2object_AT = np.char.mod(f'%0.{Numpy_round_integer}f', H2object_AT)
    H2object_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', H2object_Design_Clearance)

    V2object_BA = np.char.mod(f'%0.{Numpy_round_integer}f', V2object_BA)
    V2object_VA = np.char.mod(f'%0.{Numpy_round_integer}f', V2object_VA)
    V2object_AT = np.char.mod(f'%0.{Numpy_round_integer}f', V2object_AT)
    V2object_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', V2object_Design_Clearance)

    Categories = np.array(["Guys, communication cables, and drop wires", "Supply conductors"])

    #Creating a dictionary to turn into a dataframe
    data = {
        'Wire or Conductor': Categories,
        'Building Horizontal to Surface Basic': H2building_BA,
        'Building Horizontal to Surface Voltage Adder': H2building_VA,
        'Building Horizontal to Surface AEUC Total': H2building_AT,
        'Building Horizontal to Surface Design Clearance': H2building_Design_Clearance,

        'Building Vertical to Surface Basic': V2building_BA,
        'Building Vertical to Surface Voltage Adder': V2building_VA,
        'Building Vertical to Surface AEUC Total': V2building_AT,
        'Building Vertical to Surface Clearance': V2building_Design_Clearance,

        'Obstacle Horizontal to Surface Basic': H2object_BA,
        'Obstacle Horizontal to Surface Voltage Adder': H2object_VA,
        'Obstacle Horizontal to Surface AEUC Total': H2object_AT,
        'Obstacle Horizontal to Surface Clearance': H2object_Design_Clearance,

        'Obstacle Vertical to Surface Basic': V2object_BA,
        'Obstacle Vertical to Surface Voltage Adder': V2object_VA,
        'Obstacle Vertical to Surface AEUC Total': V2object_AT,
        'Obstacle Vertical to Surface Clearance': V2object_Design_Clearance,
        }
    
    return data
AEUC7_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_buffer_2_obstacles', 'Max_Overvoltage']}
AEUC_Table7_data = AEUC_table7_clearance(**AEUC7_input)

#The following function is for CSA Table 2
def CSA_Table_2_clearance(p2p_voltage, Buffer_Neut, Buffer_Live, Max_Overvoltage):
    
    #Adding the voltage multiplier
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 2 within the reference table file
    CSA_2 = pd.read_excel(ref_table, "CSA table 2")

    #Specify the column that is being used for this
    CSA_2_neut_clearance = CSA_2['Guys, messengers, communication, span & lightning protection wires; communication cables']

    if  0 < voltage <= 0.75: 
        #this column is for 120-660V
        CSA_2_clearance = CSA_2['0-0.75']
        voltage_range = ' 0-0.75 kV'
        column = 'III'
    elif voltage <=22: 
        #this column is for 4, 13 & 25 kV
        CSA_2_clearance = CSA_2['0.75-22']
        voltage_range = ' > 0.75 ≤ 22'
        column = 'IV'
    elif voltage <=50: 
        #this column is for 35, 69 & 72 kV
        CSA_2_clearance = CSA_2['22-50']
        voltage_range = ' > 22 ≤ 50'
        column = 'V'
    elif voltage <=90: 
        #this column is for 138 & 144 kV
        CSA_2_clearance = CSA_2['50-90']
        voltage_range = ' > 50 ≤ 90'
        column = 'VI'
    elif voltage <=120: 
        #this column is for 240 kV
        CSA_2_clearance = CSA_2['90-120']
        voltage_range = ' > 90 ≤ 120'
        column = 'VII'
    elif voltage <=150: 
        CSA_2_clearance = CSA_2['120-150']
        voltage_range = ' > 120 ≤ 150'
        column = 'VIII'
    elif voltage <=190: 
        CSA_2_clearance = CSA_2['150-190']
        voltage_range = ' > 150 ≤ 190'
        column = 'IX'
    elif voltage <= 219: 
        #this column is for 360 kV
        CSA_2_clearance = CSA_2['219 kV']
        voltage_range = ' 219 (360)'
        column = 'X'
    elif voltage <= 318: 
        #this column is for 500 kV
        CSA_2_clearance = CSA_2['318 kV']
        voltage_range = ' 318 (500)'
        column = 'XI'
    elif voltage <= 442:
        #this column is for 735 kV
        CSA_2_clearance = CSA_2['442 kV']
        voltage_range = ' 442 (735)'
        column = 'XII'
    else:
        CSA_2_clearance = CSA_2['442 kV']
        CSA_2_clearance = np.zeros(len(CSA_2_clearance))
        voltage_range = ' Voltage out of range'
        column = 'ERROR'

    Repave_Adder = np.array([0.225,0,0,0.225,0,0.3], dtype=float)

    #Need to find mean annual snow depth
    Snow_Adder = Snow_Depth * np.array([0,1,1,1,1,0], dtype=float)

    #Need to find altitude
    if Altitude > 1000:
        Altitude_Adder = (Altitude - 1000)/100 * 0.01 * CSA_2_clearance
    else:
        Altitude_Adder = np.array([0,0,0,0,0,0], dtype=float)

    #Getting the clearance arrays
    Neutral_CSA_Total = Snow_Adder + Repave_Adder + CSA_2_neut_clearance
    CSA_Total = Snow_Adder + Repave_Adder + Altitude_Adder + CSA_2_clearance

    Neutral_Design_Clearance = Buffer_Neut + Neutral_CSA_Total
    CSA_Design_Clearance = Buffer_Live + CSA_Total

    #Rounding all the values in the arrays
    
    CSA_2_neut_clearance = np.round(CSA_2_neut_clearance, Numpy_round_integer)
    Repave_Adder = np.round(Repave_Adder, Numpy_round_integer)
    Snow_Adder = np.round(Snow_Adder, Numpy_round_integer)
    Neutral_CSA_Total = np.round(Neutral_CSA_Total, Numpy_round_integer)
    Neutral_Design_Clearance = np.round(Neutral_Design_Clearance, Numpy_round_integer)
    CSA_2_clearance = np.round(CSA_2_clearance, Numpy_round_integer)
    Altitude_Adder = np.round(Altitude_Adder, Numpy_round_integer)
    CSA_Total = np.round(CSA_Total, Numpy_round_integer)
    CSA_Design_Clearance = np.round(CSA_Design_Clearance, Numpy_round_integer)

    CSA_2_neut_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_2_neut_clearance)
    Repave_Adder = np.char.mod(f'%0.{Numpy_round_integer}f', Repave_Adder)
    Snow_Adder = np.char.mod(f'%0.{Numpy_round_integer}f', Snow_Adder)
    Neutral_CSA_Total = np.char.mod(f'%0.{Numpy_round_integer}f', Neutral_CSA_Total)
    Neutral_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', Neutral_Design_Clearance)
    CSA_2_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_2_clearance)
    Altitude_Adder = np.char.mod(f'%0.{Numpy_round_integer}f', Altitude_Adder)
    CSA_Total = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_Total)
    CSA_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_Design_Clearance)

    #Adding Conductor repave and snow adders so that in error cases the neutral clearances still show
    Cond_Snow_Adder = Snow_Adder
    Cond_Repave_Adder = Repave_Adder

    #Adding errors for overvoltage conditions
    DataLength = len(CSA_2_clearance)
    if voltage > 442:
        CSA_2_clearance = np.full(DataLength, "ERROR Voltage too high")
        Cond_Repave_Adder = CSA_2_clearance
        Cond_Snow_Adder = CSA_2_clearance
        Altitude_Adder = CSA_2_clearance
        CSA_Total = CSA_2_clearance
        CSA_Design_Clearance = CSA_2_clearance

    #Creating a dictionary that will hold all the values that are outputted
    data = {
        'Neutral Basic (m)': CSA_2_neut_clearance,
        'Re-pave Adder (m)': Repave_Adder,
        'Snow Adder (m)': Snow_Adder,
        'Neutral CSA Total (m)': Neutral_CSA_Total,
        'Neutral Design Clearance (m)': Neutral_Design_Clearance,

        'Basic (m)': CSA_2_clearance,
        'Altitude Adder (m)': Altitude_Adder,
        'Cond Re-pave Adder (m)': Cond_Repave_Adder,
        'Cond Snow Adder (m)': Cond_Snow_Adder,
        'CSA Total (m)': CSA_Total,
        'Design Clearance (m)': CSA_Design_Clearance,

        #This will only be used for the excel doc
        'Voltage range': voltage_range,
        'Column #': column,
        }
    
    return data
CSA2_input = {key: inputs[key] for key in ['p2p_voltage', "Buffer_Neut", "Buffer_Live", 'Max_Overvoltage']}
CSA_Table2_data = CSA_Table_2_clearance(**CSA2_input)

#The following function is for CSA Table 3
def CSA_Table_3_clearance(p2p_voltage, Buffer_Neut, Buffer_Live, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_3 = pd.read_excel(ref_table, "CSA table 3")

    #Neutral clearance column being found on spreadsheet
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
    #Design_Clearance = Design Clearance
    Neutral_Design_Clearance = Buffer_Neut + CSA_3_neut_clearance
    CSA_Design_Clearance = Buffer_Live + CSA_Total

    #Rounding the values
    CSA_3_neut_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_3_neut_clearance)
    Neutral_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', Neutral_Design_Clearance)
    CSA_3_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_3_clearance)
    Altitude_Adder = np.char.mod(f'%0.{Numpy_round_integer}f', Altitude_Adder)
    CSA_Total = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_Total)
    CSA_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_Design_Clearance)

    data = {
        'Neutral Basic (m)': CSA_3_neut_clearance,
        'Neutral Design Clearance (m)': Neutral_Design_Clearance,

        'Basic (m)': CSA_3_clearance,
        'Altitude Adder (m)': Altitude_Adder,
        'CSA Total (m)': CSA_Total,
        'Design Clearance (m)': CSA_Design_Clearance,

        'Voltage range': voltage_range,
        }

    return data
CSA3_input = {key: inputs[key] for key in ['p2p_voltage', "Buffer_Neut", "Buffer_Live", 'Max_Overvoltage']}
CSA_Table3_data = CSA_Table_3_clearance(**CSA3_input)

#The following function is for CSA Table 5
def CSA_Table_5_clearance(p2p_voltage, Buffer_Live, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

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
    elif voltage <=90: 
        CSA_5_clearance = CSA_5['>50kV<90kV']
        Voltage_range = '> 50 ≤ 90'
    else: 
        CSA_5_clearance = CSA_5['>50kV<90kV']

    #Design_Clearance = Design Clearance
    Design_Clearance = CSA_5_clearance + Buffer_Live

    #Rounding the clearance values
    CSA_5_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_5_clearance)
    Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', Design_Clearance)

    #Adding an error condition in case of too high of a voltage
    DataLength = len(CSA_5_clearance)
    if voltage > 90:
        CSA_5_clearance = np.full(DataLength, "ERROR Voltage too high")
        Design_Clearance = CSA_5_clearance
        Voltage_range = 'Error'



    #Creating a dictionary to store all the outpt data
    data = {
        'Basic (m)': CSA_5_clearance,
        'Design Clearance (m)': Design_Clearance,

        'Voltage range': Voltage_range
        }

    return data
CSA5_input = {key: inputs[key] for key in ['p2p_voltage', "Buffer_Live", 'Max_Overvoltage']}
CSA_Table5_data = CSA_Table_5_clearance(**CSA5_input)

#The following function is for CSA Table 6
def CSA_Table_6_clearance(p2p_voltage, Buffer_Neut, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_6 = pd.read_excel(ref_table, "CSA table 6")
    
    #Again this will look a bit different since rows are being used instead of columns
    #Since .loc is being used to find the row a column must be chosen to use 
    CSA_6.set_index("Wire closest to tracks", inplace = True)

    CSA_6_neut_clearance = CSA_6.loc['Guys; messengers; communication, span, and lightning protection wires; communication cables']
    CSA_6_sub750_clearance = CSA_6.loc['Open supply-line conductors and service conductors of 0—750 V and effectively grounded continuous metallic sheathed cables of all voltages']

    if  voltage <= 22: 
        CSA_6_clearance = CSA_6.loc ['AC > 0.75 < 22 kV'].values
        Voltage_range = '> 0.75 ≤ 22 kV'
    elif voltage <=50: 
        CSA_6_clearance = CSA_6.loc ['AC > 22 < 50 kV'].values
        Voltage_range = '> 22 ≤ 50 kV'
    elif voltage <=90: 
        CSA_6_clearance = CSA_6.loc ['AC > 50 < 90 kV'].values
        Voltage_range = '> 50 ≤ 90 kV'
    elif voltage <=120: 
        CSA_6_clearance = CSA_6.loc ['AC > 90 < 120 kV'].values
        Voltage_range = '> 90 ≤ 120 kV'
    elif voltage <=150: 
        CSA_6_clearance = CSA_6.loc ['AC > 120 < 150 kV'].values
        Voltage_range = '> 120 ≤ 150 kV'
    else: 
        CSA_6_clearance = CSA_6.loc ['AC Supply conductors > 150 kV +0.01m/kV over 150kV'].values
        CSA_6_clearance = ((voltage - 150) *0.01) + CSA_6_clearance
        Voltage_range = '> 150 kV'

    #Design_Clearance = Design Clearance
    Design_Clearance_neut = CSA_6_neut_clearance + Buffer_Neut
    Design_Clearance_sub750 = CSA_6_sub750_clearance + Buffer_Neut
    Design_Clearance = CSA_6_clearance + Buffer_Neut

    #Rounding all the values
    CSA_6_clearance = np.round(CSA_6_clearance, Numpy_round_integer)
    CSA_6_neut_clearance = np.round(CSA_6_neut_clearance, Numpy_round_integer)
    CSA_6_sub750_clearance = np.round(CSA_6_sub750_clearance, Numpy_round_integer)
    Design_Clearance = np.round(Design_Clearance, Numpy_round_integer)
    Design_Clearance_neut = np.round(Design_Clearance_neut, Numpy_round_integer)
    Design_Clearance_sub750 = np.round(Design_Clearance_sub750, Numpy_round_integer)

    #Turning the rows into columns for easier display
    Basic_Main_Track = np.array([CSA_6_neut_clearance[0], CSA_6_sub750_clearance[0], CSA_6_clearance[0]])
    Design_Clearance_Main_Track = np.array([Design_Clearance_neut[0], Design_Clearance_sub750[0], Design_Clearance[0]])
    Basic_Siding = np.array([CSA_6_neut_clearance[1], CSA_6_sub750_clearance[1], CSA_6_clearance[1]])
    Design_Clearance_Siding = np.array([Design_Clearance_neut[1], Design_Clearance_sub750[1], Design_Clearance[1]])

    Basic_Main_Track = np.char.mod(f'%0.{Numpy_round_integer}f', Basic_Main_Track)
    Design_Clearance_Main_Track = np.char.mod(f'%0.{Numpy_round_integer}f', Design_Clearance_Main_Track)
    Basic_Siding = np.char.mod(f'%0.{Numpy_round_integer}f', Basic_Siding)
    Design_Clearance_Siding = np.char.mod(f'%0.{Numpy_round_integer}f', Design_Clearance_Siding)

    #Creating a dictionary to hold outputs
    data = {
        'Main Basic (m)': Basic_Main_Track,
        'Main Design Clearance (m)': Design_Clearance_Main_Track,
        'Siding Basic (m)': Basic_Siding,
        'Siding Design Clearance (m)': Design_Clearance_Siding,

        'Voltage range': Voltage_range
        }

    return data
CSA6_input = {key: inputs[key] for key in ['p2p_voltage', "Buffer_Neut", 'Max_Overvoltage']}
CSA_Table6_data = CSA_Table_6_clearance(**CSA6_input)

#The following function is for CSA Table 7
def CSA_Table_7_clearance(Buffer_Live):

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_7 = pd.read_excel(ref_table, "CSA table 7")

    #finding a column for the basic clearance
    CSA_7clearance = CSA_7['Minimum horizontal separation, m']

    #Creating an array to add a design buffer to the basic clearance
    Buffer_Live_array = np.ones(len(CSA_7)) * Buffer_Live
    CSA7_Design_Clearance = CSA_7clearance + Buffer_Live_array

    #Rounding
    CSA_7clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_7clearance)
    CSA7_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA7_Design_Clearance)

    data = {
        'Basic (m)': CSA_7clearance,
        'Design Clearance (m)': CSA7_Design_Clearance,
        }
    return data
CSA7_input = {key: inputs[key] for key in ["Buffer_Live"]}
CSA_Table7_data = CSA_Table_7_clearance(**CSA7_input)

#The following function is for CSA Table 9
def CSA_Table_9_clearance(p2p_voltage, Buffer_Neut, Buffer_Live, Max_Overvoltage):

    #Point to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 9 within the reference table file
    CSA_9 = pd.read_excel(ref_table, "CSA table 9")

    #Basic clearance columns being found on spreadsheet
    CSA_9_build_hor = CSA_9['Buildings Horiz.']
    CSA_9_build__vert = CSA_9['Buildings Vertical']
    CSA_9_obs_hor = CSA_9['To signs, billboards, lamp and traffic sign standards, and similar plant Horiz.']
    CSA_9_obs_vert = CSA_9['To signs, billboards, lamp and traffic sign standards, and similar plant Vertical']

    #Determining the clearance for >22kV
    if voltage > 22:
        #Getting the voltage adder value
        voltage_adder = (voltage - 22) * 0.01

        #Need to put this into an array so that it only adds onto the >22kV scenario
        voltage_adder_array = np.array([0,0,0,0,0,0,voltage_adder])

        #Adding the voltage adder onto the basic clearances if voltage is high enough
        CSA_9_build_hor = CSA_9_build_hor + voltage_adder_array
        CSA_9_build__vert = CSA_9_build__vert + voltage_adder_array
        CSA_9_obs_hor = CSA_9_obs_hor + voltage_adder_array
        CSA_9_obs_vert = CSA_9_obs_vert + voltage_adder_array
    else:
        voltage_adder = 0

    #Creating an adder array, first row is for neutral, and the next 5 rows are for conductors with voltage, last row considers a voltage adder as well
    adder_array = np.array([Buffer_Neut, Buffer_Live, Buffer_Live, Buffer_Live, Buffer_Live, Buffer_Live, Buffer_Live])

    #Adding the adder to get design clearance values
    Design_Clearance_build_hor = CSA_9_build_hor + adder_array
    Design_Clearance_build__vert = CSA_9_build__vert + adder_array
    Design_Clearance_obs_hor = CSA_9_obs_hor + adder_array
    Design_Clearance_obs_vert = CSA_9_obs_vert + adder_array

    #Rounding the values
    CSA_9_build_hor = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_9_build_hor)
    CSA_9_build__vert = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_9_build__vert)
    CSA_9_obs_hor = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_9_obs_hor)
    CSA_9_obs_vert = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_9_obs_vert)

    Design_Clearance_build_hor = np.char.mod(f'%0.{Numpy_round_integer}f', Design_Clearance_build_hor)
    Design_Clearance_build__vert = np.char.mod(f'%0.{Numpy_round_integer}f', Design_Clearance_build__vert)
    Design_Clearance_obs_hor = np.char.mod(f'%0.{Numpy_round_integer}f', Design_Clearance_obs_hor)
    Design_Clearance_obs_vert = np.char.mod(f'%0.{Numpy_round_integer}f', Design_Clearance_obs_vert)

    #Creating a dictionary that will have the exports from the table
    data = {
        'Buildings basic horizontal': CSA_9_build_hor,
        'Buildings basic vertical': CSA_9_build__vert,
        'Obstacles basic horizontal': CSA_9_obs_hor,
        'Obstacles basic vertical': CSA_9_obs_vert,

        'Buildings Design Clearance Horizontal': Design_Clearance_build_hor,
        'Buildings Design Clearance vertical': Design_Clearance_build__vert,
        'Obstacles Design Clearance Horizontal': Design_Clearance_obs_hor,
        'Obstacles Design Clearance vertical': Design_Clearance_obs_vert,
        }
    return data
CSA9_input = {key: inputs[key] for key in ['p2p_voltage',"Buffer_Neut", "Buffer_Live", 'Max_Overvoltage']}
CSA_Table9_data = CSA_Table_9_clearance(**CSA9_input)

#The following function is for CSA Table 10
def CSA_table10_clearance(p2p_voltage, Design_buffer_2_obstacles, Max_Overvoltage):

    #Getting Phase to Ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA table 10 within the reference table file
    CSA_10 = pd.read_excel(ref_table, "CSA table 10")

    #This will look a bit different since we'll be finding data based on rows vs. columns.
    #Since .loc is being used to find the row a column must be chosen to use 
    CSA_10.set_index("Supply conductor", inplace = True)

    #Finding correct clearance value from the voltage
    if  0 < voltage <= 0.75: 
        CSA_10_clearance = CSA_10.loc['0-750 v'].values
        Voltage_range = '0–750V kV'
    elif voltage <=22: 
        CSA_10_clearance = CSA_10.loc['> 0.75 < 22 kV'].values
        Voltage_range = '> 0.75 ≤ 22 kV'
    elif voltage <=50:
        CSA_10_clearance = CSA_10.loc['> 22 < 50 kV'].values
        Voltage_range = '> 22 ≤ 50 kV'
    else: 
        CSA_10_clearance = CSA_10.loc['> 50 kV ac + 0.010 m/kV over 50 kV'].values
        Voltage_range = '> 50 kV'

    #Creating a clearance adder from the design buffer
    clearance_adder = np.ones(len(CSA_10_clearance)) * Design_buffer_2_obstacles

    #Creating a voltage adder for voltages greater than 50 kV
    if voltage > 50:
        #Getting the voltage adder value
        voltage_adder = (voltage - 50) * 0.01

        #Need to put this into an array so that it only adds onto the >22kV scenario
        voltage_adder_array = np.ones(len(CSA_10_clearance)) * voltage_adder

        #Adding the voltage adder onto the basic clearances if voltage is high enough
        CSA_10_clearance = CSA_10_clearance + voltage_adder_array
    else:
        voltage_adder = 0

    #Clearance adder must be added on after the voltage adder
    Design_Clearance_CSA10_clearance = CSA_10_clearance + clearance_adder


    # The next to lines of code is to sort the two arrays into a format that goes: basic, Design_Clearance, basic, Design_Clearance ...etc
    # Stack the two 1D arrays vertically
    stacked_arrays = np.vstack((CSA_10_clearance, Design_Clearance_CSA10_clearance))
    # Flatten the resulting 2D array into a 1D array
    CSA10output = np.ravel(stacked_arrays, order='F')

    #Rounding the values
    CSA10output = np.char.mod(f'%0.{Numpy_round_integer}f', CSA10output)

    #Creating a dictionary to turn into a dataframe
    data = {
        #CSA 10 output is unique so far in that it is a row rather than a column and as such should be treated differently
        'CSA10_Clearance': CSA10output,

        'Voltage range': Voltage_range,
        }
    return data
CSA10_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_buffer_2_obstacles',  'Max_Overvoltage']}
CSA_Table10_data = CSA_table10_clearance(**CSA10_input)

#The following function is for CSA Table 11
def CSA_Table_11_clearance(p2p_voltage, Buffer_Neut, Buffer_Live, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_11 = pd.read_excel(ref_table, "CSA table 11")

    #This will look a bit different since we'll be finding data based on rows vs. columns.
    #Since .loc is being used to find the row a column must be chosen to use 
    CSA_11.set_index("Equipment or Conductors", inplace = True)

    #Supply Equipment values
    CSA11_SupplyEquipment = CSA_11.loc['Supply equipment'][0]
    CSA11_SupplyEquipment_Design_Clearance = CSA11_SupplyEquipment + Buffer_Live
    CSA11_SupplyEquipmentB = CSA_11.loc['Supply equipment'][1]

    #Neutral
    CSA11_neut_clearance = CSA_11.loc['Guys, messengers, span wires, communication circuits, secondary cable 0–750 V, and multi-grounded neutral conductors'][0]
    CSA11_neut_clearance_Design_Clearance = CSA11_neut_clearance + Buffer_Neut
    CSA11_neut_clearanceB = CSA_11.loc['Guys, messengers, span wires, communication circuits, secondary cable 0–750 V, and multi-grounded neutral conductors'][1]

    if  0 < voltage <= 0.75: 
        CSA11_clearance = CSA_11.loc['Guys, messengers, span wires, communication circuits, secondary cable 0–750 V, and multi-grounded neutral conductors'][0]
        CSA11_clearanceB = CSA_11.loc['Guys, messengers, span wires, communication circuits, secondary cable 0–750 V, and multi-grounded neutral conductors'][1]
        Voltage_range = '0–750V'
    elif voltage <=22: 
        CSA11_clearance = CSA_11.loc['Other supply conductors ≤ 22kV'][0]
        CSA11_clearanceB = CSA_11.loc['Other supply conductors ≤ 22kV'][1]
        Voltage_range = '> 0.75 ≤ 22'
    elif voltage <=150: 
        CSA11_clearance = CSA_11.loc['Supply conductors > 22 kV and ≤ 150 kV (6.7 + 0.01 m/kV above 22kV)'][0]
        #Adding voltage adder
        CSA11_clearance = (voltage-22)*0.01 + CSA11_clearance
        CSA11_clearanceB = CSA_11.loc['Supply conductors > 22 kV and ≤ 150 kV (6.7 + 0.01 m/kV above 22kV)'][1]
        Voltage_range = '> 22 ≤ 150'
    else: 
        CSA11_clearance = CSA_11.loc['Supply conductors > 150 kV'][0]
        #Adding voltage adder
        CSA11_clearance = (voltage-150)*0.01 + CSA11_clearance
        CSA11_clearanceB = CSA_11.loc['Supply conductors > 150 kV'][1]
        Voltage_range = '> 150'

    #Creating an array to add a design buffer to the basic clearance
    CSA11_Design_Clearance = CSA11_clearance + Buffer_Live

    #Rounding
    CSA11_SupplyEquipment = np.char.mod(f'%0.{Numpy_round_integer}f', CSA11_SupplyEquipment)
    CSA11_SupplyEquipment_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA11_SupplyEquipment_Design_Clearance
                                                         )
    CSA11_neut_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA11_neut_clearance)
    CSA11_neut_clearance_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA11_neut_clearance_Design_Clearance)

    CSA11_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA11_clearance)
    CSA11_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA11_Design_Clearance)


    #Creating the 3 columns that are needed to be outputted to form the dictionary output
    CSA11_A = np.array([CSA11_SupplyEquipment, CSA11_neut_clearance, CSA11_clearance])
    CSA11_A_Design_Clearance = np.array([CSA11_SupplyEquipment_Design_Clearance, CSA11_neut_clearance_Design_Clearance, CSA11_Design_Clearance])
    CSA11_B = np.array([CSA11_SupplyEquipmentB, CSA11_neut_clearanceB, CSA11_clearanceB])

    #Creating the dictionary output
    data = {
        'Basic': CSA11_A,
        'Design Clearance': CSA11_A_Design_Clearance,
        'B-Measured Vertically Over Land': CSA11_B,

        'Voltage range': Voltage_range,
        }
    return data
CSA11_input = {key: inputs[key] for key in ['p2p_voltage', "Buffer_Neut", "Buffer_Live",  'Max_Overvoltage']}
CSA_Table11_data = CSA_Table_11_clearance(**CSA11_input)

#The following function is for CSA Table 13
def CSA_Table_13_clearance(p2p_voltage, Design_buffer_2_obstacles, XING_P2P_Voltage, Max_Overvoltage, Max_Overvoltage_XING):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    XING_voltage = XING_P2P_Voltage / np.sqrt(3)
    #Adding Voltage Multipliers
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier
    Voltage_multiplier_XING = Max_Overvoltage_XING/100 + 1
    XING_voltage = XING_voltage * Voltage_multiplier_XING 


    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_13 = pd.read_excel(ref_table, "CSA table 13")

    #Index being set so that row can be found using name
    CSA_13.set_index("Type of line wire, cable, or other plant being crossed over", inplace = True)

    #From the way this table works a specfific cell must be selected using row and column the following if else statements will pick the proper row and column

    #Figuring out the correct row name
    if  0 < voltage <= 0.75: 
        voltage_range_over = 'AC 0-0.75'
    elif voltage <=22: 
        voltage_range_over = 'AC < 0.75 ≤ 22'
    elif voltage <=50: 
        voltage_range_over = 'AC < 22 ≤ 50'
    elif voltage <=90: 
        voltage_range_over = 'AC < 50 ≤ 90'
    elif voltage <=120: 
        voltage_range_over = 'AC < 90 ≤ 120'
    elif voltage <=150: 
        voltage_range_over = 'AC < 120 ≤ 150'
    elif voltage <=190: 
        voltage_range_over = 'AC < 150 ≤ 190'
    elif voltage <=220: 
        voltage_range_over = 'AC < 190 ≤ 220'
    elif voltage <=320: 
        voltage_range_over = 'AC < 220 ≤ 320'
    elif voltage <=425: 
        voltage_range_over = 'AC < 320 ≤ 425'
    else: 
        voltage_range_over = 'AC 0-0.75'

    #Figuring out the correct column name
    if  0 < XING_voltage <= 0.75: 
        voltage_range_XING = 'AC 0-0.75'
    elif XING_voltage <=22: 
        voltage_range_XING = 'AC < 0.75 ≤ 22'
    elif XING_voltage <=50: 
        voltage_range_XING = 'AC < 22 ≤ 50'
    elif XING_voltage <=90: 
        voltage_range_XING = 'AC < 50 ≤ 90'
    elif XING_voltage <=120: 
        voltage_range_XING = 'AC < 90 ≤ 120'
    elif XING_voltage <=150: 
        voltage_range_XING = 'AC < 120 ≤ 150'
    elif XING_voltage <=190: 
        voltage_range_XING = 'AC < 150 ≤ 190'
    elif XING_voltage <=220: 
        voltage_range_XING = 'AC < 190 ≤ 220'
    elif voltage <=320: 
        voltage_range_XING = 'AC < 220 ≤ 320'
    elif voltage <=425: 
        voltage_range_XING = 'AC < 320 ≤ 425'
    else: 
        voltage_range_XING = 'AC 0-0.75'

    #Row/column name for comms
    row_comms = 'Communication wires and cables'
    row_guys = 'Guys, span wires, and aerial grounding wires'
    col_comm_guy = 'Guys, span wires, aerial grounding conductors, and communication wires and cables'

    #Supply Equipment values
    CSA13_comm2comm = CSA_13.at[ row_comms , col_comm_guy ]
    CSA13_ac2comm = CSA_13.at[ row_comms , voltage_range_over ]

    CSA13_comm2ac = CSA_13.at[ voltage_range_XING , col_comm_guy ]
    CSA13_ac2ac = CSA_13.at[ voltage_range_XING , voltage_range_over ]

    CSA13_comm2guy = CSA_13.at[ row_guys , col_comm_guy ]
    CSA13_ac2guy = CSA_13.at[ row_guys , voltage_range_over ]

    #Creating arrays
    CAS13_guy_basic = np.array([CSA13_comm2comm, CSA13_comm2ac, CSA13_comm2guy])
    CAS13_ac_basic = np.array([CSA13_ac2comm, CSA13_ac2ac, CSA13_ac2guy])

    #Creating an array to add a design buffer to the basic clearance
    buffer_array = Design_buffer_2_obstacles * np.ones(len(CAS13_guy_basic))

    #Adding the design buffer to get design clearance (Design_Clearance)
    CSA13_from_guy_Design_Clearance = CAS13_guy_basic + buffer_array
    CSA13_from_ac_Design_Clearance = CAS13_ac_basic + buffer_array

    #Rounding
    CAS13_guy_basic = np.char.mod(f'%0.{Numpy_round_integer}f', CAS13_guy_basic)
    CAS13_ac_basic = np.char.mod(f'%0.{Numpy_round_integer}f', CAS13_ac_basic)

    CSA13_from_guy_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA13_from_guy_Design_Clearance)
    CSA13_from_ac_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA13_from_ac_Design_Clearance)

    #incase of over voltage
    if voltage > 425 or XING_voltage > 425:
        CAS13_guy_basic = np.full((3,), "ERROR voltage too high")
        CSA13_from_guy_Design_Clearance = np.full((3,), "ERROR voltage too high")
        CAS13_ac_basic = np.full((3,), "ERROR voltage too high")
        CSA13_from_ac_Design_Clearance = np.full((3,), "ERROR voltage too high")
        voltage_range_over = "ERROR Voltage too high"
        voltage_range_XING = "ERROR Voltage too high"

    #Rounding the voltage after the overvoltage function
    XING_voltage = np.char.mod(f'%0.{Numpy_round_integer}f', XING_voltage)
    voltage = np.char.mod(f'%0.{Numpy_round_integer}f', voltage)

    data = {
        'Basic guy': CAS13_guy_basic,
        'Design Clearance guy': CSA13_from_guy_Design_Clearance,
        'Basic ac': CAS13_ac_basic,
        'Design Clearance ac': CSA13_from_ac_Design_Clearance,

        'Voltage range': voltage_range_over,
        'Voltage range XING': voltage_range_XING,
        }
    
    return data
CSA13_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_buffer_2_obstacles', "XING_P2P_Voltage", 'Max_Overvoltage', 'Max_Overvoltage_XING']}
CSA_Table13_data = CSA_Table_13_clearance(**CSA13_input)

#The following function is for CSA Table 14
def CSA_Table_14_clearance(p2p_voltage, Design_buffer_2_obstacles, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_14 = pd.read_excel(ref_table, "CSA table 14")

    CSA_14_neut_clearance = CSA_14['Communication conductors and cables, span and grounding wires carried aerially']

    #Figuring out the correct row name
    if  0 < voltage <= 0.75: 
        voltage_range = '0 – 750V'
        CSA14_clearance = CSA_14['>0–750V']
    elif voltage <=22: 
        voltage_range = '> 0.75kV ≤ 22kV'
        CSA14_clearance = CSA_14['>0.75kV<22kV']
    elif voltage <=50: 
        voltage_range = '> 22kV ≤ 50kV'
        CSA14_clearance = CSA_14['>22kV<50kV']
    elif voltage <=90: 
        voltage_range = '> 50kV ≤ 90kV'
        CSA14_clearance = CSA_14['>50kV<90kV']
    elif voltage <=120: 
        voltage_range = '> 90kV ≤ 120kV'
        CSA14_clearance = CSA_14['>90kV<120kV']
    elif voltage <=150: 
        voltage_range = '> 120kV ≤ 150kV'
        CSA14_clearance = CSA_14['>120kV<150kV']
    elif voltage <=190: 
        voltage_range = '> 150kV ≤ 190kV'
        CSA14_clearance = CSA_14['>150kV<190kV']
    elif voltage <=220: 
        voltage_range = '> 190kV ≤ 220kV'
        CSA14_clearance = CSA_14['>190kV<220kV']
    elif voltage <=320: 
        voltage_range = '> 220kV ≤ 320kV'
        CSA14_clearance = CSA_14['>220kV<320kV']
    else: 
        voltage_range = '> 320kV ≤ 425kV'
        CSA14_clearance = CSA_14['>320kV<425kV']

    #Creating an array to add a design buffer to the basic clearance
    buffer_array = Design_buffer_2_obstacles * np.ones(len(CSA14_clearance))

    #Adding the design buffer to get design clearance (Design_Clearance)
    CSA14_Design_Clearance = CSA14_clearance + buffer_array
    CSA14_neut_Design_Clearance = CSA_14_neut_clearance + buffer_array

    #Rounding
    CSA14_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA14_clearance)
    CSA_14_neut_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_14_neut_clearance)

    CSA14_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA14_Design_Clearance)
    CSA14_neut_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA14_neut_Design_Clearance)

    data = {
        'Basic guy': CSA_14_neut_clearance,
        'Design Clearance guy': CSA14_neut_Design_Clearance,
        'Basic': CSA14_clearance,
        'Design Clearance': CSA14_Design_Clearance,

        'Voltage range': voltage_range,
        }
    
    return data
CSA14_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_buffer_2_obstacles', 'Max_Overvoltage']}
CSA_Table14_data = CSA_Table_14_clearance(**CSA14_input)

#The following function is for CSA Table 15
def CSA_Table_15_clearance(p2p_voltage, Design_buffer_2_obstacles, XING_P2P_Voltage, Max_Overvoltage, Max_Overvoltage_XING):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    voltage_XING = XING_P2P_Voltage / np.sqrt(3)

    #Adding Voltage Multipliers
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier
    Voltage_multiplier_XING = Max_Overvoltage_XING/100 + 1
    voltage_XING = voltage_XING * Voltage_multiplier_XING 

    comms_voltage = voltage + 0.130
    voltage = voltage + voltage_XING

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_15 = pd.read_excel(ref_table, "CSA table 15")

    #This will look a bit different since we'll be finding data based on rows vs. columns.
    #Since .loc is being used to find the row a column must be chosen to use 
    CSA_15.set_index("Sum of voltages of conductors", inplace = True)

    #Figuring out the correct row name
    if  0 < voltage <= 0.75: 
        voltage_range = '0 – 750 V'
        CSA15_clearance = CSA_15.loc['0-750V']
    else: 
        voltage_range = '> 750 V'
        CSA15_clearance = CSA_15.loc['> 750 V ac (300 + 10 mm/kV over 750)']
        CSA15_clearance = CSA15_clearance + np.ones(len(CSA15_clearance)) * 10 * (voltage - 0.75)
        #Figuring out the correct row name
    if  0 < comms_voltage <= 0.75: 
        voltage_range = '0 – 750 V'
        CSA15_Comms = CSA_15.loc['0-750V']
    else: 
        voltage_range = '> 750 V'
        CSA15_Comms = CSA_15.loc['> 750 V ac (300 + 10 mm/kV over 750)']
        CSA15_Comms = CSA15_Comms + np.ones(len(CSA15_Comms)) * 6 * (voltage - 0.75)

    #Creating an array to add a design buffer to the basic clearance
    #Adding factor of 1000 to turn meters into mm
    buffer_array = Design_buffer_2_obstacles * float(1000) * np.ones(len(CSA15_clearance))

    #Adding the design buffer to get design clearance
    CSA15_Design_Clearance = CSA15_clearance + buffer_array
    CSA15_Comms_Design_Clearance = CSA15_Comms + buffer_array

    #Rounding
    CSA15_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA15_clearance)
    CSA15_Comms = np.char.mod(f'%0.{Numpy_round_integer}f', CSA15_Comms)

    CSA15_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA15_Design_Clearance)
    CSA15_Comms_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA15_Comms_Design_Clearance)

    #Creating the arrays that will be the columns in the spreadsheet, the position has to be specified since these are arrays of len (1) o/w would give brackets around values in table
    Basic_clearance_array = np.array([CSA15_clearance[0], CSA15_Comms[0]])
    Basic_clearance_Design_Clearance = np.array([CSA15_Design_Clearance[0], CSA15_Comms_Design_Clearance[0]])

    data = {
        'Basic': Basic_clearance_array,
        'Design_Clearance': Basic_clearance_Design_Clearance,

        'Voltage range': voltage_range,
        }
    return data
CSA15_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_buffer_2_obstacles', 'XING_P2P_Voltage', 'Max_Overvoltage', 'Max_Overvoltage_XING']}
CSA_Table15_data = CSA_Table_15_clearance(**CSA15_input)

#The following function is for CSA Table 16
def CSA_Table_16_clearance(p2p_voltage, Design_buffer_2_obstacles, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_16 = pd.read_excel(ref_table, "CSA table 16")

    #This will look a bit different since we'll be finding data based on rows vs. columns.
    #Since .loc is being used to find the row a column must be chosen to use 
    CSA_16.set_index("Voltage of line conductor", inplace = True)

    #Collecting clearance for neutral and adding a design clearance buffer
    CSA16_clearance_neut = CSA_16.loc['0-5 (1000 where practicable, but in no case less than. a)150 for spans 0-6m b)230 for spans >6 and < 15m c)300 for spans >15m']
    CSA16_clearance_neut_Design_Clearance = CSA16_clearance_neut + np.ones(len(CSA16_clearance_neut)) * Design_buffer_2_obstacles
    text_neut = ' where practicable, but in no case less than. \n a)150 for spans 0-6m \n b)230 for spans >6 and < 15m \n c)300 for spans >15m'

    #Figuring out the correct row name
    if  0 < voltage <= 5: 
        #Voltage range for printing onto table + picking row from which to collect clearance
        voltage_range = '0 – 5 kV'
        CSA16_clearance = CSA_16.loc['0-5 (1000 where practicable, but in no case less than. a)150 for spans 0-6m b)230 for spans >6 and < 15m c)300 for spans >15m']
        #Creating design buffer array to add onto clearance
        buffer_array = Design_buffer_2_obstacles * float(1000) * np.ones(len(CSA16_clearance))
        CSA16_Design_Clearance = CSA16_clearance + buffer_array
        #Adding text that appears with this in CSA table
        text = ' where practicable, but in no case less than. \n a)150 for spans 0-6m \n b)230 for spans >6 and < 15m \n c)300 for spans >15m'
    elif voltage <= 22:
        #Voltage range for printing onto table + picking row from which to collect clearance
        voltage_range = '5 - 22 kV'
        CSA16_clearance = CSA_16.loc['> 5 < 22 (1000 wherever practicable, but in no case less than 500)']
        #Creating design buffer array to add onto clearance
        buffer_array = Design_buffer_2_obstacles * float(1000) * np.ones(len(CSA16_clearance))
        CSA16_Design_Clearance = CSA16_clearance + buffer_array
        #Adding text that appears with this in CSA table
        text = ' where practicable, but in no case less than 500'
    elif voltage <= 50:
        #Voltage range for printing onto table + picking row from which to collect clearance
        voltage_range = '22 – 50 kV'
        CSA16_clearance = CSA_16.loc['>22 < 50']
        #Creating design buffer array to add onto clearance
        buffer_array = Design_buffer_2_obstacles * float(1000) * np.ones(len(CSA16_clearance))
        CSA16_Design_Clearance = CSA16_clearance + buffer_array
        #Adding text that appears with this in CSA table (Not text in this case but kept in order to satisfy a statement later)
        text = ''
    else: 
        #Voltage range for printing onto table + picking row from which to collect clearance
        voltage_range = '> 50 kV'
        CSA16_clearance = CSA_16.loc['>50 (+10mm/kV over 50kV)']
        #this next row is addin the voltage adder as specified in the CSA table
        CSA16_clearance = CSA16_clearance + np.ones(len(CSA16_clearance)) * (voltage-50) * 10
        #Creating design buffer array to add onto clearance
        buffer_array = Design_buffer_2_obstacles * float(1000) * np.ones(len(CSA16_clearance))
        CSA16_Design_Clearance = CSA16_clearance + buffer_array
        #Adding text that appears with this in CSA table
        text = ''

    #Rounding
    Basic_num = np.char.mod(f'%0.{Numpy_round_integer}f', CSA16_clearance[0])
    Basic_neut = np.char.mod(f'%0.{Numpy_round_integer}f', CSA16_clearance_neut[0])

    DC = np.char.mod(f'%0.{Numpy_round_integer}f', CSA16_Design_Clearance[0])
    neut_DC = np.char.mod(f'%0.{Numpy_round_integer}f', CSA16_clearance_neut_Design_Clearance[0])

    #Adding text to values that require it
    Basic_annotated = str(Basic_num) + text
    Basic_neut_annotated = str(Basic_neut) + text_neut

    DC_annotated = str(DC) + text
    DC_neut_annotated = str(neut_DC) + text_neut

    #Creating the arrays that will be shown in the columns of the spreadsheet
    Basic = np.array([Basic_annotated, Basic_neut_annotated])
    Design_Clearance = np.array([DC_annotated, DC_neut_annotated])

    data = {
        'Basic': Basic,
        'Design_Clearance': Design_Clearance,

        'Voltage range': voltage_range,
        }
    return data
CSA16_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_buffer_2_obstacles', 'Max_Overvoltage']}
CSA_Table16_data = CSA_Table_16_clearance(**CSA16_input)

#The following function is for CSA Table 17
def CSA_Table_17_clearance(p2p_voltage, Design_buffer_2_obstacles, Span_Length, Final_Unloaded_Sag_15C, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_17 = pd.read_excel(ref_table, "CSA table 17")

    #This will look a bit different since we'll be finding data based on rows vs. columns.
    #Since .loc is being used to find the row a column must be chosen to use 
    CSA_17.set_index("Line conductor", inplace = True)


    #Figuring out the correct row name
    if  0 < voltage <= 5 and Span_Length <= 6: 
        #Voltage range and span length range to print on the spreadsheet
        voltage_range = '0 – 5 kV†'
        Span_Range = '0-6 m'
        CSA17_clearance = CSA_17.loc['0-5 kV ac 0 - 6 m'][0]
        #Creating design buffer array to add onto clearance
        CSA17_Design_Clearance = CSA17_clearance + Design_buffer_2_obstacles * float(1000)
        CSA17_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA17_clearance)
        CSA17_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA17_Design_Clearance)

    elif 0 < voltage <= 5 and Span_Length <= 50: 
        #Voltage range and span length range to print on the spreadsheet
        voltage_range = '0 – 5 kV†'
        Span_Range = '> 6 < 50 m'
        CSA17_clearance = CSA_17.loc['0-5 kV ac > 6 < 50 m'][0]
        #Creating design buffer array to add onto clearance
        CSA17_Design_Clearance = CSA17_clearance + Design_buffer_2_obstacles * float(1000)
        CSA17_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA17_clearance)
        CSA17_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA17_Design_Clearance)

    elif 0 < voltage <= 5 and Span_Length > 50 : 
        #Voltage range and span length range to print on the spreadsheet
        voltage_range = '0 – 5 kV†'
        Span_Range = '> 50 < 450 m'
        CSA17_clearance = CSA_17.loc['0-5 kV ac  > 50 < 450 m a )3 x the distance (in m) by which the span length exceeds 50 m; b)83 x the final unloaded sag (in m) at 15C conductor temperature for conductor(s) having the greatest sag; and c)10 mm/kV over 5 kV.'][0]
        CSA17_clearance = CSA17_clearance + (3*(Span_Length - 50)) + (83*(Final_Unloaded_Sag_15C))
        #Creating design buffer array to add onto clearance
        CSA17_Design_Clearance = CSA17_clearance + Design_buffer_2_obstacles * float(1000)
        CSA17_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA17_clearance)
        CSA17_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA17_Design_Clearance)

    elif voltage > 5 and Span_Length <= 50: 
        #Voltage range and span length range to print on the spreadsheet
        voltage_range = '> 5 kV†'
        Span_Range = '< 50 m'
        CSA17_clearance = CSA_17.loc['> 5kV ac < 50 m + 10 mm/kV over 1kV'][0]
        CSA17_clearance = CSA17_clearance + 10 * (voltage - 1)
        #Creating design buffer array to add onto clearance
        CSA17_Design_Clearance = CSA17_clearance + Design_buffer_2_obstacles * float(1000)
        CSA17_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA17_clearance)
        CSA17_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA17_Design_Clearance)

    elif voltage > 5 and Span_Length <= 450: 
        #Voltage range and span length range to print on the spreadsheet
        voltage_range = '> 5 kV†'
        Span_Range = '> 50 < 450 m'
        CSA17_clearance = CSA_17.loc['> 5kV ac > 50 < 450 m a )3 x the distance (in m) by which the span length exceeds 50 m; b)83 x the final unloaded sag (in m) at 15C conductor temperature for conductor(s) having the greatest sag; and c)10 mm/kV over 5 kV.'][0]
        CSA17_clearance = CSA17_clearance + (3*(Span_Length - 50)) + (83*(Final_Unloaded_Sag_15C)) + (10*(voltage-5))
        #Creating design buffer array to add onto clearance
        CSA17_Design_Clearance = CSA17_clearance + Design_buffer_2_obstacles * float(1000)
        CSA17_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA17_clearance)
        CSA17_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA17_Design_Clearance)

    else:
        #Voltage range and span length range to print on the spreadsheet
        voltage_range = '*'
        Span_Range = '> 50 < 450 m'
        CSA17_clearance = 'For spans longer than 450 m, the separation shall be based on best engineering practices, but shall be not less than the separations specified for spans of 450 m.'
        #Creating design buffer array to add onto clearance
        CSA17_Design_Clearance = 'For spans longer than 450 m, the separation shall be based on best engineering practices, but shall be not less than the separations specified for spans of 450 m.'


    #Creating the arrays that will be shown in the columns of the spreadsheet
    data = {
        'Basic': CSA17_clearance,
        'Design_Clearance': CSA17_Design_Clearance,

        'Voltage range': voltage_range,
        'Span Range': Span_Range
        }
    return data
CSA17_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_buffer_2_obstacles', 'Span_Length', 'Final_Unloaded_Sag_15C', 'Max_Overvoltage']}
CSA_Table17_data = CSA_Table_17_clearance(**CSA17_input)

#The following function is for CSA Table 18
def CSA_Table_18_clearance(p2p_voltage, Design_Buffer_Same_Structure, XING_P2P_Voltage, Is_main_wire_upper, Max_Overvoltage, Max_Overvoltage_XING):

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_18 = pd.read_excel(ref_table, "CSA table 18")

    #Index being set so that row can be found using name
    CSA_18.set_index("Conductors at lower level", inplace = True)

    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    Voltage_multiplier_XING = Max_Overvoltage_XING/100 + 1

    if Is_main_wire_upper == True:
        #Getting Phase to ground voltage if main wire is upper wire
        voltage = p2p_voltage / np.sqrt(3)
        voltage = voltage * Voltage_multiplier

        under_voltage = XING_P2P_Voltage / np.sqrt(3)
        under_voltage = under_voltage * Voltage_multiplier_XING 
    else:
        #Getting Phase to ground voltage if main wire is lower wire
        under_voltage = p2p_voltage / np.sqrt(3)
        under_voltage = under_voltage * Voltage_multiplier_XING 

        voltage = XING_P2P_Voltage / np.sqrt(3)
        voltage = voltage * Voltage_multiplier

    #From the way this table works a specfific cell must be selected using row and column the following if else statements will pick the proper row and column

    #Figuring out the correct row name
    if  0 < voltage <= 0.75: 
        voltage_range_over = 'AC 0-750 V'
    elif voltage <=5: 
        voltage_range_over = 'AC > 0.75 ≤ 5'
    elif voltage <=22: 
        voltage_range_over = 'AC > 5 ≤ 22'
    elif voltage <=50: 
        voltage_range_over = 'AC > 22 ≤ 50'
    elif voltage <=90: 
        voltage_range_over = 'AC > 50 ≤ 90'
    else: 
        voltage_range_over = 'Error'

    #Figuring out the correct column name
    if  0 < under_voltage <= 0.75: 
         voltage_range_lower = 'AC 0-750 V'
    elif under_voltage <=5: 
        voltage_range_lower = 'AC > 0.75 ≤ 5'
    elif under_voltage <=22: 
        voltage_range_lower = 'AC > 5 ≤ 22'
    elif under_voltage <=50: 
        voltage_range_lower = 'AC > 22 ≤ 50'
    elif under_voltage <=90: 
        voltage_range_lower = 'AC > 50 ≤ 90'
    else: 
        voltage_range_lower = 'Error'

    #Clearance values
    CSA18_clearance = CSA_18.at[ voltage_range_lower , voltage_range_over ]

    #Creating array for design clearance
    CSA18_Design_Clearance = []

    #Figuring out if the number has any symbols attached to it
    #If there are symbols they will be removed so addition can be done and then added back in
    if isinstance(CSA18_clearance, int):
        CSA18_Design_Clearance = CSA18_clearance + Design_Buffer_Same_Structure
        CSA18_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA18_clearance)
        CSA18_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA18_Design_Clearance)
    elif isinstance(CSA18_clearance, float):
        CSA18_Design_Clearance = CSA18_clearance + Design_Buffer_Same_Structure
        CSA18_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA18_clearance)
        CSA18_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA18_Design_Clearance)
    elif isinstance(CSA18_clearance, str):
        num = float(CSA18_clearance.rstrip('*†‡'))
        num = np.char.mod(f'%0.{Numpy_round_integer}f', num)
        Design_Buffer_Same_Structure = np.char.mod(f'%0.{Numpy_round_integer}f', Design_Buffer_Same_Structure)
        CSA18_Design_Clearance = f"{num + Design_Buffer_Same_Structure}{CSA18_clearance[len(str(int(num))):]}"
        CSA18_clearance = f'{num}{CSA18_clearance[len(str(int(num))):]}'

    #Error message incase the voltage below is highe rthan the one above
    if under_voltage > voltage:
        CSA18_clearance = "Error higher voltage below lower voltage"
        CSA18_Design_Clearance = "Error higher voltage below lower voltage"
    
    if under_voltage or voltage > 90:
        CSA18_clearance = "Error one or both of the voltages entered are too high"
        CSA18_Design_Clearance = "Error one or both of the voltages entered are too high"
        voltage_range_over = "Error one or both of the voltages entered are too high"
        voltage_range_lower = "Error one or both of the voltages entered are too high"

    data = {
        'Basic': CSA18_clearance,
        'Design Clearance': CSA18_Design_Clearance,

        'Voltage range': voltage_range_over,
        'Voltage range under': voltage_range_lower,
        }
    
    return data
CSA18_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_Buffer_Same_Structure', "XING_P2P_Voltage", 'Is_main_wire_upper', 'Max_Overvoltage', 'Max_Overvoltage_XING']}
CSA_Table18_data = CSA_Table_18_clearance(**CSA18_input)

#The following function is for CSA Table 20
def CSA_Table_20_clearance(p2p_voltage, Design_Buffer_Same_Structure, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_20 = pd.read_excel(ref_table, "CSA table 20")

    #This will look a bit different since we'll be finding data based on rows vs. columns.
    #Since .loc is being used to find the row a column must be chosen to use 
    CSA_20.set_index("Maximum circuit line-to-ground voltage, kV", inplace = True)

    #Figuring out the correct row name
    if  0 < voltage <= 0.75: 
        CSA20_clearance = CSA_20.loc['0.75 kV']
        voltage_range = '0.75'
    elif voltage <= 5: 
        CSA20_clearance = CSA_20.loc['5 kV']
        voltage_range = '5'
    elif voltage <= 10: 
        CSA20_clearance = CSA_20.loc['10 kV']
        voltage_range = '10'
    elif voltage <= 15:  
        CSA20_clearance = CSA_20.loc['15 kV']
        voltage_range = '15'
    elif voltage <= 22: 
        CSA20_clearance = CSA_20.loc['22 kV']
        voltage_range = '22'
    elif voltage <= 50:  
        CSA20_clearance = CSA_20.loc['50 kV']
        voltage_range = '50'
    else:
        CSA20_clearance = CSA_20.loc['90 kV']
        voltage_range = '90'

    CSA20_Design_Clearance = CSA20_clearance + np.ones(len(CSA20_clearance)) * Design_Buffer_Same_Structure * 1000
    #Creating the arrays that will be shown in the columns of the spreadsheet

    #Adding rounding
    Basic = np.char.mod(f'%0.{Numpy_round_integer}f', CSA20_clearance[0])
    Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA20_Design_Clearance[0])
    Between_Conductors = np.char.mod(f'%0.{Numpy_round_integer}f', CSA20_clearance[1])
    Between_Conductors_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA20_Design_Clearance[1])

    data = {
        'Basic': Basic,
        'Design_Clearance': Design_Clearance,

        'Between conductors': Between_Conductors,
        'Between conductors Design_Clearance': Between_Conductors_Design_Clearance,

        'Voltage range': voltage_range
        }
    return data
CSA20_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_Buffer_Same_Structure', 'Max_Overvoltage']}
CSA_Table20_data = CSA_Table_20_clearance(**CSA20_input)

#The following function is for CSA Table 21
def CSA_Table_21_clearance(p2p_voltage, Design_Buffer_Same_Structure, XING_P2P_Voltage, Is_main_wire_upper, Max_Overvoltage, Max_Overvoltage_XING):

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_21 = pd.read_excel(ref_table, "CSA table 21")

    #Index being set so that row can be found using name
    CSA_21.set_index("Maximum lower conductor line-to-ground voltage, kV", inplace = True)

    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    Voltage_multiplier_XING = Max_Overvoltage_XING/100 + 1

    if Is_main_wire_upper == True:
        #Getting Phase to ground voltage if main wire is upper wire
        voltage = p2p_voltage / np.sqrt(3)
        voltage = voltage * Voltage_multiplier

        under_voltage = XING_P2P_Voltage / np.sqrt(3)
        under_voltage = under_voltage * Voltage_multiplier_XING 
    else:
        #Getting Phase to ground voltage if main wire is lower wire
        under_voltage = p2p_voltage / np.sqrt(3)
        under_voltage = under_voltage * Voltage_multiplier_XING 

        voltage = XING_P2P_Voltage / np.sqrt(3)
        voltage = voltage * Voltage_multiplier
    
    #From the way this table works a specfific cell must be selected using row and column the following if else statements will pick the proper row and column

    #Figuring out the correct row name
    if  0 < voltage <= 5: 
        voltage_range_over = '5 kv'
    elif voltage <=10: 
        voltage_range_over = '10 kv'
    elif voltage <=17: 
        voltage_range_over = '17 kv'
    elif voltage <=22: 
        voltage_range_over = '22 kv'
    elif voltage <=30: 
        voltage_range_over = '30 kv'
    elif voltage <=50: 
        voltage_range_over = '50 kv'
    else: 
        voltage_range_over = '90 kv'

    #Figuring out the correct column name
    if  0 < under_voltage <= 0.75: 
         voltage_range_lower = '0.75 kv'
    elif under_voltage <=5: 
        voltage_range_lower = '5 kv'
    elif under_voltage <=10: 
        voltage_range_lower = '10 kv'
    elif under_voltage <=17: 
        voltage_range_lower = '17 kv'
    elif under_voltage <=22: 
        voltage_range_lower = '22 kv'
    elif under_voltage <=30: 
        voltage_range_lower = '30 kv'
    elif under_voltage <=50: 
        voltage_range_lower = '50 kv'
    else: 
        voltage_range_lower = '90 kv'

    #Clearance values
    CSA21_clearance = CSA_21.at[ voltage_range_lower , voltage_range_over ]
    CSA21_Design_Clearance = float(CSA21_clearance) + Design_Buffer_Same_Structure

    #Rounding
    CSA21_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA21_clearance)
    CSA21_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA21_Design_Clearance)

    #Error message incase the voltage below is highe rthan the one above
    if voltage_range_lower > voltage_range_over:
        CSA21_clearance = "Error higher voltage below lower voltage"
        CSA21_Design_Clearance = "Error higher voltage below lower voltage"
    
    if voltage > 90:
        CSA21_clearance = "Out of Range"
        CSA21_Design_Clearance = "Out of Range"
        voltage_range_over = "Out of Range"
        voltage_range_lower = "Out of Range"

    data = {
        'Basic': CSA21_clearance,
        'Design Clearance': CSA21_Design_Clearance,

        'Voltage range': voltage_range_over,
        'Voltage range under': voltage_range_lower,
        }
    
    return data
CSA21_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_Buffer_Same_Structure', "XING_P2P_Voltage", 'Is_main_wire_upper', 'Max_Overvoltage', 'Max_Overvoltage_XING']}
CSA_Table21_data = CSA_Table_21_clearance(**CSA21_input)

#The following function is for CSA Table 22
def CSA_Table_22_clearance(p2p_voltage, Design_Buffer_Same_Structure, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_22 = pd.read_excel(ref_table, "CSA table 22")

    #Figuring out the correct row name
    if  0 < voltage <= 0.75: 
        CSA22_clearance = CSA_22['0-750 V']
        voltage_range = '0 - 750 V'
    elif voltage <= 5: 
        CSA22_clearance = CSA_22['> 0.75 kV < 5 kV']
        voltage_range = '> 0.75 ≤ 5 kV'
    elif voltage <= 10: 
        CSA22_clearance = CSA_22['> 5 < 22 kV']
        voltage_range = '> 5 ≤ 22 kV'
    elif voltage <= 10: 
        CSA22_clearance = CSA_22['> 22 < 50 kV']
        voltage_range = '> 22 ≤ 50 kV'
    else:
        CSA22_clearance = CSA_22['> 50 kV']
        voltage_range = '> 50 kV'

    #Defining an array that will contain the design clearance
    CSA22_clearance_rounded = []

    #For loop round values in CSA_22 clearance
    for item in CSA22_clearance:
        #First checks if value is an integer and if so adds the clearance
        if isinstance(item, int):
            val = np.char.mod(f'%0.{Numpy_round_integer}f', item)
            CSA22_clearance_rounded.append(val)
        #Then checks if value is a float and if so adds the clearance
        elif isinstance(item, float):
            val = np.char.mod(f'%0.{Numpy_round_integer}f', item)
            CSA22_clearance_rounded.append(val)
        #If the value is a string it will search trough the string for integers or floats
        elif isinstance(item, str):
            match = re.search(r'\d+(\.\d+)?', item)
            #if a number is found in the string it will take the number out and add onto that number and then return a new number
            if match:
                num = float(match.group())
                new_num = num
                new_num = np.char.mod(f'%0.{Numpy_round_integer}f', new_num)
                new_item = item[:match.start()] + str(new_num) + item[match.end():]
                #This new number is then appended to the array
                CSA22_clearance_rounded.append(new_item)
            #If no number is found the original string passes through into the new array
            else:
                CSA22_clearance_rounded.append(item)
        #If there is something that is not a string, float or integer it will pass into the new array unchanged
        else:
            CSA22_clearance_rounded.append(item)

    #Defining an array that will contain the design clearance
    CSA_22_Design_Clearance = []

    #For loop to add clearance onto Table 22 values regardless of if there are symbols on the number
    for item in CSA22_clearance:
        #First checks if value is an integer and if so adds the clearance
        if isinstance(item, int):
            val = item + Design_Buffer_Same_Structure
            val = np.char.mod(f'%0.{Numpy_round_integer}f', val)
            CSA_22_Design_Clearance.append(val)
        #Then checks if value is a float and if so adds the clearance
        elif isinstance(item, float):
            val = item + Design_Buffer_Same_Structure
            val = np.char.mod(f'%0.{Numpy_round_integer}f', val)
            CSA_22_Design_Clearance.append(item + val)
        #If the value is a string it will search trough the string for integers or floats
        elif isinstance(item, str):
            match = re.search(r'\d+(\.\d+)?', item)
            #if a number is found in the string it will take the number out and add onto that number and then return a new number
            if match:
                num = float(match.group())
                new_num = num + Design_Buffer_Same_Structure
                new_num = np.char.mod(f'%0.{Numpy_round_integer}f', new_num)
                new_item = item[:match.start()] + str(new_num) + item[match.end():]
                #This new number is then appended to the array
                CSA_22_Design_Clearance.append(new_item)
            #If no number is found the original string passes through into the new array
            else:
                CSA_22_Design_Clearance.append(item)
        #If there is something that is not a string, float or integer it will pass into the new array unchanged
        else:
            CSA_22_Design_Clearance.append(item)

    CSA_22_Design_Clearance = np.array(CSA_22_Design_Clearance)

    data = {
        'Basic': CSA22_clearance_rounded,
        'Design_Clearance': CSA_22_Design_Clearance,

        'Voltage range': voltage_range
        }
    return data
CSA22_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_Buffer_Same_Structure', 'Max_Overvoltage']}
CSA_Table22_data = CSA_Table_22_clearance(**CSA22_input)

#The following function is for CSA Table 23
def CSA_Table_23_clearance(p2p_voltage, Design_Buffer_Same_Structure, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_23 = pd.read_excel(ref_table, "CSA table 23")

    #Figuring out the correct row name
    if  0 < voltage <= 0.75: 
        CSA23_clearance = CSA_23['0-750 V']
        voltage_range = '0 - 750 V'
        voltage_adder = 0
    elif voltage <= 22: 
        CSA23_clearance = CSA_23['> 0.75 < 22 kV']
        voltage_range = '> 0.75 ≤ 5 kV'
        voltage_adder = 0
    else:
        CSA23_clearance = CSA_23['> 22 < 50 kV*']
        if voltage <= 50:
            voltage_range = '> 22 ≤ 50 kV'
            voltage_adder = 0
        else:
            voltage_adder = (voltage-50) * 0.010
            voltage_range = '> 50 kV'
            voltage_adder = np.round(voltage_adder.astype(float), Numpy_round_integer)

    #Defining an array that will contain the basic array with voltage adder
    CSA23_clearance_basic = []

    #For loop to add the voltage adder
    for item in CSA23_clearance:
        #First checks if value is an integer and if so adds the clearance
        if isinstance(item, int):
            val = item + voltage_adder
            val = np.char.mod(f'%0.{Numpy_round_integer}f', val)
            CSA23_clearance_basic.append(val)
        #Then checks if value is a float and if so adds the clearance
        elif isinstance(item, float):
            val = item + voltage_adder
            val = np.char.mod(f'%0.{Numpy_round_integer}f', val)
            CSA23_clearance_basic.append(val)
        #If the value is a string it will search trough the string for integers or floats
        elif isinstance(item, str):
            match = re.search(r'\d+(\.\d+)?', item)
            #if a number is found in the string it will take the number out and add onto that number and then return a new number
            if match:
                num = float(match.group())
                new_num = num + voltage_adder
                new_num = np.char.mod(f'%0.{Numpy_round_integer}f', new_num)
                new_item = item[:match.start()] + str(new_num) + item[match.end():]
                #This new number is then appended to the array
                CSA23_clearance_basic.append(new_item)
            #If no number is found the original string passes through into the new array
            else:
                CSA23_clearance_basic.append(item)
        #If there is something that is not a string, float or integer it will pass into the new array unchanged
        else:
            CSA23_clearance_basic.append(item)

    CSA23_clearance_basic = np.array(CSA23_clearance_basic)

    #Defining an array that will contain the design clearance
    CSA_23_Design_Clearance = []


    #For loop to add clearance onto values regardless of if there are symbols on the number
    for item in CSA23_clearance_basic:
        #First checks if value is an integer and if so adds the clearance
        if isinstance(item, int):
            val = item + Design_Buffer_Same_Structure
            val = np.char.mod(f'%0.{Numpy_round_integer}f', val)
            CSA_23_Design_Clearance.append(val)
        #Then checks if value is a float and if so adds the clearance
        elif isinstance(item, float):
            val = item + Design_Buffer_Same_Structure
            val = np.char.mod(f'%0.{Numpy_round_integer}f', val)
            CSA_23_Design_Clearance.append(val)
        #If the value is a string it will search trough the string for integers or floats
        elif isinstance(item, str):
            match = re.search(r'\d+(\.\d+)?', item)
            #if a number is found in the string it will take the number out and add onto that number and then return a new number
            if match:
                num = float(match.group())
                new_num = num + Design_Buffer_Same_Structure
                new_num = np.char.mod(f'%0.{Numpy_round_integer}f', new_num)
                new_item = item[:match.start()] + str(new_num) + item[match.end():]
                #This new number is then appended to the array
                CSA_23_Design_Clearance.append(new_item)
            #If no number is found the original string passes through into the new array
            else:
                CSA_23_Design_Clearance.append(item)
        #If there is something that is not a string, float or integer it will pass into the new array unchanged
        else:
            CSA_23_Design_Clearance.append(item)

    CSA_23_Design_Clearance = np.array(CSA_23_Design_Clearance)


    data = {
        'Basic': CSA23_clearance_basic,
        'Design_Clearance': CSA_23_Design_Clearance,

        'Voltage range': voltage_range
        }
    return data
CSA23_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_Buffer_Same_Structure', 'Max_Overvoltage']}
CSA_Table23_data = CSA_Table_23_clearance(**CSA23_input)

#The following function is for CSA Table 24
def CSA_Table_24_clearance(p2p_voltage, Design_Buffer_Same_Structure, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_24 = pd.read_excel(ref_table, "CSA table 24")

    #Index being set so that row can be found using name
    CSA_24.set_index("Voltage of supply conductor", inplace = True)

    #Figuring out the correct row name
    if  0 < voltage <= 0.75: 
        CSA24_clearance = CSA_24.loc['0—750 V with other covering or bare'][0]
        voltage_range = '0 - 750 V'
    elif voltage <= 22: 
        CSA24_clearance = CSA_24.loc['> 0.75 kV and ≤ 15 kV'][0]
        voltage_range = '> 0.75 kV and ≤ 15 kV'
    elif voltage <= 22: 
        CSA24_clearance = CSA_24.loc['> 15 kVand ≤ 22 kV'][0]
        voltage_range = '> 15 kV and ≤ 22 kV'
    else:
        CSA24_clearance = CSA_24.loc['> 22 kV and ≤ 250 kV (+10mm/kV over 22kV)'][0]
        CSA24_clearance = CSA24_clearance + (10 * (voltage-22))
        CSA24_clearance = np.round(CSA24_clearance.astype(float), Numpy_round_integer)
        voltage_range = '> 22 kV and ≤ 250 kV'

    #Conversion of design buffer to mm
    Design_Buffer_Same_Structure = Design_Buffer_Same_Structure *1000

    #Figuring out if the number has any symbols attached to it
    #If there are symbols they will be removed so addition can be done and then added back in
    if isinstance(CSA24_clearance, int):
        CSA_24_Design_Clearance = CSA24_clearance + Design_Buffer_Same_Structure
    elif isinstance(CSA24_clearance, float):
        CSA_24_Design_Clearance = CSA24_clearance + Design_Buffer_Same_Structure
    elif isinstance(CSA_24_Design_Clearance, str):
        num = float(CSA_24_Design_Clearance.rstrip('*†‡'))
        CSA_24_Design_Clearance = f"{num + Design_Buffer_Same_Structure}{CSA24_clearance[len(str(int(num))):]}"

    CSA24_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA24_clearance)
    CSA_24_Design_Clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_24_Design_Clearance)

    if voltage > 250:
        CSA24_clearance = "Out of Range"
        CSA_24_Design_Clearance = "Out of Range"
        voltage_range = "Out of Range"

    data = {
        'Basic': CSA24_clearance,
        'Design_Clearance': CSA_24_Design_Clearance,

        'Voltage range': voltage_range
        }
    return data
CSA24_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_Buffer_Same_Structure', 'Max_Overvoltage']}
CSA_Table24_data = CSA_Table_24_clearance(**CSA24_input)

#The following function is for CSA Table 25
def CSA_Table_25_clearance(p2p_voltage, Design_Buffer_Same_Structure, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_25 = pd.read_excel(ref_table, "CSA table 25")

    #Index being set so that row can be found using name
    CSA_25.set_index("Type of plant near which the guy passes", inplace = True)

    CSA_25_comms = CSA_25.loc['Communication line plant'][0]
    CSA_25_guy = CSA_25.loc['Supply guy wires and span wires'][0]

    #Figuring out the correct row name
    if  0 < voltage <= 0.75: 
        CSA25_clearance = CSA_25.loc['0 - 750 v'][0]
        voltage_range = '0 - 750 V'
    elif voltage <= 22: 
        CSA25_clearance = CSA_25.loc['> 0.75 kV and ≤ 22 kV'][0]
        voltage_range = '> 15 kV and ≤ 22 kV'
    else:
        CSA25_clearance = CSA_25.loc['> 22kV (+0.01 m/kV over 22kV)'][0]
        CSA25_clearance = CSA25_clearance + (0.01 * (voltage-22))
        CSA25_clearance = np.round(CSA25_clearance.astype(float), Numpy_round_integer)
        voltage_range = '> 22 kV'

    #Getting Basic design clearance array
    Basic_clearances = np.array([CSA_25_comms, CSA25_clearance, CSA_25_guy])
    Design_clearances = Basic_clearances + (np.ones(len(Basic_clearances)) * Design_Buffer_Same_Structure)

    #Making an array and a for loop for rounding CSA table 25 DC values
    CSA_25_Design_Clearance = []

    for item in Design_clearances:
            val = np.char.mod(f'%0.{Numpy_round_integer}f', item)
            CSA_25_Design_Clearance.append(val)

    #Now redefining the basic clearances (just rounded)
    CSA_25_comms = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_25_comms)
    CSA25_clearance = np.char.mod(f'%0.{Numpy_round_integer}f', CSA25_clearance)
    CSA_25_guy = np.char.mod(f'%0.{Numpy_round_integer}f', CSA_25_guy)
    Basic_clearances = np.array([CSA_25_comms, CSA25_clearance, CSA_25_guy])

    data = {
        'Basic': Basic_clearances,
        'Design_Clearance': CSA_25_Design_Clearance,


        'Voltage range': voltage_range
        }
    return data
CSA25_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_Buffer_Same_Structure', 'Max_Overvoltage']}
CSA_Table25_data = CSA_Table_25_clearance(**CSA25_input)

#The following function is for CSA Table 26
def CSA_Table_26_clearance(p2p_voltage, Design_Buffer_Same_Structure, Max_Overvoltage):

    #Getting Phase to ground voltage
    voltage = p2p_voltage / np.sqrt(3)
    #Adding Voltage Multiplier
    Voltage_multiplier = Max_Overvoltage/100 + 1
    voltage = voltage * Voltage_multiplier

    #Start by opening the sheet for CSA Table 3 within the reference table file
    CSA_26 = pd.read_excel(ref_table, "CSA table 26")

    #Index being set so that row can be found using name
    CSA_26.set_index("Type of plant over or near which the guy passes", inplace = True)

    CSA_26_comms = CSA_26.loc['Communication line plant']

    #Figuring out the correct row name
    if  0 < voltage <= 0.75: 
        CSA26_clearance = CSA_26.loc['0-750 v']
        voltage_range = '0 - 750 V'
    elif voltage <= 5: 
        CSA26_clearance = CSA_26.loc['> 0.75 kV and ≤ 5 kV']
        voltage_range = '> 0.75 kV and ≤ 5 kV'
    elif voltage <= 15: 
        CSA26_clearance = CSA_26.loc['> 5kV and ≤ 15 kV']
        voltage_range = '> 5kV and ≤ 15 kV'
    elif voltage <= 22: 
        CSA26_clearance = CSA_26.loc['> 15 kV and ≤ 22 kV']
        voltage_range = '> 15 kV and ≤ 22 kV'
    else:
        CSA26_clearance = CSA_26.loc['> 22 kV (+10mm/kV over 22kV)']
        CSA26_clearance = CSA26_clearance + np.ones(len(CSA26_clearance))*(10 * (voltage-22))
        CSA26_clearance = np.round(CSA26_clearance.astype(float), Numpy_round_integer)
        voltage_range = '> 22 kV'

    #Defining an array that will contain the design clearance
    CSA_26_Design_Clearance = []
    CSA_26_Design_Clearance_Comms = []

    #For loop to add clearance onto comms
    for item in CSA_26_comms:
        #First checks if value is an integer and if so adds the clearance
        if isinstance(item, int):
            val = item + Design_Buffer_Same_Structure
            val = np.char.mod(f'%0.{Numpy_round_integer}f', val)
            CSA_26_Design_Clearance_Comms.append(val)
        #Then checks if value is a float and if so adds the clearance
        elif isinstance(item, float):
            val = item + Design_Buffer_Same_Structure
            val = np.char.mod(f'%0.{Numpy_round_integer}f', val)
            CSA_26_Design_Clearance_Comms.append(val)
        #If the value is a string it will search trough the string for integers or floats
        elif isinstance(item, str):
            match = re.search(r'\d+(\.\d+)?', item)
            #if a number is found in the string it will take the number out and add onto that number and then return a new number
            if match:
                num = float(match.group())
                new_num = num + Design_Buffer_Same_Structure
                new_num = np.char.mod(f'%0.{Numpy_round_integer}f', new_num)
                new_item = item[:match.start()] + str(new_num) + item[match.end():]
                #This new number is then appended to the array
                CSA_26_Design_Clearance_Comms.append(new_item)
            #If no number is found the original string passes through into the new array
            else:
                CSA_26_Design_Clearance_Comms.append(item)
        #If there is something that is not a string, float or integer it will pass into the new array unchanged
        else:
            CSA_26_Design_Clearance_Comms.append(item)

    CSA_26_Design_Clearance_Comms = np.array(CSA_26_Design_Clearance_Comms)

    #For loop to add clearance onto conductors
    for item in CSA26_clearance:
        #First checks if value is an integer and if so adds the clearance
        if isinstance(item, int):
            val = item + Design_Buffer_Same_Structure
            val = np.char.mod(f'%0.{Numpy_round_integer}f', val)
            CSA_26_Design_Clearance.append(val)
        #Then checks if value is a float and if so adds the clearance
        elif isinstance(item, float):
            val = item + Design_Buffer_Same_Structure
            val = np.char.mod(f'%0.{Numpy_round_integer}f', val)
            CSA_26_Design_Clearance.append(val)
        #If the value is a string it will search trough the string for integers or floats
        elif isinstance(item, str):
            match = re.search(r'\d+(\.\d+)?', item)
            #if a number is found in the string it will take the number out and add onto that number and then return a new number
            if match:
                num = float(match.group())
                new_num = num + Design_Buffer_Same_Structure
                new_num = np.char.mod(f'%0.{Numpy_round_integer}f', new_num)
                new_item = item[:match.start()] + str(new_num) + item[match.end():]
                #This new number is then appended to the array
                CSA_26_Design_Clearance.append(new_item)
            #If no number is found the original string passes through into the new array
            else:
                CSA_26_Design_Clearance.append(item)
        #If there is something that is not a string, float or integer it will pass into the new array unchanged
        else:
            CSA_26_Design_Clearance.append(item)

    CSA_26_Design_Clearance = np.array(CSA_26_Design_Clearance)

    CSA_26_comms_clearance = np.array([CSA_26_comms[0], CSA_26_Design_Clearance_Comms[0], CSA_26_comms[1], CSA_26_Design_Clearance_Comms[1]])
    CSA_26_conductor_clearance = np.array([CSA26_clearance[0], CSA_26_Design_Clearance[0], CSA26_clearance[1], CSA_26_Design_Clearance[1]])

    data = {
        'Comms': CSA_26_comms_clearance,
        'Conductors': CSA_26_conductor_clearance,

        'Voltage range': voltage_range
        }
    return data
CSA26_input = {key: inputs[key] for key in ['p2p_voltage', 'Design_Buffer_Same_Structure', 'Max_Overvoltage']}
CSA_Table26_data = CSA_Table_26_clearance(**CSA26_input)

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
    AEUC5_cell00 = 'Over walkways or land normally accessible only to pedestrians, snowmobiles, and all terrain vehicles not exceeding 3.6m  ‡'
    AEUC5_cell10 = 'Over rights of way of underground pipelines operating at a pressure of over 700 kilopascals; equipment not exceeding 4.15m'
    AEUC5_cell20 = 'Over land likely to be travelled by road vehicles (including roadways, streets, lanes, alleys, driveways, and entrances); equipment not exceeding 4.15m **'
    AEUC5_cell30 = 'Over land likely to be travelled by road vehicles (including highways, roadways, streets, lanes, alleys, driveways, and entrances); equipment not exceeding 5.3m †'
    AEUC5_cell40 = 'Over land likely to be travelled by agricultural or other equipment; equipment not exceeding 5.3m ††'
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
        ['Location of Wires or Conductors *', 'Guys, Messengers, Span & Lightening Protection Wires and Communications Wires and Cables', ' ', ' ', ' ', ' ', 'Voltage of Open Supply Conductors and Service Conductors Voltage Line to Ground kV AC except where note (Voltages in Parentheses are AC Phase to Phase) ' + str(voltage_range)],
        [' '],
        [' ', 'Col. I', ' ', ' ', ' ', ' ' , col],
        [' ', 'Basic (m)', 'Re-pave Adder (m)', "Snow Adder (m)", "AEUC Total (m)", "Design Clearance (m)", 'Basic (m)', "Altitude Adder (m)", 'Re-pave Adder (m)', "Snow Adder (m)", "AEUC Total (m)", "Design Clearance (m)"],
              ]
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(6):
        row = [AEUC5_titles[i], AEUC_Table5_data['AEUC base neutral'][i], AEUC_Table5_data['Re-pave Adder'][i], AEUC_Table5_data['Snow Adder'][i], AEUC_Table5_data['AEUC total neutral'][i], AEUC_Table5_data['Design clearance neutral'][i], \
               AEUC_Table5_data['AEUC base clearance'][i], AEUC_Table5_data['Altitude Adder'][i], AEUC_Table5_data['Cond Re-pave Adder'][i], AEUC_Table5_data['Cond Snow Adder'][i], AEUC_Table5_data['AEUC total'][i], AEUC_Table5_data['Design clearance'][i]]
        AEUC_5.append(row)

    #This is retrieving the number of rows in each table array
    n_row_AEUC_table_5 = len(AEUC_5)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_AEUC_table_5 = len(AEUC_5[7])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_AEUC_5 = []
    for i in range(n_row_AEUC_table_5):
        if i < 7:
            list_range_color_AEUC_5.append((i + 1, 1, i + 1, n_column_AEUC_table_5, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_AEUC_5.append((i + 1, 1, i + 1, n_column_AEUC_table_5, color_bkg_data_1))
        else:
            list_range_color_AEUC_5.append((i + 1, 1, i + 1, n_column_AEUC_table_5, color_bkg_data_2))

    # define cell format
    cell_format_AEUC_5 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_AEUC_table_5, 'center'), (4, 2, 5, 6, 'center'), (4, 1, 7, 1, 'center'), (4, 7, 5, 12, 'center'), (6, 2, 6, 6, 'center'), (6, 7, 6, 12, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_AEUC_5,
        'range_border': [(1, 1, 1, n_column_AEUC_table_5), (2, 1, n_row_AEUC_table_5, n_column_AEUC_table_5)],
        'row_height': [(1, 50)]+[(4, 20)]+[(5, 20)],
        'column_width': [(1, 30)] + [(i + 2, 10) for i in range(n_column_AEUC_table_5)],
    }

    # define some footer notes
    footer_AEUC5 = ['*Where a line runs parallel to land accessible to vehicles but is over land not requiring clearance for vehicles, the wire can swing out over the area accessible to vehicles or, at voltages over 200 kV AC, vehicles can be subjected to a hazard from induced voltages.' ,
                    'These vertical clearances apply where the conductor (in the swing condition, where specified) is over, or within the following horizontal distances from the edge of, land accessible to vehicles:', 
                    '(a) 0.0 m for communication circuits and O to 50 kV phase to phase AC supply circuits;' ,
                    '(b) 0.9 m for 50 to 90 kV phase to phase AC supply circuits;' , 
                    '(c) 1.7 m for 120 to 150 kV phase to phase AC supply circuits;',
                    '(d) 6.1 mfor 250 to 350 kV phase to phase AC supply circuits;', 
                    '**Generally Restricted to Urban Areas',
                    '†Provincial and municipal authorities may designate certain roads and highways as high load corridors and set specific ground clearances for these routes.',
                    '††This category includes farm fields and access roads to farm fields, as well as entrances to farm yards.',
                    "‡For voltages from 0-750V this clearance can be reduced to 3.5 m in the last span connecting the overhead supply to the consumer's service point of attachment.",
                    'Note: See high load corridors on map of Provincial Highways for vehicle hights of 9.0m and 12.8m.',
                    'AESO rules section 502.2 clause 17 (3) requires a minimum clearance of 12.2 m over agricultural land.',
                    'Basic clearances from AEUC code 2022 table 5.']

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
        if i < 6:
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
        'column_width': [(1, 40)] + [(i + 1, 10) for i in range(1, n_column_AEUC_table_7)],
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
        [' '],
        [' ', 'Guys, messengers, communication, span & lightning protection wires; communication cables', ' ', ' ', ' ', ' ', 'Open Supply Conductors and Service Conductors, ac' + str(voltage_range) + ' kV'],
        [' '],
        [' ', 'Col. II', ' ', ' ', ' ', ' ' , col],
        [' ', 'Basic (m)', 'Re-pave Adder (m)', "Snow Adder (m)", "CSA Total (m)", "Design Clearance (m)", 'Basic (m)', "Altitude Adder (m)", 'Re-pave Adder (m)', "Snow Adder (m)", "CSA Total (m)", "Design Clearance (m)"],
              ]
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(6):
        row = [CSA2_titles[i], CSA_Table2_data['Neutral Basic (m)'][i], CSA_Table2_data['Re-pave Adder (m)'][i], CSA_Table2_data['Snow Adder (m)'][i], CSA_Table2_data['Neutral CSA Total (m)'][i], CSA_Table2_data['Neutral Design Clearance (m)'][i], \
               CSA_Table2_data['Basic (m)'][i], CSA_Table2_data['Altitude Adder (m)'][i], CSA_Table2_data['Cond Re-pave Adder (m)'][i], CSA_Table2_data['Cond Snow Adder (m)'][i], CSA_Table2_data['CSA Total (m)'][i], CSA_Table2_data['Design Clearance (m)'][i]]
        CSA2.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_2 = len(CSA2)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_2 = len(CSA2[9])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA2 = []
    for i in range(n_row_CSA_table_2):
        if i < 8:
            list_range_color_CSA2.append((i + 1, 1, i + 1, n_column_CSA_table_2, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA2.append((i + 1, 1, i + 1, n_column_CSA_table_2, color_bkg_data_1))
        else:
            list_range_color_CSA2.append((i + 1, 1, i + 1, n_column_CSA_table_2, color_bkg_data_2))

    # define cell format
    cell_format_CSA_2 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 4, n_column_CSA_table_2, 'center'), (5, 2, 6, 6, 'center'), (5, 1, 8, 1, 'center'), (5, 7, 6, 12, 'center'), (7, 2, 7, 6, 'center'), (7, 7, 7, 12, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA2,
        'range_border': [(1, 1, 1, n_column_CSA_table_2), (2, 1, n_row_CSA_table_2, n_column_CSA_table_2)],
        'row_height': [(1, 30)],
        'column_width': [(1, 50)] + [(i + 2, 10) for i in range(n_column_CSA_table_2)],
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
        if i < 10:
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
        'column_width':[(1, 10)] + [(2, 50)] + [(i + 3, 10) for i in range(n_column_CSA_table_3)],
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
        if i < 6:
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
        if i < 6:
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
        if i < 5:
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

#CSA Table 9
#region
    #Creating the title blocks. etc
    CSA9_cell00 = 'Guys, communication cables, and drop wires'
    CSA9_cell10 = '0-750 V'
    CSA9_cell20 = ' '
    CSA9_cell30 = ' '
    CSA9_cell40 = '> 0.75kV <=22 kV'
    CSA9_cell50 = ' '
    CSA9_cell60 = '> 22 kV**††'

    CSA9_cell01 = ' '
    CSA9_cell11 = 'Insulated or grounded'
    CSA9_cell21 = 'Enclosed in effectively grounded metallic sheath'
    CSA9_cell31 = 'Not insulated, grounded, or enclosed in effectively grounded metallic sheath'
    CSA9_cell41 = 'Enclosed in effectively grounded metallic sheath'
    CSA9_cell51 = 'Not enclosed in effectively grounded metallic sheath'
    CSA9_cell61 = '-'

    voltage = inputs['p2p_voltage']

    CSA9_titles = np.array([CSA9_cell00, CSA9_cell10, CSA9_cell20, CSA9_cell30, CSA9_cell40, CSA9_cell50, CSA9_cell60])
    CSA9_loc_titles = np.array([CSA9_cell01, CSA9_cell11, CSA9_cell21, CSA9_cell31, CSA9_cell41, CSA9_cell51, CSA9_cell61])
    
    CSA9 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 9 \n Minimum Design Clearances from Wires and Conductors not attached to Buildings, Signs, and similar Plant, ac \n (See clauses 5.7.3.1 to 5.7.3.3.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        [' ', ' ', 'Minimum clearance, m'],
        [' ', ' ', 'To Buildings*† and above ground pipelines', ' ', ' ', ' ', 'To signs, billboards, lamp and traffic sign standards, and similar plant', ' ', ' ', ' '],
        [' ', ' ', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance'],
        [' ', ' ', 'Horizontal to surface', ' ', 'Vertical to surface', ' ', 'Horizontal to surface', ' ', 'Vertical to surface', ' '],
              ]
    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(7):
        row = [CSA9_titles[i], CSA9_loc_titles[i], CSA_Table9_data['Buildings basic horizontal'][i], CSA_Table9_data['Buildings Design Clearance Horizontal'][i], CSA_Table9_data['Buildings basic vertical'][i], CSA_Table9_data['Buildings Design Clearance vertical'][i] \
               , CSA_Table9_data['Obstacles basic horizontal'][i], CSA_Table9_data['Obstacles Design Clearance Horizontal'][i], CSA_Table9_data['Obstacles basic vertical'][i], CSA_Table9_data['Obstacles Design Clearance vertical'][i]]
        CSA9.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_9 = len(CSA9)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_9 = len(CSA9[7])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA9 = []
    for i in range(n_row_CSA_table_9):
        if i < 7:
            list_range_color_CSA9.append((i + 1, 1, i + 1, n_column_CSA_table_9, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA9.append((i + 1, 1, i + 1, n_column_CSA_table_9, color_bkg_data_1))
        else:
            list_range_color_CSA9.append((i + 1, 1, i + 1, n_column_CSA_table_9, color_bkg_data_2))

    # define cell format
    cell_format_CSA_9 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_9, 'center'),(4, 1, 7, 2, 'center'), (4, 3, 4, 10, 'center'), (5, 3, 5, 6, 'center'), (5, 7, 5, 10, 'center'), (7, 3, 7, 4, 'center'), (7, 5, 7, 6, 'center'), (7, 7, 7, 8, 'center'), (7, 9, 7, 10, 'center'),\
                        (8, 1, 8, 2, 'center'),(9, 1, 11, 1, 'center'),(12, 1, 13, 1, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA9,
        'range_border': [(1, 1, 1, n_column_CSA_table_9), (2, 1, n_row_CSA_table_9, n_column_CSA_table_9)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_9)],
    }

    # define some footer notes
    footer_CSA9 = ['* Clearances over or adjacent to portions of a building normally traversed by persons or vehicles are specified in Tables 2 and 4.',
                   '† Clearances are applicable to non-metallic buildings or buildings whose metallic parts are effectively grounded. For other buildings, an assessment might be needed to determine additional clearances for electrostatic induction (see Clause 5.7.3.3).', 
                   '0-750V 1.Insulated or grounded  2.Not insulated, grounded, or enclosed in effectively grounded metallic sheath, Vertical to Surface may be reduced to 1 m for portions of the building considered normally inaccessible.',
                   '** and 0.75 - 22kV,Not enclosed in effectively grounded metallic sheath, Vertical to Surface. Conductors of these voltage classes should not be carried over buildings where other suitable construction can be used.',
                   '†† Where it is necessary to carry conductors of these voltage classes over buildings, it should be determined whether additional measures, including increased clearances, are needed to ensure that the crossed-over buildings can be used safely and effectively.',
                   '0.75 - 22kV, Not enclosed in effectively grounded metallic sheath, Horizontal to surface value may be reduced to 1.5 m where the building does not have fire escapes, balconies, and windows that can be opened adjacent to the conductor.',
                   'Voltages are rms line-to-ground'
                   ]

    # define the worksheet
    CSA_Table_9 = {
        'ws_name': 'CSA Table 9',
        'ws_content': CSA9,
        'cell_range_style': cell_format_CSA_9,
        'footer': footer_CSA9
    }
#endregion

#CSA Table 10
#region
    voltage = inputs['p2p_voltage']
    voltage_range = CSA_Table10_data['Voltage range']

    #Creating the title blocks. etc
    CSA10_cell00 = str(voltage_range)
    
    CSA10 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 10 \n Minimum Design Clearances from Supply Conductors to Bridges \n (See clauses 5.7.4.2.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        [' ', 'Minimum design clearance from supply conductor to bridge, m'],
        [' ', 'Readily accessible portions*',' ',' ',' ',' ' ,' ',' ',' ',' ',' ',' ',' ','Inaccessable portions'],
        [' ', 'Horizontal', ' ', ' ', ' ', 'Vertical', ' ', ' ', ' ', ' ', ' ', ' ', ' ', 'Horizontal', ' ', ' ', ' ', 'Vertical'],
        [' ', 'Conductor attached to bridge †', ' ', 'Conductor not attached †', ' ', 'Conductor attached to bridge', ' ', ' ', ' ', 'Conductor not attached', ' ', ' ', ' ', \
         'Conductor attached to bridge', ' ', 'Conductor not attached', ' ', 'Conductor attached to bridge', ' ', ' ', ' ', 'Conductor not attached'],
        [' ', ' ', ' ', ' ', ' ', 'Over Bridge', ' ', 'Under Bridge', ' ', 'Over Bridge', ' ', 'Under Bridge', ' ', ' ', ' ', ' ', ' ', 'Over Bridge', ' ', 'Under Bridge', ' ', 'Over Bridge', ' ', 'Under Bridge', ' '],
        ['Supply Conductor', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance', \
         'Basic', 'Design Clearance', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance'],
        [CSA10_cell00] + [CSA_Table10_data['CSA10_Clearance'][i] for i in range(24)],
              ]

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_10 = len(CSA10)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_10 = len(CSA10[8])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA10 = []
    for i in range(n_row_CSA_table_10):
        if i < 9:
            list_range_color_CSA10.append((i + 1, 1, i + 1, n_column_CSA_table_10, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA10.append((i + 1, 1, i + 1, n_column_CSA_table_10, color_bkg_data_1))
        else:
            list_range_color_CSA10.append((i + 1, 1, i + 1, n_column_CSA_table_10, color_bkg_data_2))

    # define cell format
    cell_format_CSA_10 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_10, 'center'), (4, 1, 8, 1, 'center'), (4, 2, 4, 25, 'center'), (5, 2, 5, 13, 'center'), (5, 14, 5, 25, 'center'), (6, 2, 6, 5, 'center'), (6, 6, 6, 13, 'center'), (6, 14, 6, 17, 'center'), (6, 18, 6, 25, 'center'), \
                        (7, 2, 8, 3, 'center'), (7, 4, 8, 5, 'center'), (7, 6, 7, 9, 'center'), (7, 10, 7, 13, 'center'), (7, 14, 8, 15, 'center'), (7, 16, 8, 17, 'center'), (7, 18, 7, 21, 'center'), (7, 22, 7, 25, 'center'),\
                        (8, 6, 8, 7, 'center'), (8, 8, 8, 9, 'center'), (8, 10, 8, 11, 'center'), (8, 12, 8, 13, 'center'), (8, 18, 8, 19, 'center'), (8, 20, 8, 21, 'center'), (8, 22, 8, 23, 'center'), (8, 24, 8, 25, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA10,
        'range_border': [(1, 1, 1, n_column_CSA_table_10), (2, 1, n_row_CSA_table_10, n_column_CSA_table_10)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 10) for i in range(n_column_CSA_table_10)],
    }

    # define some footer notes
    footer_CSA10 = ['* Readily accessible portions include portions of a bridge likely to be used by workers and portions readily accessible to the public. Bridge seats of steel bridges supported by masonry, brick, or concrete abutments that need periodic inspection should be considered readily accessible portions. Concrete, brick, or masonry portions that need to be passed by workers so that they can gain access to other areas shall be considered readily accessible portions.',
                   '† Below 750 V this clearance may be reduced to 1 m where the portion of the bridge involved is readily accessible to workers but not to the public.', 
                   '0-750V 1.Insulated or grounded  2.Not insulated, grounded, or enclosed in effectively grounded metallic sheath, Vertical to Surface may be reduced to 1 m for portions of the building considered normally inaccessible.',
                   'Voltages are rms line-to-ground.'
                   'Grounding conductors or conductors installed in effectively grounded conduit do not require clearance under a bridge.'
                   ]

    # define the worksheet
    CSA_Table_10 = {
        'ws_name': 'CSA Table 10',
        'ws_content': CSA10,
        'cell_range_style': cell_format_CSA_10,
        'footer': footer_CSA10
    }
#endregion

#CSA Table 11
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    voltage_range = CSA_Table11_data['Voltage range']

    #Creating the title blocks. etc
    CSA11_cell00 = 'Supply equipment'
    CSA11_cell10 = 'Guys, messengers, span wires, communication circuits, secondary cable 0–750 V, and multi-grounded neutral conductors'
    CSA11_cell20 = 'Supply conductors' + str(voltage_range) + ' kV*'

    CSA11_titles = np.array([CSA11_cell00, CSA11_cell10, CSA11_cell20])
    
    CSA11 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 11 \n Minimum Separation of Equipment and Clearances of Conductors from Swimming Pools, ac \n (See clauses 5.7.5 and A.5.7.5 and Figure 1.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Wire closest to tracks', ' ', 'Minimum clearance, m'],
        [' ', ' ', 'A — Measured in any direction from the water level, edge of pool, or diving platform', ' ', 'B — Measured vertically over land'],
        [' ', ' ', 'Basic', 'Design Clearance'],
              ]
    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(3):
        row = [CSA11_titles[i], ' ',  CSA_Table11_data['Basic'][i], CSA_Table11_data['Design Clearance'][i], CSA_Table11_data['B-Measured Vertically Over Land'][i]]
        CSA11.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_11 = len(CSA11)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_11 = len(CSA11[6])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA11 = []
    for i in range(n_row_CSA_table_11):
        if i < 6:
            list_range_color_CSA11.append((i + 1, 1, i + 1, n_column_CSA_table_11, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA11.append((i + 1, 1, i + 1, n_column_CSA_table_11, color_bkg_data_1))
        else:
            list_range_color_CSA11.append((i + 1, 1, i + 1, n_column_CSA_table_11, color_bkg_data_2))

    # define cell format
    cell_format_CSA_11 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_11, 'center'), (4, 1, 6, 2, 'center'), (4, 3, 4, 5, 'center'), (7, 1, 7, 2, 'center'), (8, 1, 8, 2, 'center'), (5, 3, 5, 4, 'center'), (5, 5, 6, 5, 'center'), (9, 1, 9, 2, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA11,
        'range_border': [(1, 1, 1, n_column_CSA_table_11), (2, 1, n_row_CSA_table_11, n_column_CSA_table_11)],
        'row_height': [(1, 50),(5, 100)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_11)],
    }

    # define some footer notes
    footer_CSA11 = ['Voltages are rms line-to-ground.',
                   'See Figure 1 for an illustration of swimming pool clearance limits.', 
                   '*Supply conductors greater than 150 kV shall not cross over swimming pools or be closer to the edge of a pool than 8.0m + 0.01m/kV above 150 kV',
                   ]

    # define the worksheet
    CSA_Table_11 = {
        'ws_name': 'CSA Table 11',
        'ws_content': CSA11,
        'cell_range_style': cell_format_CSA_11,
        'footer': footer_CSA11
    }
#endregion

#CSA Table 13
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    XING_voltage = inputs["XING_P2P_Voltage"]
    XING_voltage = np.char.mod(f'%.{Numpy_round_integer}f', XING_voltage)
    line2ground_voltage = voltage / np.sqrt(3)
    line2ground_voltage = np.char.mod(f'%.{Numpy_round_integer}f', line2ground_voltage)
    voltage_range = CSA_Table13_data['Voltage range']
    voltage_range_XING = CSA_Table13_data['Voltage range XING']

    #Creating the title blocks. etc
    CSA13_cell00 = 'Communication wires and cables'
    CSA13_cell10 = 'Open supply conductors, ac, kV'
    CSA13_cell20 = 'Guys, span wires, and aerial grounding wires'

    CSA13_cell01 = ' '
    CSA13_cell11 = str(voltage_range)
    CSA13_cell21 = ' '

    CSA13_titles = np.array([CSA13_cell00, CSA13_cell10, CSA13_cell20])
    CSA13_1in = np.array([CSA13_cell01, CSA13_cell11, CSA13_cell21])
    
    CSA13 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 13 \n Minimum Design Vertical Clearances between Wires Crossing Each Other and supported by Different Supporting Structures, ac \n (See clauses 5.8.1.1, 5.8.1.2, A.5.8.1, and A.5.8.1.3.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Type of line wire, cable, or other plant being crossed over', ' ', 'Line wire or cable at upper level, minimum clearance, m'],
        [' ', ' ', '(line to ground voltage:' + str(XING_voltage) + ' kV)'],
        [' ', ' ', 'Guys, span wires, aerial grounding conductors, and communication wires and cables', ' ', 'Open supply-line conductors and service wires, ac, ' + str(voltage_range_XING) +' kV'],
        ['(line to ground voltage ' + str(line2ground_voltage) + ' kV)', ' ', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance'],
              ]
    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(3):
        row = [CSA13_titles[i], CSA13_1in[i],  CSA_Table13_data['Basic guy'][i], CSA_Table13_data['Design Clearance guy'][i], CSA_Table13_data['Basic ac'][i], CSA_Table13_data['Design Clearance ac'][i]]
        CSA13.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_13 = len(CSA13)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_13 = len(CSA13[8])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA13 = []
    for i in range(n_row_CSA_table_13):
        if i < 6:
            list_range_color_CSA13.append((i + 1, 1, i + 1, n_column_CSA_table_13, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA13.append((i + 1, 1, i + 1, n_column_CSA_table_13, color_bkg_data_1))
        else:
            list_range_color_CSA13.append((i + 1, 1, i + 1, n_column_CSA_table_13, color_bkg_data_2))

    # define cell format
    cell_format_CSA_13 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_13, 'center'), (4, 1, 6, 2, 'center'), (4, 3, 4, 6, 'center'), (5, 3, 5, 6, 'center'), (6, 3, 6, 4, 'center'), (6, 5, 6, 6, 'center'), (7, 1, 7, 2, 'center'), (8, 1, 8, 2, 'center'), (10, 1, 10, 2, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA13,
        'range_border': [(1, 1, 1, n_column_CSA_table_13), (2, 1, n_row_CSA_table_13, n_column_CSA_table_13)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_13)],
    }

    # define some footer notes
    footer_CSA13 = ['Voltages are rms line-to-ground.',
                   'Refer to Table A.1 for typical nominal system voltages and overvoltage values.', 
                   ' Refer to Clause A.5.8.1 for explanation of the calculation method used to produce the clearance values shown in this Table.',
                   ]

    # define the worksheet
    CSA_Table_13 = {
        'ws_name': 'CSA Table 13',
        'ws_content': CSA13,
        'cell_range_style': cell_format_CSA_13,
        'footer': footer_CSA13
    }
#endregion

#CSA Table 14
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table14_data['Voltage range']

    #Creating the title blocks. etc
    CSA14_cell00 = 'Aerial tramways Gondolas and similar apparatus providing a roof over passengers'
    CSA14_cell10 = 'Chairlifts, T-bars, and similar apparatus or towers of any type of aerial tramway'

    CSA14_titles = np.array([CSA14_cell00, CSA14_cell10])
    
    CSA14 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 14 \n Minimum Design Vertical Clearances for Crossings over Aerial Tramways \n (See clauses 5.8.1.1.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Type of plant being crossed over', 'Line wire or cable at upper level, minimum clearance, m'],
        [' ', 'Communication conductors and cables, span and grounding wires carried aerially', ' ', 'Open supply-line conductors ac,' + str(voltage_range)],
        [' ', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance'],
            ]
    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(2):
        row = [CSA14_titles[i],  CSA_Table14_data['Basic guy'][i], CSA_Table14_data['Design Clearance guy'][i], CSA_Table14_data['Basic'][i], CSA_Table14_data['Design Clearance'][i]]
        CSA14.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_14 = len(CSA14)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_14 = len(CSA14[6])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA14 = []
    for i in range(n_row_CSA_table_14):
        if i < 6:
            list_range_color_CSA14.append((i + 1, 1, i + 1, n_column_CSA_table_14, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA14.append((i + 1, 1, i + 1, n_column_CSA_table_14, color_bkg_data_1))
        else:
            list_range_color_CSA14.append((i + 1, 1, i + 1, n_column_CSA_table_14, color_bkg_data_2))

    # define cell format
    cell_format_CSA_14 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_14, 'center'), (4, 1, 6, 1, 'center'), (4, 2, 4, 5, 'center'), (5, 2, 5, 3, 'center'), (5, 4, 5, 5, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA14,
        'range_border': [(1, 1, 1, n_column_CSA_table_14), (2, 1, n_row_CSA_table_14, n_column_CSA_table_14)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_14)],
    }

    # define some footer notes
    footer_CSA14 = ['Voltages are rms line-to-ground.']

    # define the worksheet
    CSA_Table_14 = {
        'ws_name': 'CSA Table 14',
        'ws_content': CSA14,
        'cell_range_style': cell_format_CSA_14,
        'footer': footer_CSA14
    }
#endregion

#CSA Table 15
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table15_data['Voltage range']

    #Creating the title blocks. etc
    CSA15_cell00 = str(voltage_range)
    CSA15_cell10 = 'Communication conductors'

    CSA15_titles = np.array([CSA15_cell00, CSA15_cell10])
    
    CSA15 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 15 \n Clearances between conductors supported by different structures but not crossing each other — \n Clearance increments to be added to horizontal displacement \n (See clauses 5.8.2 and A.5.8.2.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        [' '],
        ['Sum of voltages of conductors*', 'Clearance increment, mm'],
        [' ', 'Basic', 'Design Clearance'],
            ]
    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(2):
        row = [CSA15_titles[i],  CSA_Table15_data['Basic'][i], CSA_Table15_data['Design_Clearance'][i]]
        CSA15.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_15 = len(CSA15)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_15 = len(CSA15[6])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA15 = []
    for i in range(n_row_CSA_table_15):
        if i < 6:
            list_range_color_CSA15.append((i + 1, 1, i + 1, n_column_CSA_table_15, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA15.append((i + 1, 1, i + 1, n_column_CSA_table_15, color_bkg_data_1))
        else:
            list_range_color_CSA15.append((i + 1, 1, i + 1, n_column_CSA_table_15, color_bkg_data_2))

    # define cell format
    cell_format_CSA_15 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 4, n_column_CSA_table_15, 'center'), (5, 1, 6, 1, 'center'), (5, 2, 5, 3, 'center'), (5, 4, 5, 5, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA15,
        'range_border': [(1, 1, 1, n_column_CSA_table_15), (2, 1, n_row_CSA_table_15, n_column_CSA_table_15)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_15)],
    }

    # define some footer notes
    footer_CSA15 = ['* Communication conductors are assumed to operate at less than 130 V Design_Clearance at 0.25 A or less per pair (1500 pairs or fewer).', 'Voltages are rms line-to-ground.']

    # define the worksheet
    CSA_Table_15 = {
        'ws_name': 'CSA Table 15',
        'ws_content': CSA15,
        'cell_range_style': cell_format_CSA_15,
        'footer': footer_CSA15
    }
#endregion

#CSA Table 16
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table16_data['Voltage range']

    #Creating the title blocks. etc
    CSA16_cell00 = str(voltage_range)
    CSA16_cell10 = 'Communication conductors'

    CSA16_titles = np.array([CSA16_cell00, CSA16_cell10])
    
    CSA16 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 16 \n Minimum design clearances between conductors of one line and supporting structures of another line \n (See clause 5.8.3.1.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Voltage of line conductor*, kV', 'Minimum clearance between conductor and supporting structure, mm'],
        [' ', 'Basic', 'Design Clearance'],
            ]
    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(2):
        row = [CSA16_titles[i],  CSA_Table16_data['Basic'][i], CSA_Table16_data['Design_Clearance'][i]]
        CSA16.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_16 = len(CSA16)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_16 = len(CSA16[6])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA16 = []
    for i in range(n_row_CSA_table_16):
        if i < 5:
            list_range_color_CSA16.append((i + 1, 1, i + 1, n_column_CSA_table_16, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA16.append((i + 1, 1, i + 1, n_column_CSA_table_16, color_bkg_data_1))
        else:
            list_range_color_CSA16.append((i + 1, 1, i + 1, n_column_CSA_table_16, color_bkg_data_2))

    # define cell format
    cell_format_CSA_16 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_16, 'center'), (4, 1, 5, 1, 'center'), (4, 2, 4, 3, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA16,
        'range_border': [(1, 1, 1, n_column_CSA_table_16), (2, 1, n_row_CSA_table_16, n_column_CSA_table_16)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_16)],
    }

    # define some footer notes
    footer_CSA16 = ['* Communication conductors are assumed to operate at less than 130 V Design_Clearance at 0.25 A or less per pair (1500 pairs or fewer).', 'Voltages are rms line-to-ground.', 'Spans are assumed to be greater than 15m.', 'For line voltages (line to ground) between 0 and 22 kV, clearance of 1 m shall be met wherever practical, but see CSA C22.3 No.1 for details on allowed reductions if necessary.']

    # define the worksheet
    CSA_Table_16 = {
        'ws_name': 'CSA Table 16',
        'ws_content': CSA16,
        'cell_range_style': cell_format_CSA_16,
        'footer': footer_CSA16
    }
#endregion

#CSA Table 17
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table17_data['Voltage range']
    span_range = CSA_Table17_data['Span Range']

    #Creating the title blocks. etc
    CSA17_cell00 = str(voltage_range)
    
    CSA17 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 17 \n Minimum horizontal separations of supply-line conductors attached to the same supporting structure \n (See clauses 5.9.1.1, 5.9.1.2, and 5.9.1.3.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Line conductor', 'Minimum horizontal separation of conductors for spans mm'],
        [' ', 'span ' + str(span_range)],
        [' ', 'Basic', 'Design Clearance'],
        [CSA17_cell00 ,  CSA_Table17_data['Basic'], CSA_Table17_data['Design_Clearance']]
            ]
    

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_17 = len(CSA17)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_17 = len(CSA17[5])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA17 = []
    for i in range(n_row_CSA_table_17):
        if i < 6:
            list_range_color_CSA17.append((i + 1, 1, i + 1, n_column_CSA_table_17, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA17.append((i + 1, 1, i + 1, n_column_CSA_table_17, color_bkg_data_1))
        else:
            list_range_color_CSA17.append((i + 1, 1, i + 1, n_column_CSA_table_17, color_bkg_data_2))

    # define cell format
    cell_format_CSA_17 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_17, 'center'), (4, 1, 6, 1, 'center'), (4, 2, 4, 3, 'center'), (5, 2, 5, 3, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA17,
        'range_border': [(1, 1, 1, n_column_CSA_table_17), (2, 1, n_row_CSA_table_17, n_column_CSA_table_17)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_17)],
    }

    # define some footer notes
    footer_CSA17 = [' For spans longer than 450 m, the separation shall be based on best engineering practices, but shall be not less than the separations specified for spans of 450 m.', '† Phase-to-phase voltages for the same circuit or the sum of phase-to-ground voltages for different circuits']

    # define the worksheet
    CSA_Table_17 = {
        'ws_name': 'CSA Table 17',
        'ws_content': CSA17,
        'cell_range_style': cell_format_CSA_17,
        'footer': footer_CSA17
    }
#endregion

#CSA Table 18
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table18_data['Voltage range']
    Voltage_range_under = CSA_Table18_data['Voltage range under']

    #Creating the title blocks. etc
    CSA18_cell00 = str(Voltage_range_under)
    
    CSA18 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 18 \n Minimum vertical separations between supply-line conductors attached to the same supporting structure, ac \n (See clause 5.9.2.1 and A.4.1.5.) \n System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Conductors at lower level', 'Minimum vertical separation at supporting structure, m'],
        [' ', 'Conductors at higher level: ' +  str(voltage_range)],
        [' ', 'Basic', 'Design Clearance'],
        [CSA18_cell00 ,  CSA_Table18_data['Basic'], CSA_Table18_data['Design Clearance']]
            ]
    

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_18 = len(CSA18)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_18 = len(CSA18[5])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA18 = []
    for i in range(n_row_CSA_table_18):
        if i < 6:
            list_range_color_CSA18.append((i + 1, 1, i + 1, n_column_CSA_table_18, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA18.append((i + 1, 1, i + 1, n_column_CSA_table_18, color_bkg_data_1))
        else:
            list_range_color_CSA18.append((i + 1, 1, i + 1, n_column_CSA_table_18, color_bkg_data_2))

    # define cell format
    cell_format_CSA_18 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_18, 'center'), (4, 1, 6, 1, 'center'), (4, 2, 4, 3, 'center'), (5, 2, 5, 3, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA18,
        'range_border': [(1, 1, 1, n_column_CSA_table_18), (2, 1, n_row_CSA_table_18, n_column_CSA_table_18)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_18)],
    }

    # define some footer notes
    footer_CSA18 = ['† This value does not apply to conductors on adjacent supports of the same circuit or circuits.', 'Voltages are rms line-to-ground.']

    # define the worksheet
    CSA_Table_18 = {
        'ws_name': 'CSA Table 18',
        'ws_content': CSA18,
        'cell_range_style': cell_format_CSA_18,
        'footer': footer_CSA18
    }
#endregion

#CSA Table 20
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table20_data['Voltage range']

    #Creating the title blocks. etc
    CSA20_cell00 = str(voltage_range)
    
    CSA20 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 20 \n \
         Minimum in-span vertical clearances between supply conductors of the same circuit that are attached to the same supporting structure \n \
         (See clause 5.9.2.3.) \n \
         System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Maximum circuit line-to-ground voltage, kV', 'Minimum clearance, mm'],
        [' ', 'Between multi-grounded neutral and circuit conductor', ' ', 'Between circuit conductors'],
        [' ', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance'],
        [CSA20_cell00 ,  CSA_Table20_data['Basic'], CSA_Table20_data['Design_Clearance'],  CSA_Table20_data['Between conductors'], CSA_Table20_data['Between conductors Design_Clearance']]
            ]
    

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_20 = len(CSA20)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_20 = len(CSA20[6])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA20 = []
    for i in range(n_row_CSA_table_20):
        if i < 6:
            list_range_color_CSA20.append((i + 1, 1, i + 1, n_column_CSA_table_20, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA20.append((i + 1, 1, i + 1, n_column_CSA_table_20, color_bkg_data_1))
        else:
            list_range_color_CSA20.append((i + 1, 1, i + 1, n_column_CSA_table_20, color_bkg_data_2))

    # define cell format
    cell_format_CSA_20 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_20, 'center'), (4, 1, 6, 1, 'center'), (4, 2, 4, 5, 'center'), (5, 2, 5, 3, 'center'),(5, 4, 5, 5, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA20,
        'range_border': [(1, 1, 1, n_column_CSA_table_20), (2, 1, n_row_CSA_table_20, n_column_CSA_table_20)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_20)],
    }

    # This table has no footer notes

    # define the worksheet
    CSA_Table_20 = {
        'ws_name': 'CSA Table 20',
        'ws_content': CSA20,
        'cell_range_style': cell_format_CSA_20,
    }
#endregion

#CSA Table 21
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table21_data['Voltage range']
    voltage_range_XING = CSA_Table21_data['Voltage range under']

    #Creating the title blocks. etc
    CSA21_cell00 = str(voltage_range)
    
    CSA21 = [
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 21 \n \
         Minimum in-span vertical clearances between supply conductors of different circuits that are attached to the same supporting structure \n \
         System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Maximum lower conductor line-to-ground voltage, kV', 'Minimum in-span vertical clearance, mm'],
        [' ', 'Maximum upper conductor line-to-ground voltage, kV', ' '],
        [' ', str(voltage_range_XING), ' '],
        [' ', 'Basic', 'Design Clearance'],
        [CSA21_cell00 ,  CSA_Table21_data['Basic'], CSA_Table21_data['Design Clearance']]
            ]
    

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_21 = len(CSA21)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_21 = len(CSA21[6])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA21 = []
    for i in range(n_row_CSA_table_21):
        if i < 7:
            list_range_color_CSA21.append((i + 1, 1, i + 1, n_column_CSA_table_21, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA21.append((i + 1, 1, i + 1, n_column_CSA_table_21, color_bkg_data_1))
        else:
            list_range_color_CSA21.append((i + 1, 1, i + 1, n_column_CSA_table_21, color_bkg_data_2))

    # define cell format
    cell_format_CSA_21 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_21, 'center'), (4, 1, 7, 1, 'center'), (4, 2, 4, 3, 'center'), (5, 2, 5, 3, 'center'),(6, 2, 6, 3, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA21,
        'range_border': [(1, 1, 1, n_column_CSA_table_21), (2, 1, n_row_CSA_table_21, n_column_CSA_table_21)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_21)],
    }

    # This table has no footer notes

    # define the worksheet
    CSA_Table_21 = {
        'ws_name': 'CSA Table 21',
        'ws_content': CSA21,
        'cell_range_style': cell_format_CSA_21,
    }
#endregion

#CSA Table 22
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table22_data['Voltage range']

    #Creating the title blocks. etc
    CSA22_cell00 = 'Supply lateral, vertical, or line conductors and supply lateral or vertical conductors of the same or different circuits but not connected together'
    CSA22_cell10 = 'Supply lateral, vertical, or line conductors and surface of structure, crossarms, and other non-energized supply plant, including grounding conductors'
    CSA22_cell20 = 'Supply lateral, vertical, or line conductors and span or guy wire, except where conductors are supported by the span wire'
    CSA22_cell30 = 'Supply-line conductor and lightning protection wire parallel to the line'

    CSA22titles = np.array([CSA22_cell00, CSA22_cell10, CSA22_cell20, CSA22_cell30])
    
    CSA22 = ([
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 22 \n \
         Minimum separations or clearances in any direction from supply conductors to other supply plant attached to the same supporting structure \n \
         (See clause 5.9.5 and A.5.9.5.) \n \
         System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Between', 'Separation or clearance, mm'],
        [' ', 'Voltage of conductor(s), ac*'],
        [' ', str(voltage_range)],
        [' ', 'Basic', 'Design Clearance'],
            ])
    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(4):
        row = [CSA22titles[i],  CSA_Table22_data['Basic'][i], CSA_Table22_data['Design_Clearance'][i]]
        CSA22.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_22 = len(CSA22)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_22 = len(CSA22[9])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA22 = []
    for i in range(n_row_CSA_table_22):
        if i < 7:
            list_range_color_CSA22.append((i + 1, 1, i + 1, n_column_CSA_table_22, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA22.append((i + 1, 1, i + 1, n_column_CSA_table_22, color_bkg_data_1))
        else:
            list_range_color_CSA22.append((i + 1, 1, i + 1, n_column_CSA_table_22, color_bkg_data_2))

    # define cell format
    cell_format_CSA_22 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': [(1, 1, 3, n_column_CSA_table_22, 'center'), (4, 1, 7, 1, 'center'), (4, 2, 4, 3, 'center'), (5, 2, 5, 3, 'center'), (6, 2, 6, 3, 'center')],
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA22,
        'range_border': [(1, 1, 1, n_column_CSA_table_22), (2, 1, n_row_CSA_table_22, n_column_CSA_table_22)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_22)],
    }

    # define some footer notes
    footer_CSA22 = (['* Clearances from connecting wires to switches or arresters may be less than that from line wires to switches or arresters.', 
                     '† For voltages exceeding 50 kV, clearances shall be based on best engineering practices, but shall be not less than the separations specified for 50 kV.',
                     '** This clearance may be reduced to 0 mm for conductors and cables 0–750 V, provided that the conductors are attached to the surface of the structure and grounded or covered by material of adequate insulating and mechanical properties.', 
                     '†† Where this clearance cannot be achieved, adequate insulation shall be applied to the supply conductor.', 
                     '‡‡ This clearance shall be increased to 300 mm for voltages 0.75–8 kV for a span wire or span guy wire running parallel to the supply conductor.', 
                     '§§ The clearance shall be not less than the separations specified in Clause 5.9.1 (table 17).',
                     'Voltages are rms line-to-ground.'])

    # define the worksheet
    CSA_Table_22 = {
        'ws_name': 'CSA Table 22',
        'ws_content': CSA22,
        'cell_range_style': cell_format_CSA_22,
        'footer': footer_CSA22
    }
#endregion

#CSA Table 23
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table23_data['Voltage range']

    #Creating the title blocks. etc
    CSA23_cell00 = 'Live or current-carrying supply plant (including neutrals) and communication line plant'
    CSA23_cell10 = 'Non-energized supply plant (excluding luminaire span wire and brackets) and communication line plant'
    CSA23_cell20 = ' '
    CSA23_cell30 = ' '
    CSA23_cell40 = ' '
    CSA23_cell50 = 'Trolley span wires or brackets and communication line plant'
    CSA23_cell60 = 'Luminaire span wires or brackets and communication line plant'
    CSA23_cell70 = ''
    CSA23_cell80 = 'Point of attachment of combined communication drop and supply service conductor and communication line plant'
    CSA23_cell90 = 'Housing containing communication power supply, communication compressor dehydrator, or other communication equipment (effectively grounded or insulated) and communication cable'

    CSA23_cell01 = ' '
    CSA23_cell11 = 'Option A§'
    CSA23_cell21 = ' '
    CSA23_cell31 = 'Option B§'
    CSA23_cell41 = ' '
    CSA23_cell51 = ' '
    CSA23_cell61 = ' '
    CSA23_cell71 = ' '
    CSA23_cell81 = ' '
    CSA23_cell91 = ' '

    CSA23_cell02 = ' '
    CSA23_cell12 = 'Ungrounded'
    CSA23_cell22 = 'Effectively Grounded'
    CSA23_cell32 = 'Ungrounded'
    CSA23_cell42 = 'Effectively Grounded'
    CSA23_cell52 = ' '
    CSA23_cell62 = 'Ungrounded'
    CSA23_cell72 = 'Effectively Grounded'
    CSA23_cell82 = ' '
    CSA23_cell92 = ' '

    CSA23titles = np.array([CSA23_cell00, CSA23_cell10, CSA23_cell20, CSA23_cell30, CSA23_cell40, CSA23_cell50, CSA23_cell60, CSA23_cell70, CSA23_cell80, CSA23_cell90])
    CSA23_1in = np.array([CSA23_cell01, CSA23_cell11, CSA23_cell21, CSA23_cell31, CSA23_cell41, CSA23_cell51, CSA23_cell61, CSA23_cell71, CSA23_cell81, CSA23_cell91])
    CSA23_2in = np.array([CSA23_cell02, CSA23_cell12, CSA23_cell22, CSA23_cell32, CSA23_cell42, CSA23_cell52, CSA23_cell62, CSA23_cell72, CSA23_cell82, CSA23_cell92])
    
    CSA23 = ([
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 23 \n \
         Minimum vertical separations at a joint-use structure \n \
         (See clauses 5.10.1.1, 5.10.1.6, 5.10.1.7, 5.10.6.2 and A.5.10.1 and figures A.10 and A.11.) \n \
         System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Between', ' ', ' ','Minimum vertical separation, m'],
        [' ', ' ', ' ', 'Voltage of supply conductors'],
        [' ', ' ', ' ', str(voltage_range)],
        [' ', ' ', ' ', 'Basic', 'Design Clearance'],
            ])
    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(10):
        row = [CSA23titles[i], CSA23_1in[i], CSA23_2in[i], CSA_Table23_data['Basic'][i], CSA_Table23_data['Design_Clearance'][i]]
        CSA23.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_23 = len(CSA23)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_23 = len(CSA23[9])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA23 = []
    for i in range(n_row_CSA_table_23):
        if i < 7:
            list_range_color_CSA23.append((i + 1, 1, i + 1, n_column_CSA_table_23, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA23.append((i + 1, 1, i + 1, n_column_CSA_table_23, color_bkg_data_1))
        else:
            list_range_color_CSA23.append((i + 1, 1, i + 1, n_column_CSA_table_23, color_bkg_data_2))

    # define cell format
    cell_format_CSA_23 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': ([(1, 1, 3, n_column_CSA_table_23, 'center'), (4, 1, 7, 3, 'center'), (4, 4, 4, 5, 'center'), (5, 4, 5, 5, 'center'), (6, 4, 6, 5, 'center'), (8, 1, 8, 3, 'center'),
                         (9, 1, 12, 1, 'center'), (9, 2, 10, 2, 'center'), (11, 2, 12, 2, 'center'), (13, 1, 13, 3, 'center'), (14, 1, 15, 2, 'center'), (16, 1, 16, 3, 'center'), (17, 1, 17, 3, 'center')]),
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA23,
        'range_border': [(1, 1, 1, n_column_CSA_table_23), (2, 1, n_row_CSA_table_23, n_column_CSA_table_23)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_23)],
    }

    # define some footer notes
    footer_CSA23 = (['† On lateral communication drop wire plant, this separation may be reduced to 0.6 m.', 
                     '‡ See Clause 5.2.7 for requirements for neutral conductors.',
                     '§ Option A or Option B shall be selected in accordance with Clauses 5.10.1.6 and 5.10.1.7.',
                     'Voltages are rms line-to-ground.'])

    # define the worksheet
    CSA_Table_23 = {
        'ws_name': 'CSA Table 23',
        'ws_content': CSA23,
        'cell_range_style': cell_format_CSA_23,
        'footer': footer_CSA23
    }
#endregion

#CSA Table 24
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table24_data['Voltage range']
    
    CSA24 = ([
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 24 \n \
         Minimum in-span vertical clearances between supply and communication conductors \n \
         (See clauses 5.10.3.2 and A.5.10.3.) \n \
         System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Voltage of supply conductor','Minimum clearance of supply conductor above line of sight of points of support of highest communication wire or cable, mm'],
        [' ', 'Basic', 'Design Clearance'],
        [str(voltage_range), CSA_Table24_data['Basic'], CSA_Table24_data['Design_Clearance']]
            ])

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_24 = len(CSA24)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_24 = len(CSA24[5])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA24 = []
    for i in range(n_row_CSA_table_24):
        if i < 5:
            list_range_color_CSA24.append((i + 1, 1, i + 1, n_column_CSA_table_24, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA24.append((i + 1, 1, i + 1, n_column_CSA_table_24, color_bkg_data_1))
        else:
            list_range_color_CSA24.append((i + 1, 1, i + 1, n_column_CSA_table_24, color_bkg_data_2))

    # define cell format
    cell_format_CSA_24 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': ([(1, 1, 3, n_column_CSA_table_24, 'center'), (4, 1, 5, 1, 'center'), (4, 2, 4, 3, 'center')]),
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA24,
        'range_border': [(1, 1, 1, n_column_CSA_table_24), (2, 1, n_row_CSA_table_24, n_column_CSA_table_24)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_24)],
    }

    # define some footer notes
    footer_CSA24 = (['* For effectively grounded neutral, see Clause 5.10.3.3.', 
                     '† While the use of the design limit of 75 mm can yield a minimum actual clearance of approximately 300 mm under the worst expected conditions, most situations will result in clearances in excess of 600 mm.',
                     'Voltages are rms line-to-ground.'])

    # define the worksheet
    CSA_Table_24 = {
        'ws_name': 'CSA Table 24',
        'ws_content': CSA24,
        'cell_range_style': cell_format_CSA_24,
        'footer': footer_CSA24
    }
#endregion

#CSA Table 25
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table25_data['Voltage range']
    
    CSA25_cell00 = 'Communication line plant'
    CSA25_cell10 = 'Current-carrying supply plant, ac' + str(voltage_range)
    CSA25_cell20 = 'Supply guy wires and span wires'

    CSA25titles = np.array([CSA25_cell00, CSA25_cell10, CSA25_cell20])

    CSA25 = ([
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 25 \n \
         Minimum clearances from guys to plant of another system \n \
         (See clauses 5.11.2.1 and A.5.11.2.) \n \
         System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Type of plant near which the guy passes','Minimum clearance, m'],
        [' ', 'Basic', 'Design Clearance'],
            ])

    
    #The following will fill out the rest of the table with numbers in their respective problems
    for i in range(3):
        row = [CSA25titles[i], CSA_Table25_data['Basic'][i], CSA_Table25_data['Design_Clearance'][i]]
        CSA25.append(row)

    #This is retrieving the number of rows in each table array
    n_row_CSA_table_25 = len(CSA25)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_25 = len(CSA25[6])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA25 = []
    for i in range(n_row_CSA_table_25):
        if i < 5:
            list_range_color_CSA25.append((i + 1, 1, i + 1, n_column_CSA_table_25, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA25.append((i + 1, 1, i + 1, n_column_CSA_table_25, color_bkg_data_1))
        else:
            list_range_color_CSA25.append((i + 1, 1, i + 1, n_column_CSA_table_25, color_bkg_data_2))

    # define cell format
    cell_format_CSA_25 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': ([(1, 1, 3, n_column_CSA_table_25, 'center'), (4, 1, 5, 1, 'center'), (4, 2, 4, 3, 'center')]),
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA25,
        'range_border': [(1, 1, 1, n_column_CSA_table_25), (2, 1, n_row_CSA_table_25, n_column_CSA_table_25)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_25)],
    }

    # define some footer notes
    footer_CSA25 = (['Voltages are rms line-to-ground.'])

    # define the worksheet
    CSA_Table_25 = {
        'ws_name': 'CSA Table 25',
        'ws_content': CSA25,
        'cell_range_style': cell_format_CSA_25,
        'footer': footer_CSA25
    }
#endregion

#CSA Table 26
#region

    #Importing voltage and voltage range to be used in the table titles
    voltage = inputs['p2p_voltage']
    line2ground_voltage = voltage / np.sqrt(3)
    voltage_range = CSA_Table25_data['Voltage range']
    
    CSA26_cell00 = 'Communication line plant'
    CSA26_cell10 = 'Current-carrying supply plant' + str(voltage_range)

    CSA26titles = np.array([CSA26_cell00, CSA26_cell10])

    CSA26 = ([
    #The following is the title header before there is data
        ['CSA C22-3 No. 1-20 Table 26 \n \
         Minimum clearance or separation between guys and other plant attached to the joint-use structure \n \
         (See clauses 5.11.2.2 and A.5.11.2.) \n \
         System Voltage: ' + str(voltage) + ' kV (AC 3-phase)'],
        [' '],
        [' '],
        ['Type of plant over or near which the guy passes','Minimum clearance or separation from guy, mm'],
        [' ','Guy not parallel to plant', ' ', 'Guy parallel to plant'],
        [' ', 'Basic', 'Design Clearance', 'Basic', 'Design Clearance'],
        [CSA26titles[0], CSA_Table26_data['Comms'][0], CSA_Table26_data['Comms'][1], CSA_Table26_data['Comms'][2], CSA_Table26_data['Comms'][3]],
        [CSA26titles[1], CSA_Table26_data['Conductors'][0], CSA_Table26_data['Conductors'][1], CSA_Table26_data['Conductors'][2], CSA_Table26_data['Conductors'][3]]
            ])


    #This is retrieving the number of rows in each table array
    n_row_CSA_table_26 = len(CSA26)

    #This is retrieving the number of columns in the row specified in the columns
    n_column_CSA_table_26 = len(CSA26[5])

    #Creating an empty variable to determine the colour format that is used in the table
    list_range_color_CSA26 = []
    for i in range(n_row_CSA_table_26):
        if i < 6:
            list_range_color_CSA26.append((i + 1, 1, i + 1, n_column_CSA_table_26, color_bkg_header))
        elif i % 2 == 0:
            list_range_color_CSA26.append((i + 1, 1, i + 1, n_column_CSA_table_26, color_bkg_data_1))
        else:
            list_range_color_CSA26.append((i + 1, 1, i + 1, n_column_CSA_table_26, color_bkg_data_2))

    # define cell format
    cell_format_CSA_26 = {
        #range_merge is used to merge cells with the format for instructions within the tuple list being: start_row (int), start_column (int), end_row (int), end_column (int), horizontal_align (str, optional) merged cell will be aligned: vertical centered, horizontal per spec
        'range_merge': ([(1, 1, 3, n_column_CSA_table_26, 'center'), (4, 1, 6, 1, 'center'), (4, 2, 4, 5, 'center'), (5, 2, 5, 3, 'center'), (5, 4, 5, 5, 'center')]),
        'range_font_bold' : [(1, 1, 2, 1)],
        'range_color': list_range_color_CSA26,
        'range_border': [(1, 1, 1, n_column_CSA_table_26), (2, 1, n_row_CSA_table_26, n_column_CSA_table_26)],
        'row_height': [(1, 50)],
        'column_width': [(i + 1, 30) for i in range(n_column_CSA_table_26)],
    }

    # define some footer notes
    footer_CSA26 = (['* This clearance may be reduced to 75 mm where adequate insulation is provided.',
                     'Voltages are rms line-to-ground.'])

    # define the worksheet
    CSA_Table_26 = {
        'ws_name': 'CSA Table 26',
        'ws_content': CSA26,
        'cell_range_style': cell_format_CSA_26,
        'footer': footer_CSA26
    }
#endregion

    #This determines the workbook and the worksheets within the workbook
    workbook_content = ([AEUC_Table_5, AEUC_Table_7, CSA_Table_2, CSA_Table_3, CSA_Table_5, CSA_Table_6, CSA_Table_7, CSA_Table_9, CSA_Table_10,
                         CSA_Table_11, CSA_Table_13, CSA_Table_14, CSA_Table_15, CSA_Table_16, CSA_Table_17, CSA_Table_18, CSA_Table_20, 
                         CSA_Table_21, CSA_Table_22, CSA_Table_23, CSA_Table_24, CSA_Table_25, CSA_Table_26])

    #This will create the workbook with the filename specified at the top of this function
    report_xlsx_general.create_workbook(workbook_content=workbook_content, filename=filename)
    return()

#this function is just here to run the def_report_excel function
if __name__ == '__main__':
    create_report_excel()
