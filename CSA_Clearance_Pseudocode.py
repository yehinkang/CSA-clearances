#Pseudocode for CSA Clearance Calc
 
"""

Inputs:
Create a class that takes all of these inputs
    P2P_Voltage
    Max_Overvoltage

    Design_Buffer_Non-Energised
    Design_Buffer_Energised

    Design_Buffer_2_Obs
    Design_Buffer_Same_Structure

    Clearance_Rounding

    Location

    Latitude_Northing
        Degrees_N
        Mins_N
        Seconds_N

    Latitude_Northing
        Degrees_W
        Mins_W
        Seconds_W

    Something that takes coordinates and spits out nearest location

    Loc_Elevation

    Custom_Elevation

    Table 17 Horizontal Conductor Seperations
        Span_Length
        Final_Unloaded_Sag_15C

    Crossing_or_Underbuild
        Is main wire upper or lower wire
        P2P_Voltage
        Max_Overvoltage


    
Tables:
define functions for all of these individual tables using the inputs to the clearance data class
AEUC Table 5 - Min. Vertical Design Clearances above Ground or Rails
AEUC Table 7 - Min. Design Clearance from Wires and Conductors Not Attached to Buildings, Signs, and Similar Plant
CSA Table 2 - Minimum Vertical Design Clearances above Ground or Rails, ac
CSA Table 3 - Minimum Vertical Design Clearances above Waterways*, ac
CSA Table 5 - Minimum Separations (heights) of Supply Equipment from Ground, ac
CSA Table 6 - Minimum Horizontal Desgin Clearances between Wires and Railway Tracks, ac
CSA Table 7 - Minimum Horizontal Separations from Supporting Structures to Railway Tracks, ac
CSA Table 9 - Minimum Design Clearances from Wires and Conductors not attached to Buildings, Signs, and similar Plant, ac
CSA Table 10 - Minimum Design Clearances from Supply Conductors to Bridges
CSA Table 11 - Minimum Separation of Equipment and Clearances of Conductors from Swimming Pools, ac
CSA Table 13 - Minimum Design Vertical Clearances between Wires Crossing Each Other and supported by Different Supporting Structures, ac
CSA Table 14 - Minimum Design Vertical Clearances for Crossings over Aerial Tramways
CSA Table 15 - Clearances between conductors supported by different structures but not crossing each other — Clearance increments to be added to horizontal displacement
CSA Table 16 - Minimum design clearances between conductors of one line and supporting structures of another line
CSA Table 17 - Minimum horizontal separations of supply-line conductors attached to the same supporting structure
CSA Table 18 - Minimum vertical separations between supply-line conductors attached to the same supporting structure, ac
CSA Table 20 - Minimum in-span vertical clearances between supply conductors of the same circuit that are attached to the same supporting structure
CSA Table 21 - Minimum in-span vertical clearances between supply conductors of different circuits that are attached to the same supporting structure
CSA Table 22 - Minimum separations or clearances in any direction from supply conductors to other supply plant attached to the same supporting structure
CSA Table 23 - Minimum vertical separations at a joint-use structure
CSA Table 24 - Minimum in-span vertical clearances between supply and communication conductors
CSA Table 25 - Minimum clearances from guys to plant of another system
CSA Table 26 - Minimum clearance or separation between guys and other plant attached to the joint-use structure


The code had an assortment of helper functions for rounding .etc
"""
