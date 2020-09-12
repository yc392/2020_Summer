import numpy as np
import math
from statistics import mean
import datetime



# declaration of constants
ca = 1.005        # specific heat of dry air; Kj/(kgºK)
cv = 1.85        # specific heat of water vapour; Kj/(kgºK)
cw = 4.186       # specific heat of water; Kj/(kgºK)
hv = 2500.8         # latent heat of vapourization of free water; Kj/(kgºK)
Pa = 101300        # atmospheric pressure; ??

# properties of moist air (ASAE 1997)
R1 = 22105649.25
a1 = -27405.526
b1 = 97.5413
c1 = -0.146244
d1 = 0.00012558
e1 = -0.000000048502
f1 = 4.34903
g1 = 0.0039381

# declare constants
pi = 3.14159265358979  # this is the pi number to compute are of the bin
PaToInWa = 0.0040146  # this is the conversion of 1 pa to inches of water
cfmTom3min = 0.0283168  # this is the conversion from 1 cfm to m3/min of airflow
HPtoKW = 0.746  # this is the conversion from 1 HP to KW of power
KjToKWH = 0.000278  # this is the conversion of 1 kJ to KWH
KWHtoBTU = 3412.141633  # this is the conversion of 1 KWH to Btu (IT)

# declare variables
class GRAIN:
    def __init__(self):
        self.ArrGrain = np.zeros((30, 15)).tolist()
        self.LineIndex = 0
        self.GrainIndex = 0
        self.FileName = ''
        self.TablePath = 0

# Form main
class AIR:
    def __init__(self):
        self.AirflowRate = 0  # Public AirflowRate As Float ' this is the average airflow rate, m3/min/t
        self.TotalAirflow = 0  # Public TotalAirflow As Float 'this is the total airflow rate in the bin, m3/min
        self.AirflowRate_cfm = 0  # Public AirflowRate_cfm As Float ' this is the average airflow rate, cfm/bu
        self.AirflowNonUnif = 0.15  # Public AirflowNonUnif As Float 'this is the non-uniformity factor to determine the center and side airflow rates, dec
        self.AirflowCenter = 0  # Public AirflowCenter As Float 'this is the airflow rate at the center of the bin, m3/min/t
        self.AirflowSide = 0  # Public AirflowSide As Float 'this is the airflow rate at the side of the bin, m3/min/t
        self.AirfResistance = 0  # Public AirfResistance As Float ' this is the resistance of the airflow, Pa
        self.AirfResistance_Wa = 0  # Public AirfResistance_Wa As Float ' this is the resistance of the airflow, inches of water
        self.AirVel_C = 0  # As Float   'this is the air velocity at the center of the bin, m/s
        self.AirVel_S = 0  # As Float   'this is the air velocity at the side of the bin, m/s
        self.AirTemp_C = 0  # As Float 'this is the temperature of the drying air, ºC
        self.AirRH = 0  # As Float 'this is the relative humidity of the drying air, %



class FAN:
    def __init__(self):
        self.FanStatus = True  # As Boolean 'this is the fan status, true when fan is "on" and false when fan is "off"
        self.FanPower_HP = 0 # As Float ' this is the estimated fan power, HP
        self.FanPower_KW = 0 #As Float ' this is the estimated fan power, KW
        self.FanEfficiency = 50 #Single ' this is the fan efficiency, %
        self.FanPreWarming_C = 0  # Public FanPreWarming_C As Float ' this is the estimated fan prewarming, ºC
        self.FanRunHours = 0  # As Float 'this is the total fan run hours of the current year simulation, hours
        self.Per_FanRun = 0  # As Float 'this is the % of fan run time of the current year simulation, %
        self.FanKWH = 0  # As Float 'this is the cumulative KWH (energy consumption) of the fan during the current simulation year, kwh
        self.FanPrint = ''  # As String ' this is the string with information about the fan status (F On or F Off) to be printed in the fan oputput file


class HEATER:
    def __init__(self):
        self.HeaterStatus = False  # As Boolean 'this is the burner status, true when burner is "on" and false when burner is "off"
        self.HeatingEnergy_KW = 0  # As Float ' this is the energy required to heat the air, KWH
        self.HeatingEnergy_BTU = 0  # As Float ' this is the energy required to heat the air, BTU
        self.HeaterRHflag = 0 #As Integer 'this flag indicates if the heater is on or off due to the ambient RH (0+off, 1=on)
        self.HeaterEMCflag = 0 #As Integer 'this flag indicates if the heater is on or off due to the ambient EMC (0+off, 1=on)
        self.HeaterFlag = 0 #As Integer 'this flag indicates if the heater is on or off due to the ambient coditions (0+off, 1=on)
        self.HeaterRunHours = 0  # As Float 'this is the total heater run hours of the current year simulation, hours
        self.Per_HeaterRun = 0  # As Float 'this is the % of heater run time of the current year simulation, %
        self.HeaterKWH = 0  # As Float 'this is the cumulative KWH (energy consumption) of the heater during the current simulation year, kw
        self.HeaterPrint = ''  # As String ' this is the string with information about the heater status (H On or H Off) to be printed in the fan oputput file
        self.BurnerEnergy_KW = 0  # Single 'this is the estimated burner power, KW
        self.BurnerEnergy_BTU = 0  # As Float 'this is the estimated burner power, BTU
        self.BurnerEfficiency = 80 #As Float ' this is the burner efficiency, %
        self.Tinc_C = 2 #As Float ' this is the temperature increase demanded to the burner, ºC



class BIN:
    def __init__(self):
        self.BinDiameter = 0 # As Float ' this is the diameter of the bin, m
        self.BinHeight = 0 #As Float 'this is the height of the bin, m
        self.BinArea = 0 #As Float ' this is the area of the bin, m2
        self.BinCapacity_t = 0 #As Float
        self.BinCapacity_bu = 0 #As Float
        self.PackingFactor = 1  # Public PackingFactor As Float ' packing factor for the Shedd's curves


class LAYER:
    def __init__(self, numLayer, layerDepth):
        #self.SingleLayerDepth1 = 0 #As Float ' initial estimate of layer depth, m
        self.SingleLayerDepth = layerDepth #As Float ' layer depth, m
        self.NumberOfLayers = numLayer #As Integer ' number of layers in the bin
        self.LayerDepth = [0] * numLayer #As Float ' detpth of each layer at begining of simulation (default value is 0.5m when simulatin starts), meter
        self.LayerDepth_C = [0] * numLayer #As Float 'depth of each grain layer at the center of the bin anytime after simulation started, m
        self.LayerDepth_S = [0] * numLayer #As Float 'depth of each grain layer at the side of the bin anytime after simulation started, m

class GrainInfo:
    def __init__(self, numLayer):
        self.AvgInGrainMC = 0 #As Float
        self.AvgInGrainTemp = 0 #As Float
        self.GrainMC_In_WB_C = [0] * numLayer #As Float ' MC of the grain in each layer at the center of the bin, %, wb
        self.GrainMC_In_WB_S = [0] * numLayer #As Float ' MC of the grain in each layer at the side of the bin, %, wb
        self.GrainTemp_In_C_C = [0] * numLayer #As Float ' Temp of the grain in each layer at the center of the bin, ºC
        self.GrainTemp_In_C_S = [0] * numLayer #As Float ' Temp of the grain in each layer at the side of the bin, ºC
        self.GrainMC_WB_C = [0] * numLayer #As Float 'this is the MC of the grain in each layer at the center of the bin at any time, %, wb
        self.GrainMC_WB_S = [0] * numLayer #As Float 'this is the MC of the grain in each layer at the side of the bin at any time, %, wb
        self.GrainTemp_C_C = [0] * numLayer #As Float 'this is the temperature of the grain in each layer at the center of the bin at any time, ºC
        self.GrainTemp_C_S = [0] * numLayer #As Float 'this is the temperature of the grain in each layer at the side of the bin at any time, ºC
        self.AvgGrainDensity = 0  # As Float ' average initial density of the grain, kg/m3 (used to compute airflow resistance)
        self.GrainDensity_C  = [0] * numLayer #As Float ' Desity of the grain in the layer at the center of the bin, kg/m3
        self.GrainDensity_S  = [0] * numLayer #As Float ' Desity of the grain in the layer at the side of the bin, kg/m3


class AirInfo:
    def __init__(self):
        self.Ps = 0 #As Float ' this is the saturated vapor pressure of the air
        self.Pv = 0 #As Float 'this is the vapor pressure of the air
        self.TC = 0 #As Float 'this is the temperature of the air in ºC
        self.TK = 0 #As Float 'this is the temperature of the air in ºK
        self.RH = 0 #As Float ' this is the RH of the air, %
        self.AbsHum = 0 #As Float ' this is the absolute humidity (H) of the air, g of water/kg of dry air

class DamageInfo:
    def __init__(self, numLayer):
        self.DML_Mult_Damage = 0 #As Float 'this is the DML multiplier for damage kernel
        self.DML_Mult_Fungicide = 0 #As Float ' this is the DML multiplier for fungicide application
        self.DML_Mult_Genetics = 0 #As Float 'this is the DML multiplier for hybrid
        self.DML_Mult_Temp = 0 #As Float 'this is the DML multiplier for temperature
        self.DML_Mult_MC = 0 #As Float 'this is the DML multiplier for MC
        self.GT = 0 #As Float 'this is the temperature of the grain for DML computation, ºC
        self.GMC = 0 #As Float 'this is the MC of the grain  for DML computation, %, wb
        self.GDamage = 0 #As Float ' this is the percentage of damage grain for DML computation, %
        self.tr = 0 #As Float ' this is the equivalent storage time for the DML computation, hours
        self.DML_GramsPerKg = 0 #As Float 'this is the predicted DML for the timestep, grams of CO2 produced per Kg of Dry Matter
        self.DML_BTUPerKg = 0 #As Float 'this is the amount of BUT produced when 1 gram of CO2 per KG of dry matter is realesed
        self.DML_TempIncr_C = 0 #As Float 'this is the temperature increase of the grain mass due to the heat generated by respiration (DML generation), ºC
        self.DMLFlag = '' #As String 'this flag indicates if the grain selected is corn, if DMLFlag is 1=true, then the grain is corn and DML will be computed
        self.LayerDML = [0] * numLayer # As Float 'this is the DML (%) of the layer of corn
        self.LayerDML_C = [0] * numLayer # As Float 'this is the DML (%) of the layer of corn at the center of the bin
        self.LayerDML_S = [0] * numLayer # As Float 'this is the DML (%) of the layer of corn at the side of the bin

SimStatus = 0 #As Boolean   'this is the status for the simulation, when it is true, simulation will run for next timestep, if it is false it will stop at current time step

class DryingInfo():
    def __init__(self, numLayer):
        self.dx = [0] * numLayer # As Float        ' depth of the thin layer; m
        self.dxf = [0] * numLayer # As Float        ' depth of the thin layer at the end of the timestep
        self.MfW = [0] * numLayer # As Float     ' final grain moisture content; %, wb
        self.GfC = [0] * numLayer # As Float     ' final grain temperature, ºC
        self.va = 0 #Dim va As Float             ' velocity of the drying air, m/s
        self.M0W = [0] * numLayer #Dim grainInfo.GrainMC_WB_C(31) As Float     ' initial grain moisture content; %, wb
        self.G0C  = [0] * numLayer ##Dim G0C(31) As Float     ' initial grain temperature; ºC

class OutputFile:
    def __init__(self, numLayer):
        self.CurrentDir = ''  # As String ' this string carries the information about the location of the VB-PHAST program
        self.BaseName = ''  # As String ' this is the basename of the output files
        self.FileName1 = '' #As String ' this is the filename of the file in which the output will be written
        self.FileNumber = 0 #As Integer 'this is the filenumber indicator of a given open file, represents the number in the open file FileName as #1
        self.PrintVar = [0] * numLayer # As Float 'this is the variable to print in the file
        self.TimeStamp = '' #As String ' this is the time stamp (hour, i.e.; 1:00) for the outputs files
        self.DateStamp = '' #As String 'this is the date stamp (month-day-year) for the output files
        self.PrintTime = '' #As String 'this is the time stamp and date stamp combined for the outputs files
        self.PrintHeading = '' #As String 'this is the heading string for the output MC, Temp and DML files
        self.PrintHeading_Fan = '' #As String 'this is the heading string for the output Fan file
        self.PrintHeading_End = '' #As String 'this is the heading string for the output Summary file
        self.PrintVar_Fan = '' #As String ' this is the string that carries all the variables for each hour to be printed in the fan output file
        self.PrintVar_End = '' #As String ' this is the string that carries all the variables for each year to be printed in the Summary output file
        self.PrintYear = '' #As String 'this is the Year stamp for the Summary outputs files

#'variabled associated to the subroutine StopSimulation
class StopCriteria:
    def __init__(self,numLayer):
        self.MCCriteria = False #As Boolean ' criteria to consider simulation completed, when MCCriteria = true, then MC is the selected criteria
        self.TempCriteria = False #As Boolean ' criteria to consider simulation completed, when TempCriteria = true, then temperature is the selected criteria
        self.DMLCriteria = False #As Boolean ' criteria to consider simulation completed, when DMLCriteria = true, then DML is the selected criteria
        self.DateCriteria = False #As Boolean ' criteria to consider simulation completed, when DateCriteria = true, then Date is the selected criteria
        self.ArrCriteriaC = [0] * numLayer # As Float ' this is the array that carries the information from the center of the bin to decide if simulation was completed or not
        self.ArrCriteriaS = [0] * numLayer # As Float ' this is the array that carries the information from the side of the bin to decide if simulation was completed or not
        self.ArrCriteria = [0] * (2 * numLayer) # As Float ' this is the array that carries the information to decide if simulation was completed or not
        self.CurrentHours = 0 #As Integer ' this is the variable that carries the information about the current number of hours since simulation started
        self.StopSim = False #As Boolean 'this is the output of the sub, when true simulation was completed
        self.MaxLayer = 0 #As Float 'this is the number of layers considered in the simulation, used to determine the upper bound of the array
        self.MCavg_Stop = 0 #As Float 'this is the desired average final moisture content value to consider completed the simulation
        self.MCmax_Stop = 0 #As Float ' this is the desired maximum final moisture content value to consider completed the simulation
        self.Tempavg_Stop = 0 #As Float 'this is the desired average final temperature value to consider completed the simulation
        self.Tempmax_Stop = 0 #As Float ' this is the desired maximum final temperature value to consider completed the simulation
        self.DMLavg_Stop = 0 #As Float 'this is the desired average final DML value to consider completed the simulation
        self.Hours_Stop = 0 #As Integer 'this is the desired number of hours to consider completed the simulation

#'variables associated with the simulation time period
class SimTime:
    def __init__(self):
        self.InitialSimYear = 2019 #As Float 'this is the year at which the first simulation will start
        self.InitialSimMonth = 4 #As Float 'this is the month at which the first simulation will start
        self.InitialSimDay = 11 #As Float 'this is the day at which the first simulation will start
        self.FinalSimYear = 2019 #As Float 'this is the year at which the first simulation will finish
        self.FinalSimMonth = 4 #As Float 'this is the month at which the first simulation will finish
        self.FinalSimDay = 13 #As Float 'this is the day at which the first simulation will finish

#'variables associated with the multiple years simulation period
class MultYear:
    def __init__(self):
        self.IMultYear = 0 #As Float 'this is the first year of the multiple years simulation
        self.FMultYear = 0 #As Float 'this is the last year of the multiple years simulation
        self.MultYears = False #As Boolean ' this variable indicates if mulatiple years simulation is requested

#'varaibles asociated with the weather file
class FileSimTime:
    def __init__(self):
        self.InitialFileYear = 0 #As Integer 'this is the first year in the weather file
        self.InitialFileMonth = 0 #As Integer 'this is the first month in the weather file
        self.InitialFileDay = 0 #As Integer 'this is the first day in the weather file
        self.FinalFileYear = 0 #As Integer 'this is the last year in the weather file
        self.FinalFileMonth = 0 #As Integer 'this is the last month in the weather file
        self.FinalFileDay = 0 #As Integer 'this is the last day in the weather file



#'variables associated with the location of the line to read in the weather file
class WFILEInfo:
    def __init__(self):
        self.TempColumn = 0  # As Integer 'this is the number of the temperature column in the weather file
        self.RHColumn = 0  # As Integer 'this is the number of the RH column in the weather file
        self.TotalDaysToFile = 0 #As Long ' number of days from January 1 of 1960 to the first year month and day of the weather file
        self.TotalDaysStartSim = 0 #As Long 'number of days from January 1 of 1960 to the first year month and day of the analysis
        self.DaysToStart = 0 #As Long 'number of days since the beginin of the weather data to the begining of the analysis
        self.FirstLineRead = 0 #As Long 'number of hours (lines) between the begining of the weather file and the first simulation hour
        self.TotalDaysFinishSim = 0 #As Long 'days from January 1 of 1960 to the last year month and day of the analysis
        self.DaysToFinish = 0 #As Long 'number of days since the beginin of the weather data to the last day of the analysis
        self.LastLineRead = 0 #As Long 'number of hours (lines) between the begining of the weather file and the last simulation hour
        self.CurrentYear = 0 #As Long 'this is the current year of the simulation
        self.FileNameW = 0 #As String 'this is the filename and path of the weather file used for the simulation
        self.LineFromFile = 0 #As String 'this string is the line readed from the weather file
        self.sItems = '' #As String 'thi  is the string that will contain the information read from each line of the weather file


#'varaibles associated with the selecting criteria of useful hours for the CNA strategy
class CNALimit:
    def __init__(self):
        self.MinTempSelect = -30 #As Float 'this is the minimum temperature limit for the temperature window of the CNA strategy, ºC
        self.MaxTempSelect = 40 #As Float 'this is the maximum temperature limit for the temperature window for the CNA strategy, ºC
        self.MinRHSelect = 0 #As Float 'this is the minimum RH limit for the RH window of the CNA strategy, %
        self.MaxRHSelect = 100 #As Float 'this is the maximum RH limit for the RH window of the CAN strategy, %
        self.MinEMCSelect = 0 #As Float 'this is the minimum drying EMC limit for the EMC window of the CNA strategy, %, wb
        self.MaxEMCSelect = 30 #As Float 'this is the maximum drying EMC limit for the EMC window of the CNA strategy, %, wb

#'variables associated with weather data
class WeatherData:
    def __init__(self):
        self.RHflag = 0 #As Integer ' this is the flag that indicates that the RH read from the file fits into the selected window of fan operation, 1 = true, 0= false
        self.Tempflag = 0 #As Integer ' this is the flag that indicates that the temperature read from the file fits into the selected window of fan operation, 1 = true, 0= false
        self.EMCflag = 0 #As Integer ' this is the flag that indicates that the EMC computed from Temp and RH read from the file fits into the selected window of fan operation, 1 = true, 0= false
        self.BadDataflag = 0 #As Integer ' this is the flag that indicates that the temperature and RH data read from the file are good (tempe >-29ºC and <40ºC and RH >1 and <=100%), 1 = true, 0= false
        self.HourFlag = 0 #As Integer 'this is the flag that indicates that the temperature,RH , EMC and BadData flags are true, 1=true, 0=false
        self.FileTemp = 0 #As Float 'this is the ambient temperature data read from the weather file, ºC
        self.FileRH = 0 #As Float 'this is the ambient RH data read from the weather file, %
        self.FileEMC_db = 0 #As Float 'this is the ambient EMC data computed from the ambient T and RH read from the weather file, %, db
        self.FileEMC_wb = 0 #As Float 'this is the ambient EMC data computed from the ambient T and RH read from the weather file, %, wb

#'variables associated with the drying air condition at the plenum of the bin
class PLENUM:
    def __init__(self):
        self.PlenumTemp_C  = 0 #As Float ' this is the plenum air drying temperature, ºC
        self.PlenumRH = 0 #As Float ' this is the plenum air RH, %
        self.PlenumRHfan = 0 #As Float ' this is the rh in the plenum after the fan prewarming, %
        self.PsAmbient = 0 #As Float ' this is the Vapor Pressure at Saturation of the ambient air
        self.PvAmbient = 0 #As Float ' this is the Vapor Pressure of the ambient air
        self.PsPlenum = 0 #As Float ' this is the Vapor Pressure at Saturation of the drying air at the plenum
        self.PvPlenum = 0 #As Float ' this is the Vapor Pressure of the drying air at the plenum
        self.PlenumEMC_db = 0 #As Float 'this is the plenum EMC data computed from the plenum T and RH, %, db
        self.PlenumEMC_wb = 0 #As Float 'this is the plenum EMC data computed from the plenum T and RH, %, wb


#'variables related to the simulation of a set of years
class YearSummery:
    def __init__(self, numYear):
        self.S_Moisture_Avg = [0] * numYear # As Float 'this array constain the final average moisture content for each one of the years of the simulation, %, wb
        self.S_Moisture_Min = [0] * numYear #  As Float 'this array constain the final minimum moisture content for each one of the years of the simulation, %, wb
        self.S_Moisture_Max = [0] * numYear #  As Float 'this array constain the final maximum moisture content for each one of the years of the simulation, %, wbPublic S_Temperature_Avg(50) As Float 'this array constain the final average temperature for each one of the years of the simulation,ºC
        self.S_Temperature_Min = [0] * numYear #  As Float 'this array constain the final minimum temperature for each one of the years of the simulation,ºC
        self.S_Temperature_Max = [0] * numYear #  As Float 'this array constain the final maximum temperature for each one of the years of the simulation,ºC
        self.S_DML_Avg = [0] * numYear #  As Float 'this array constain the final average DML for each one of the years of the simulation, %
        self.S_DML_Max = [0] * numYear #  As Float 'this array constain the final maximum DML for each one of the years of the simulation, %
        self.S_DryingHs = [0] * numYear #  As Float 'this array constain the final drying hours for each one of the years of the simulation, hours
        self.S_FanHs = [0] * numYear #  As Float 'this array constain the final fan run hours for each one of the years of the simulation, hours
        self.S_PerFanHs = [0] * numYear #  As Float 'this array constain the final percentaje of fan run hours for each one of the years of the simulation, %
        self.S_HeaterHs = [0] * numYear #  As Float 'this array constain the final heater run hours for each one of the years of the simulation, hours
        self.S_PerHeaterHs = [0] * numYear #  As Float 'this array constain the final percentaje of heater run hours for each one of the years of the simulation, %
        self.S_FanKWH = [0] * numYear #  As Float 'this array constain the final energy consumption of the fan for each one of the years of the simulation, KWH
        self.S_HeaterKWH = [0] * numYear #  As Float 'this array constain the final energy consumption of the heater for each one of the years of the simulation, KWH
        self.S_OutputSummary = False #As Boolean 'this variable indicates if only Summary results are required, or all results are requires (true=summary, false=all)
        self.S_TotDryingCost = [0] * numYear # As Float 'this is the drying cost for each simulation run, $/tonne




SAVH_FinalMC = 0 #As Float ' this is the desired final moisture content for the savh strategy, %, wb
EstDryingTime = 0 #As Float ' this is the estimated drying time for the savh strategy, % drying time = 15*50/cfm/bu
MC_Highest = 0 #As Float ' this is the highest MC of the first grain layer at any given hour, %, wb
MC_Lowest = 0 #As Float ' this is the lowest MC of the first grain layer at any given hour, %, wb
MC_HighLimit = 0 #As Float ' this is the high MC limit at any given hour, %, wb
MC_LowLimit = 0 #As Float ' this is the low MC limit at any given hour, %, wb

#'variables related to the computation of the drying cost
class Cost:
    def __init__(self):
        self.GrainPrice = 0 #As Float 'this is the price of the grain, in $/tonne
        self.ElectricityCost = 0 # As Float 'this is the cost of electricity, in $/kwh
        self.PropaneCost = 0 # As Float 'this is the cost of propane (for the burner), in $/gallon
        self.EnergyCost = 0 # As Float 'this is the cost of the energy required by the Fan and the Heater, $
        self.ShrinkCost = 0 # As Float 'this is the cost of overdrying (from the desired final MC) and DML,$
        self.DesiredFinMC = 0 # As Float 'this is the desired final moisture content (%, wb) to conpute the shrink cost
        self.FanCost = 0 # As Float 'this is the total cost of the fan, $
        self.HeaterGalons = 0 # As Float 'this is the total gallons of propane consummed (3414 BTU/KW and 92000 BTU/Gallon), gallons
        self.HeaterCost = 0 # As Float 'this is the total cost of the heater, $
        self.DMLCost = 0 # As Float 'this is te cost related to DML, $
        self.DMLBin = 0 # As Float 'this is the average dml of the bin used for cost calculations, %
        self.OverdryingCost = 0 # As Float ' this is the cost related to the overdrying of the grain, $
        self.HeaterType = 0 # As Boolean 'this is the type of heater considered, true=propane gas and false=electrical
        self.FinalTonGrain = 0 # As Float 'this is the final tons of grains at the desired final MC
        self.InTonnes = 0 #As Float 'thi is the initial number of tonnes of grain in the bin, used for drying cost computations
        self.AvgFinMC = 0 #As Float 'this is the average final mc of the grain, used for drying cost computation

class StringInfo:
    def __init__(self):
        self.StWeatherFile = ''  # As String 'this string carries the information about the weather file used in the simulation
        self.StStrategy = ''  # As String 'this string carries the information about the strategy selected for the simulation
        self.StBin = ''  # As String 'this string carries the information about the dimensions of the bin used for the simulation
        self.StGrainEMC = ''  # As String 'this string carries the information about the EMC parameters used in the simulation
        self.StAirflow = ''  # As String 'this string carries the information af the airflow rate of the simulation
        self.StGrain = ''  # As String 'this string carries the information aboput the initial grain condition (MC and T)
        self.StStartDate = ''  # As String 'this string carries the information about the start date of the simulation (month and day)
        self.StStopSimulation = ''  # As String 'this string carries the information about the stop criteria for the simulation
        self.StRunInfo = ''  # As String ' this string combines the information of the 8 strings above



def ReadGrainLIst(FileName):
    # Using readlines()
    file1 = open(FileName, 'r')
    Lines = file1.readlines()

    newgrain = GRAIN()
    count = 0
    # Strips the newline character
    for line in Lines:
        j = 0
        linesplit = line.split('\t')
        linesplit.pop()
        for each in linesplit:
            if j == 0:
                newgrain.ArrGrain[count][j] = each
            else:
                newgrain.ArrGrain[count][j] = float(each)
            j += 1
        count += 1
    file1.close()
    return newgrain



def DisplayGrainInfo(GrainIndex):
    return
    #print(ArrGrain[GrainIndex])

def UpdateF_GrainEdit(editIndex):
    """if editIndex == 0:
        # add
    if editIndex == 1:
        #edit
    if editIndex == 2:
        #view
    """


def SetGrainLayerDepth(Layer):
#' this sub is to set the initial grain layer depth
    for i in range(Layer.NumberOfLayers):
        Layer.LayerDepth[i] = Layer.SingleLayerDepth

    #LayerDepth[Layers] = Height - int(Height)

def ComputeGrainDensity(GrainIndex, GrainMC1, ArrGrain):
#this function computes the density of the grain (as function of grainMC) based on the ASAED241.4 standard
#Grain MC is the moisture content of the grin, dec, w.b
#Grain density is the density of the grain in kg/m3
    return ArrGrain[GrainIndex][7] - ArrGrain[GrainIndex][8] * (GrainMC1 / 100) + ArrGrain[GrainIndex][9] * (GrainMC1 / 100) * (GrainMC1 / 100)


def AirFlowResistance(GrainIndex, ArrGrain, TotalAirflow, BinArea, PackingFactor):

#this function computes the pressure drop per meter according to a given airflow rate (m3/m2/sec)
# for a specific grain, Pa/m
    return (ArrGrain[GrainIndex][12] * ((TotalAirflow / 60 / BinArea) * (TotalAirflow / 60 / BinArea)) / math.log(1 + ArrGrain[GrainIndex][13] * (TotalAirflow / 60 / BinArea))) * PackingFactor


def CompFanPreWarm(AirfResistance_Wa):
    # this function computes the fan prewarming based on airflow resistance in inches of water
    # fan prewarming is in ºC, and assumes that 0.28ºC of temp increase for each inch of water of
    # static pressure
    return 0.28 * AirfResistance_Wa

def CompFanPower(TotalAirflow, AirfResistance_Wa, FanEfficiency):
    #this function computes the fan power required for the given airflow rate (cfm) and the given
    #static pressure (inch of water) for the given fan efficiency.
    #the output is in HP
    return (TotalAirflow / cfmTom3min * AirfResistance_Wa) / (63.46 * FanEfficiency)

def Kelvin(TC):
#'declare functions
#' to convert ºC to ºK
#' TC is temperature in ºC
#' TK is temperature in ºK
    return TC + 273.15

def Sat_press(TK):
# compute saturated vapour pressure
# TK is temperature, ºK
# Ps is saturated vapour pressure
    return R1 * (math.exp((a1 + b1 * TK + c1 * (math.pow(TK,2)) + d1 * (math.pow(TK,3)) + e1 * (math.pow(TK,4))) / (f1 * TK - g1 * (math.pow(TK,2)))))


def CompVaporPress(Ps, RH):
#'compute the vapor pressure of the air
#'Ps is the saturated vapor pressure of the air
#'RH is the relative humidity of the air, %
    return Ps * RH / 100

def CompAirDens(Pv, TK):
    #' compute air density in kg/m3
    #'PV is the vapor pressure of the air
    #'TK is the temperature of the air, ºK
    return (Pa - Pv) / (287 * TK)


def CompAbsHum(Pv):
#'compute asbsolute humidity (H) of the air, g of water/kg of dry air
#'Pv is the vapor pressure of the air
#'Pa is the atmospheric pressure of the air
    return 0.6219 * Pv / (Pa - Pv)


def CompGrainSpHeat(ArrGrain, GrainIndex, GrainMC):
#'compute the specific heat of the grain based on the grain moisture content according to ASAE D243.4
#' grain specific heat is in kJ/(kg*ºK)
#' grain MC is in %, w.b
    return ArrGrain[GrainIndex][10] + ArrGrain[GrainIndex][11] * GrainMC



def CompBurnerPower(Tinc_C, TC, RH, TotalAirflow):
#'compute the power of the heater based on the temperature increase and the airflow
#'TC is the temperature of the air, ºC
    global TK,Ps,Pv,AbsHum
    TK = Kelvin(TC)
    Ps = Sat_press(TK)
    Pv = CompVaporPress(Ps, RH)
    AbsHum = CompAbsHum(Pv)
    return ((ca + cv * AbsHum) * (Tinc_C) * (TotalAirflow * 60 * CompAirDens(Pv, TK))) * KjToKWH


def CompTempMult(GT, GMC):
#'compute the Temperature Multiplier for the DML equation
#'the temperature multiplier is computed based on the ASAE Standard (X535) in revision (12-22-04)
#'GT is grain temperature, C
#'GMC is grain moisture content, %, wb
    if GT < 15.6:
        CompTempMult = 128.389 * math.exp(-4.86 * (1.8 * GT + 32) / 60)
    elif GT < 26.7:
        if GMC < 19:
            CompTempMult = 32.3 * math.exp(-3.48 * (1.8 * GT + 32) / 60)
        elif GMC < 28:
            CompTempMult = math.exp(-0.00493277 + (0.05 * (1.8 * GT + 32) - 3) * (math.log(0.0795012 + 0.012315 * GMC)))
        else:
            CompTempMult = math.exp(2.56683 - 0.0428628 * (1.8 * GT + 32))
    else:
        if GMC < 19:
            CompTempMult = 32.3 * math.exp(-3.48 * (1.8 * GT + 32) / 60)
        elif GMC < 28:
            CompTempMult = 32.3 * math.exp(-3.48 * (1.8 * GT + 32) / 60 + ((GMC - 19) / 100 * math.exp(0.61 * (1.8 * GT - 28) / 60)))
        else:
            CompTempMult = 32.3 * math.exp(-3.48 * (1.8 * GT + 32) / 60) + 0.09 * math.exp(0.61 * (1.8 * GT - 28) / 60)

    return CompTempMult


def CompMoistMult(GMC):
#'compute the Moisture Multiplier for the DML equation
#'the moisture multiplier is computed based on the ASAE Standard (X535) in revision (12-22-04)
#GMC is grain moisture content, %, wb
    return 0.103 * (math.exp(455 / ((GMC * 100 / (100 - GMC)) ^ 1.53)) - (0.00845 * (GMC * 100 / (100 - GMC))) + 1.558)

def CompDamageMult(GDamage):
#'compute the Damage Multiplier for the DML equation
#'the damage multiplier is computed based on the ASAE Standard (X535) in revision (12-22-04)
#'GDamage is the percentage of damaged grain, %
    return 2.08 * math.exp(-0.0239 * GDamage)


def CompEqStorageTime(GT, GMC, GDamage, DML_Mult_Fungicide, DML_Mult_Genetics):
#' compute the equivalent storage time for the DML equation
#'the equivalent storage time is computed on the basis of the Temperature, MC, grain Damage, grain Gentetics and Fungicide Multipliers
#'the results is in hours
    global DML_Mult_Temp, DML_Mult_MC, DML_Mult_Damage
    DML_Mult_Temp = CompTempMult(GT, GMC)
    DML_Mult_MC = CompMoistMult(GMC)
    DML_Mult_Damage = CompDamageMult(GDamage)
    return 1 / (DML_Mult_Temp * DML_Mult_MC * DML_Mult_Damage * DML_Mult_Fungicide * DML_Mult_Genetics)


def CompCO2Prod(GT, GMC, GDamage, DML_Mult_Fungicide, DML_Mult_Genetics):
#'compute the CO2 production of the stored corn
#'compute the grain DML in grams of CO2 produced per kilogram of initial dry matter
#' CompCO2Prod is computed according to the procedure proposed by Saul and Steel (1966)
    tr = CompEqStorageTime(GT, GMC, GDamage, DML_Mult_Fungicide, DML_Mult_Genetics)
    return 1.3 * (math.exp(0.006 * tr) - 1) + 0.015 * tr



def CompDML_TempInc(DML_GramsPerKg, ArrGrain, GrainIndex, MC_WB):
#' compute the temperature increase (ºC) in the grain mass due to the heat generated during respiration
#'the computing is based on the following relationship:
#'respiration of 1 mol of glucose (180 g) produces: 264 g of CO2; 108 g of water; and 2816 kJoules
#'based on this relationship, per gram of glucose respired: 1.47 g of CO2; 0.6 g of water; and 15.64 kJoules
    return (15.64 * DML_GramsPerKg) / (ArrGrain[GrainIndex][10] + ArrGrain[GrainIndex][11] * MC_WB)


def CF_EMC_D(TC, RH, ArrGrain, GrainIndex) :
#' compute the desorption (drying) EMC with the modified Chung-Pfost equation
#' TC is temperature in ºC
#' RH is the air relative humidity, decimal
#' EMC is the equilibrium moisture content, %, db
    if TC < -30:
        TC = -30
#'chung-pfost model does not work with temperaures below -30ºC
    return (math.log(math.log(RH) * -1 * ((TC + ArrGrain[GrainIndex][3]) / ArrGrain[GrainIndex][1]))) / -ArrGrain[GrainIndex][2]


def CF_EMC_R(TC, RH, ArrGrain, GrainIndex) :
    # ' compute the adsorption (re-wetting) EMC with the modified Chung-Pfost equation
    # ' TC is temperature in ºC
    # ' RH is the air relative humidity, decimal
    # ' EMC is the equilibrium moisture content, %, db
    if TC < -30:
        TC = -30 #'chung-pfost model does not work with temperaures below -30ºC
    return (math.log(math.log(RH) * -1 * ((TC + ArrGrain[GrainIndex][6]) / ArrGrain[GrainIndex][4]))) / -ArrGrain[GrainIndex][5]


def CF_ERH_D(TC, MC, ArrGrain, GrainIndex) :
#' compute the desorption (drying) ERH with the modified Chung-Pfost equation
#' TC is temperature in ºC
#' RH is the equilibrium relative humidity, decimal
#' MC is the grain moisture content, %, db
    if TC < -30:
        TC = -30 #'chung-pfost model does not work with temperaures below -30ºC
    return math.exp(-(ArrGrain[GrainIndex][1] / (TC + ArrGrain[GrainIndex][3])) * math.exp(-ArrGrain[GrainIndex][2] * MC))


def CF_ERH_R(TC, MC, ArrGrain, GrainIndex):
#' compute the adsorption (re-wetting) ERH with the modified Chung-Pfost equation
# ' TC is temperature in ºC
# ' RH is the equilibrium relative humidity, decimal
# ' MC is the grain moisture content, %, db
    if TC < -30:
        TC = -30 #'chung-pfost model does not work with temperaures below -30ºC
    return math.exp(-(ArrGrain[GrainIndex][4] / (TC + ArrGrain[GrainIndex][6])) * math.exp(-ArrGrain[GrainIndex][5] * MC))



# ' convert moisture content from wet basis to dry basis, decimal
# ' Mwb: moisture content wet basis, decimal
# ' Mdb: moisture content dry basis, decimal
def Mwb_Mdb(Mwb) :
    return Mwb / (1 - Mwb)

# ' convert moisture content from dry basis to wet basis, decimal
# ' Mwb: moisture content wet basis, decimal
# ' Mdb: moisture content dry basis, decimal
def Mdb_Mwb(Mdb) :
    return Mdb / (1 + Mdb)


#' this function is to adjust the grain layer depth according to the change in moisture content during the timestep
def CompLayerDepthF(GrainIndex, MCInitial, MCFinal, ArrGrain, LayerDepth0):
    DensityI = 0  #'this is the density of the cor at the begining of the time step
    CornMassI = 0 #'this is the mass of corn at MC initial
    DryMassCorn = 0 #'this is the dy mass of corn in the layer
    DensityF = 0 #'this is the density of the corn at the end of the time step (at MC final)
    CorMassF = 0 # 'this is the mass of corn at MC final
#'CompLayerDepth is the depth of the corn layer at the end of the time step, adjusted by change in MC and grain density
    DensityI = ArrGrain[GrainIndex][7] - ArrGrain[GrainIndex][8] * (MCInitial / 100) + ArrGrain[GrainIndex][9] * (MCInitial / 100) * (MCInitial / 100)
    CornMassI = DensityI * LayerDepth0
    DryMassCorn = CornMassI * (1 - MCInitial / 100)
    DensityF = ArrGrain[GrainIndex][7] - ArrGrain[GrainIndex][8] * (MCFinal / 100) + ArrGrain[GrainIndex][9] * (MCFinal / 100) * (MCFinal / 100)
    CornMassF = DryMassCorn / (1 - MCFinal / 100)
    return CornMassF / DensityF


# def DryingCost(FanKWH, HeaterKWH, InTonnes, DMLBin, AvgFinMC):
# #'this sub is to compute the total drying cost for the current yer (Ricardo Bartosik Master thesis)
# #'drying cost=energy cost + shrink cost
#     FanCost = FanKWH * ElectricityCost
#     if HeaterType == True:
#         HeaterCost = HeaterKWH * 3414 / 92000 * PropaneCost
#     else:
#         HeaterCost = HeaterKWH * ElectricityCost
#
#     EnergyCost = FanCost + HeaterCost
#
#     DMLCost = InTonnes * DMLBin / 100 * GrainPrice
#     FinalTonGrain = InTonnes * (1 - AvgInGrainMC / 100) / (1 - DesiredFinMC / 100)
#     if AvgFinMC < DesiredFinMC:
#         OverdryingCost = (FinalTonGrain - (InTonnes * (1 - AvgInGrainMC / 100) / (1 - AvgFinMC / 100))) * GrainPrice
#     else:
#         OverdryingCost = 0
#
#     ShrinkCost = DMLCost + OverdryingCost
#
#     return (EnergyCost + ShrinkCost) / FinalTonGrain


# 'declare procedure
# ' To determine if a year is "bisiesto"
# ' Year is the year that questioned
# ' If Year is Bisiesto, then Bisiesto1 = 1
# def Bisiesto1(Year):
#     if (int(Year / 4) == (Year / 4)):
#         return 1
#     else:
#         return 0





def Drying(Grain, Layer, grainInfo, AirTemp_C, AirRH, AirVel, fanPara, center=True):
#grainInfo.GrainMC_WB_C, G0C, AirTemp_C, AirRH, dx, ArrGrain, GrainIndex, va, FanStatus

# ' Comments at 8/5/2004 by REB
# ' This drying model was develop on the basis of the Thompson et al., 1972 equilibrium model. Adapted from Romualdo Martinez (2001)
# ' thesis: Modelling and Simulation of the Two Stage Rice Drying System in the Philippines,
# ' Hohenheim, 2001.




    T0C = [0] * Layer.NumberOfLayers # As Float     ' initial air temperature;
    T0K = [0] * Layer.NumberOfLayers # As Float     ' initial air temperature;
    RH0 = [0] * Layer.NumberOfLayers # As Float     ' initial air relative humidity; %
    Ps0 = [0] * Layer.NumberOfLayers # As Float     ' initial saturated vapour pressure; ??
    Pv0 = [0] * Layer.NumberOfLayers # As Float     ' initial vapour pressure at given air condition; ??
    H0 = [0] * Layer.NumberOfLayers # As Float      ' initial absolute humidity of the air; kg/kg

    M0d = [0] * Layer.NumberOfLayers # As Float     ' initial grain moisture content; %, db
    M0d1 = [0] * Layer.NumberOfLayers # As Float    ' grain moisture content at the beginning of the simulation; %, db

    G0K = [0] * Layer.NumberOfLayers # As Float     ' initial grain temperature;
    R = [0] * Layer.NumberOfLayers # As Float       ' dry mater to dry air ratio; kg/kg
    GD = [0] * Layer.NumberOfLayers # As Float      ' grain density; kg/m3
    ad = [0] * Layer.NumberOfLayers # As Float      ' air density; kg/m3
    Cg1 = [0] * Layer.NumberOfLayers # As Float     ' specific heat of corn; kJ/(kg)
    Cg = [0] * Layer.NumberOfLayers # As Float      ' specific heat of corn in relation to the dry mater to dry air ratio; kJ/(kg)
    Hf1 = [0] * (2*Layer.NumberOfLayers) # As Float     ' absolute humidity used to find feasible final RH conditions
    Mfd1 = [0] * (2*Layer.NumberOfLayers) # As Float
    TfC1 = [0] * (2*Layer.NumberOfLayers) # As Float
    TfK1 = [0] * (2*Layer.NumberOfLayers) # As Float
    Psf1 = [0] * (2*Layer.NumberOfLayers) # As Float
    Pvf1 = [0] * (2*Layer.NumberOfLayers) # As Float
    RHf1A = [0] * (2*Layer.NumberOfLayers) # As Float
    RHf1B = [0] * (2*Layer.NumberOfLayers) # As Float
    dRH1 = [0] * (2*Layer.NumberOfLayers) # As Float

    Mfd = [0] * Layer.NumberOfLayers # As Float       ' final grain moisture content; %, db
    dH = [0] * Layer.NumberOfLayers # As Float      ' change in absolute humidity of the air; kg/kg
    Hf = [0] * Layer.NumberOfLayers # As Float      ' final absolute humidity of the air; kg/kg
    Rgas = [0] * Layer.NumberOfLayers # As Float    ' R constant of water vapour at final conditions; kJ/(kg篕)
    dL = [0] * Layer.NumberOfLayers # As Float      ' Latent heat of vaporization of water; KJ/(kg篕)
    TfC = [0] * Layer.NumberOfLayers # As Float     ' final air temperature; 篊
    TfK = [0] * Layer.NumberOfLayers # As Float     ' final air temperature; 篕
    Psf = [0] * Layer.NumberOfLayers # As Float     ' final Ps
    Pvf = [0] * Layer.NumberOfLayers # As Float     ' final Pv
    RHf = [0] * Layer.NumberOfLayers # As Float     ' final RH, %

    Process = 0 #As Integer      'indicates if this is a drying or rewetting process: 0= drying, 1=rewetting, 2= average between drying and rewetting
    TimeStepEMC_D = 0 #As Float   'this is the computed drying EMC for the timested, used to determine if during the current timestep a drying or rewetting equation should be used
    TimeStepEMC_R = 0 #Single   'this is the computed rewetting EMC for the timested, used to determine if during the current timestep a drying or rewetting equation should be used
    t = 1 #'timestep is 1 hour

    dryingInfo = DryingInfo(Layer.NumberOfLayers)
    if center:
        dryingInfo.M0W = grainInfo.GrainMC_WB_C
        dryingInfo.G0C = grainInfo.GrainTemp_C_C
        dryingInfo.dx = Layer.LayerDepth_C
        dryingInfo.va = AirVel
    else:
        dryingInfo.M0W = grainInfo.GrainMC_WB_S
        dryingInfo.G0C = grainInfo.GrainTemp_C_S
        dryingInfo.dx = Layer.LayerDepth_S
        dryingInfo.va = AirVel


    damageInfo = DamageInfo(Layer.NumberOfLayers)

    #i = 0    ' set array index to 1 to start in the first thin layer
    for i in range(Layer.NumberOfLayers):

        if i == 0:   #' to set the T0 to ambient for thin layer 1, or to Tf of the layer before
            T0C[i] = AirTemp_C
            RH0[i] = AirRH
        else:
            T0C[i] = TfC[i - 1]
            RH0[i] = RHf[i - 1]


        #' to compute moisture and temperature change only if fan is on
        if fanPara.FanStatus:
            M0d[i] = Mwb_Mdb(dryingInfo.M0W[i] / 100) * 100
            T0K[i] = Kelvin(T0C[i])
            Ps0[i] = Sat_press(T0K[i])
            Pv0[i] = CompVaporPress(Ps0[i], RH0[i]) #' compute Pvs
            H0[i] = CompAbsHum(Pv0[i])  #' compute absolute humidity 0
            GD[i] = Grain.ArrGrain[Grain.GrainIndex][7] - Grain.ArrGrain[Grain.GrainIndex][8] * (dryingInfo.M0W[i] / 100) + Grain.ArrGrain[Grain.GrainIndex][9] * (dryingInfo.M0W[i] / 100) * (dryingInfo.M0W[i] / 100) #' compute grain density
            ad[i] = CompAirDens(Pv0[i], T0K[i])  #' compute air density
            R[i] = (GD[i] * dryingInfo.dx[i] * (1 - dryingInfo.M0W[i] / 100)) / (dryingInfo.va * t * 3600 * ad[i])  #'compute R: dry matter to dry air ratio, kg/kg
            Cg1[i] = CompGrainSpHeat(Grain.ArrGrain, Grain.GrainIndex, dryingInfo.M0W[i])       #' compute specific heat of grain
            Cg[i] = R[i] * Cg1[i]  #' compute specific heat of grain, converted to Kj/(kg ar K)

            # 'detremine if the current step if drying or rewetting
            # 'a third possibility is also considered, this is when the grain MC of the layer is in between the drying EMC and the rewetting EMC
            # 'in this case the average of the drying and re-wetting curves is used
            Process = 0
            TimeStepEMC_D = (Mdb_Mwb(CF_EMC_D(T0C[i], RH0[i] / 100, Grain.ArrGrain, Grain.GrainIndex) / 100)) * 100
            TimeStepEMC_R = (Mdb_Mwb(CF_EMC_R(T0C[i], RH0[i] / 100, Grain.ArrGrain, Grain.GrainIndex) / 100)) * 100
            if TimeStepEMC_D == TimeStepEMC_R:
                Process = 0     #'in case that drying and rewetting parameters are the same
            else:
                if dryingInfo.M0W[i] >= TimeStepEMC_D:
                    Process = 0
                elif dryingInfo.M0W[i] >= TimeStepEMC_R:
                    Process = 1
                else:
                    Process = 2

            flag = 1
            n = 0
            while flag == 1:
                if n == 0:
                    Hf1[n] = H0[i] * 0.99
                elif n == 1:
                    Hf1[n] = H0[i] * 1.01
                else:
                    Hf1[n] = Hf1[n - 1] - dRH1[n - 1] * ((Hf1[n - 2] - Hf1[n - 1]) / (dRH1[n - 2] - dRH1[n - 1]))

                Mfd1[n] = M0d[i] - 100 * (Hf1[n] - H0[i]) / R[i]
                TfC1[n] = ((ca + cv * H0[i]) * T0C[i] - (Hf1[n] - H0[i]) * (hv - cw * dryingInfo.G0C[i]) + Cg[i] * dryingInfo.G0C[i]) / (ca + cv * Hf1[n] + Cg[i])
                TfK1[n] = Kelvin(TfC1[n])
                if Process == 0:
                    RHf1A[n] = CF_ERH_D(TfC1[n], Mfd1[n], Grain.ArrGrain, Grain.GrainIndex) * 100
                elif Process == 1:
                    RHf1A[n] = CF_ERH_R(TfC1[n], Mfd1[n], Grain.ArrGrain, Grain.GrainIndex) * 100
                else:
                    RHf1A[n] = ((CF_ERH_D(TfC1[n], Mfd1[n], Grain.ArrGrain, Grain.GrainIndex) * 100) + (CF_ERH_R(TfC1[n], Mfd1[n], Grain.ArrGrain, Grain.GrainIndex) * 100)) / 2
                Psf1[n] = Sat_press(TfK1[n])
                Pvf1[n] = Hf1[n] * Pa / (0.6219 + Hf1[n])
                RHf1B[n] = Pvf1[n] / Psf1[n] * 100
                dRH1[n] = RHf1A[n] - RHf1B[n]
                if math.fabs(dRH1[n]) < 0.001:
                    flag = 0
                    Hf[i] = Hf1[n]
                    Mfd[i] = Mfd1[n]
                    dryingInfo.MfW[i] = Mdb_Mwb(Mfd[i] / 100) * 100
                    TfC[i] = TfC1[n]
                    dryingInfo.GfC[i] = TfC[i]
                    RHf[i] = RHf1B[n]
                    if RHf[i] >= 100:
                        RHf[i] = 99
                else:
                    flag = 1
                n = n + 1

            #'update layer depth at the end of the timestep
            dryingInfo.dxf[i] = CompLayerDepthF(Grain.GrainIndex, dryingInfo.M0W[i], dryingInfo.MfW[i], Grain.ArrGrain, dryingInfo.dx[i])

        #'compute DML of the layer if grain is corn
        if Grain.ArrGrain[Grain.GrainIndex][14] == "1":
            #global GT,GMC,DML_GramsPerKg,DML_TempIncr_C
            damageInfo.GT = dryingInfo.GfC[i]
            damageInfo.GMC = dryingInfo.MfW[i]
            damageInfo.DML_GramsPerKg = CompCO2Prod(damageInfo.GT, damageInfo.GMC, damageInfo.GDamage, damageInfo.DML_Mult_Fungicide, damageInfo.DML_Mult_Genetics) / 14.7
            Layer.LayerDML[i] = Layer.LayerDML[i] + damageInfo.DML_GramsPerKg
            damageInfo.DML_TempIncr_C = CompDML_TempInc(damageInfo.DML_GramsPerKg, Grain.ArrGrain, Grain.GrainIndex, dryingInfo.GMC)
            TfC[i] = TfC[i] + damageInfo.DML_TempIncr_C

    return dryingInfo, damageInfo


def PrintToFile(PrintVar, m, j, PrintTime, PrintHeading):

    #' to print the values on the table
    Valuestrim = "\t"
    for i in range(m+3):
        if i < m:
            VarValue = '%.2f' % PrintVar[i]
        elif i == m:
            nonZero = [k for k in PrintVar if k != 0]
            VarValue = '%.2f' % mean(nonZero)
        elif i == (m + 1):
            nonZero = [k for k in PrintVar if k != 0]
            VarValue = '%.2f' % min(nonZero)
        else:
            VarValue = '%.2f' % max(PrintVar)

        Valuestrim = Valuestrim + VarValue + '\t'

    if j == 0:
    #' to print the headings of the table
        return PrintHeading + '\n' +  PrintTime + Valuestrim + '\n' # FileNumber, PrintHeading
    else:
        return PrintTime + Valuestrim + '\n'



def FixInlet_Strat(Grain, Layer, grainInfo, outputFile, binPara, airPara, fanPara, stopPara):
    Counter = 0
    #'set the initial temperature and MC conditions for each grain layer
    #global GrainMC_WB_C, GrainTemp_C_C, LayerDepth_C,GrainMC_WB_S, GrainTemp_C_S,LayerDepth_S
    for i in range(Layer.NumberOfLayers):
        grainInfo.GrainMC_WB_C[i] = grainInfo.GrainMC_In_WB_C[i]
        grainInfo.GrainMC_WB_S[i] = grainInfo.GrainMC_In_WB_S[i]
        grainInfo.GrainTemp_C_C[i] = grainInfo.GrainTemp_In_C_C[i]
        grainInfo.GrainTemp_C_S[i] = grainInfo.GrainTemp_In_C_S[i]

#    'set the initial grain layer depth
    for i in range(Layer.NumberOfLayers):
        Layer.LayerDepth_C[i] = Layer.LayerDepth[i]
        Layer.LayerDepth_S[i] = Layer.LayerDepth[i]

    f3 = open("output/_mccenter.txt",'w')
    f4 = open("output/_mcside.txt", "w")
    f6 = open("output/_tempcenter.txt", "w")
    f7 = open("output/_tempside.txt", "w")
    f9 = open("output/_dmlcenter.txt", "w")
    f10 = open("output/_dmlside.txt", "w")
    #f3.write("abafasdfagagagfgdg")
    stopPara.StopSim = False
    while not stopPara.StopSim:
    #'compute drying for the center of the bin
        dryingInfo, damageInfo = Drying(Grain, Layer, grainInfo, airPara.AirTemp_C, airPara.AirRH, airPara.AirVel_C, fanPara, center=True)
        #GrainMC_WB_C, GrainTemp_C_C, AirTemp_C, AirRH, LayerDepth_C, ArrGrain, GrainIndex, AirVel_C, FanStatus
    #'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the center of the bin
        for i in range(Layer.NumberOfLayers):
            grainInfo.GrainMC_WB_C[i] = dryingInfo.MfW[i]
            grainInfo.GrainTemp_C_C[i] = dryingInfo.GfC[i]
            Layer.LayerDepth_C[i] = dryingInfo.dxf[i]
            damageInfo.LayerDML_C[i] = damageInfo.LayerDML[i]


    #'write the headings for the output file
        textstrim = ""
        for i in range(Layer.NumberOfLayers+3):
            if i < Layer.NumberOfLayers:
                varprint = "Layer" + str(i + 1)
            elif i == Layer.NumberOfLayers:
                varprint = "Average"
            elif i == Layer.NumberOfLayers+1:
                varprint = "Minimun"
            else:
                varprint = "Maximum"

            textstrim = textstrim + varprint + '\t'

        PrintHeading = "Hours" + '\t' + textstrim

        #'print in a file the moisture content of each layers at the center of the bin
        #FileNumber = f3
        output = PrintToFile(grainInfo.GrainMC_WB_C, Layer.NumberOfLayers, Counter, str(Counter+1), PrintHeading)
        f3.write(output)
        #'print in a file the temperature of each layers at the center of the bin
        #FileNumber = f6
        output = PrintToFile(grainInfo.GrainTemp_C_C, Layer.NumberOfLayers , Counter, str(Counter+1), PrintHeading)
        f6.write(output)
        if Grain.ArrGrain[Grain.GrainIndex][14] == '1':
            #'print in a file the DML of each layers at the center of the bin
            #FileNumber = f9
            output = PrintToFile(damageInfo.LayerDML_C, Layer.NumberOfLayers , Counter, str(Counter+1), PrintHeading)
            f9.write(output)

        #'compute drying for the side of the bin
        dryingInfo, damageInfo = Drying(Grain, Layer, grainInfo, airPara.AirTemp_C, airPara.AirRH, airPara.AirVel_S, fanPara, center=False)

        #GrainMC_WB_S, GrainTemp_C_S, AirTemp_C, AirRH, LayerDepth_S, ArrGrain, GrainIndex, AirVel_S, FanStatus
        #'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the side of the bin

        for i in range(Layer.NumberOfLayers):
            grainInfo.GrainMC_WB_S[i] = dryingInfo.MfW[i]
            grainInfo.GrainTemp_C_S[i] = dryingInfo.GfC[i]
            Layer.LayerDepth_S[i] = dryingInfo.dxf[i]
            damageInfo.LayerDML_S[i] = damageInfo.LayerDML[i]

        #'print in a file the moisture content of each layers at the side of the bin
        #FileNumber = f4
        output = PrintToFile(grainInfo.GrainMC_WB_S, Layer.NumberOfLayers, Counter, str(Counter+1), PrintHeading)
        f4.write(output)
        #'print in a file the temperature of each layers at the side of the bin
        #FileNumber = f7
        output = PrintToFile(grainInfo.GrainTemp_C_S, Layer.NumberOfLayers, Counter, str(Counter+1), PrintHeading)
        f7.write(output)

        if Grain.ArrGrain[Grain.GrainIndex][14] == "1":
            #'print in a file the DML of each layers at the side of the bin
            #FileNumber = f10
            output = PrintToFile(damageInfo.LayerDML_S, Layer.NumberOfLayers, Counter, str(Counter+1), PrintHeading)
            f10.write(output)

       # StopSimulation(MCCriteria, TempCriteria, DMLCriteria, GrainMC_WB_C, GrainMC_WB_S, GrainTemp_C_C, GrainTemp_C_S, LayerDML_C, LayerDML_S, Counter + 1, NumberOfLayers - 1, MCavg_Stop, MCmax_Stop, Tempavg_Stop, Tempmax_Stop, DMLavg_Stop, Hours_Stop)

        Counter = Counter + 1
        if Counter >200:
            stopPara.StopSim = True

    f3.close()
    f4.close()
    f6.close()
    f7.close()
    f9.close()
    f10.close()

    return



def getTime(fileSimTime):
    file = open('Weather.txt', 'r')
    Lines = file.readlines()
    first_line = Lines[0]
    fileSimTime.InitialFileYear = int(first_line.split('\t')[0])
    fileSimTime.InitialFileMonth = int(first_line.split('\t')[1])
    fileSimTime.InitialFileDay = int(first_line.split('\t')[2])
    last_line = Lines[-1]
    fileSimTime.FinalFileYear = int(last_line.split('\t')[0])
    fileSimTime.FinalFileMonth = int(last_line.split('\t')[1])
    fileSimTime.FinalFileDay = int(last_line.split('\t')[2])

    file.close()

    return


def CNA_Strat(Grain, Layer, grainInfo, outputFile, binPara, airPara, fanPara, stopPara):

    fileSimTime = FileSimTime()
    getTime(fileSimTime)

    heaterInfo = HEATER()
    heaterInfo.HeaterStatus = False
    fanInfo = FAN()

    simTime = SimTime()
    WFileInfo = WFILEInfo()
    weatherData = WeatherData()

    WF_first_day = datetime.datetime(fileSimTime.InitialFileYear,fileSimTime.InitialFileMonth,fileSimTime.InitialFileDay)
    WF_last_day = datetime.datetime(fileSimTime.FinalFileYear,fileSimTime.InitialFileMonth,fileSimTime.FinalFileDay)

    sim_fist_day = datetime.datetime(simTime.InitialSimYear,simTime.InitialSimMonth,simTime.InitialSimDay)
    sim_last_day = datetime.datetime(simTime.FinalSimYear,simTime.FinalSimMonth,simTime.FinalSimDay)

    WFileInfo.DaysToStart = (sim_fist_day - WF_first_day).days
    WFileInfo.FirstLineRead = WFileInfo.DaysToStart * 24

    WFileInfo.DaysToFinish = (sim_last_day - WF_first_day).days + 1
    WFileInfo.LastLineRead = WFileInfo.DaysToFinish * 24

    multYear = MultYear()
    if simTime.FinalSimYear != simTime.InitialSimYear:
        multYear.MultYears = True
    multYear.IMultYear = simTime.InitialSimYear
    multYear.FMultYear = simTime.FinalSimYear

    CNAlimit = CNALimit()

    for CurrentYear in range(multYear.IMultYear,multYear.FMultYear+1):

        f3 = open("output/"+str(CurrentYear)+"_mccenter.txt", 'w')
        f4 = open("output/"+str(CurrentYear)+"_mcside.txt", "w")
        f6 = open("output/"+str(CurrentYear)+"_tempcenter.txt", "w")
        f7 = open("output/"+str(CurrentYear)+"_tempside.txt", "w")
        f9 = open("output/"+str(CurrentYear)+"_dmlcenter.txt", "w")
        f10 = open("output/"+str(CurrentYear)+"_dmlside.txt", "w")
        f11 = open("output/"+str(CurrentYear)+"_Fan.txt", "w")

        for i in range(Layer.NumberOfLayers):
            grainInfo.GrainMC_WB_C[i] = grainInfo.GrainMC_In_WB_C[i]
            grainInfo.GrainMC_WB_S[i] = grainInfo.GrainMC_In_WB_S[i]
            grainInfo.GrainTemp_C_C[i] = grainInfo.GrainTemp_In_C_C[i]
            grainInfo.GrainTemp_C_S[i] = grainInfo.GrainTemp_In_C_S[i]

        #    'set the initial grain layer depth
        for i in range(Layer.NumberOfLayers):
            Layer.LayerDepth_C[i] = Layer.LayerDepth[i]
            Layer.LayerDepth_S[i] = Layer.LayerDepth[i]

        Counter = 0

        file = open('Weather.txt', 'r')
        Lines = file.readlines().rstrip()

        for i in range(WFileInfo.FirstLineRead, WFileInfo.LastLineRead):
            Line = Lines[i]
            weatherData.BadDataflag = 1
            weatherData.FileTemp = Line.split('\t')[6]
            weatherData.FileRH = Line.split('\t')[7]
            if weatherData.FileTemp < -29.9 or weatherData.FileTemp > 40 or weatherData.FileRH < 1 or weatherData.FileRH > 40:
                weatherData.BadDataflag = 0

            if weatherData.BadDataflag == 1:
                weatherData.FileEMC_db = CF_EMC_D(weatherData.FileTemp, weatherData.FileRH/100, Grain.ArrGrain, Grain.GrainIndex)
                weatherData.FileEMC_wb = 100 * weatherData.FileEMC_db / (100 + weatherData.FileEMC_db)

                if weatherData.FileTemp in range(CNAlimit.MinTempSelect,CNAlimit.MaxTempSelect):
                    weatherData.Tempflag = 1
                else:
                    weatherData.Tempflag = 0

                if weatherData.FileRH in range(CNAlimit.MinRHSelect,CNAlimit.MaxRHSelect):
                    weatherData.RHflag = 1
                else:
                    weatherData.RHflag = 0

                if weatherData.FileEMC_wb in range(CNAlimit.MinEMCSelect, CNAlimit.MaxEMCSelect):
                    weatherData.EMCflag = 1
                else:
                    weatherData.EMCflag = 0

            weatherData.HourFlag = weatherData.Tempflag * weatherData.RHflag * weatherData.EMCflag * weatherData.BadDataflag


            if weatherData.HourFlag == 1:
                fanInfo.FanStatus = True
                plenum = PLENUM()
                plenum.PlenumTemp_C = weatherData.FileTemp + fanInfo.FanPreWarming_C
                plenum.PsAmbient = Sat_press(Kelvin(weatherData.FileTemp))
                plenum.PvPlenum = CompVaporPress(plenum.PsAmbient, weatherData.FileRH)
                plenum.PvPlenum = plenum.PvAmbient
                plenum.PsPlenum = Sat_press(Kelvin(plenum.PlenumTemp_C))
                plenum.PlenumRH = plenum.PvPlenum / plenum.PsPlenum * 100
                plenum.PlenumEMC_db = CF_EMC_D(plenum.PlenumTemp_C, plenum.PlenumRH/100, Grain.ArrGrain, Grain.GrainIndex)
                plenum.PlenumEMC_wb = 100 * plenum.PlenumEMC_db / (100 + plenum.PlenumEMC_db)
            else:
                fanInfo.FanStatus = False

            fanInfo.FanPrint = "Fan Off"
            heaterInfo.HeaterPrint = "Heater Off"

            if fanInfo.FanStatus:
                fanInfo.FanRunHours += 1
                fanInfo.FanPrint = "Fan On"

            fanInfo.Per_FanRun = fanInfo.FanRunHours / (Counter + 1)
            fanInfo.FanKWH = fanInfo.FanPower_KW * fanInfo.FanRunHours

            if heaterInfo.HeaterStatus:
                heaterInfo.HeaterRunHours += 1
                heaterInfo.HeaterPrint = "Heater On"

            heaterInfo.Per_HeaterRun = heaterInfo.HeaterRunHours/(Counter+1)
            heaterInfo.HeaterKWH = heaterInfo.BurnerEnergy_KW * heaterInfo.HeaterRunHours

        #     If
        #     S_OutputSummary = False
        #     Then
        #     'print the heading for the fan output file
        #     PrintHeading_Fan = "Date  /   Time" & vbTab & vbTab & "A Temp" & vbTab & "A RH" & vbTab & "A EMC" & vbTab & "Pl Temp" & vbTab & "Pl RH" & vbTab & "Pl EMC" & vbTab & "Fan St" & vbTab & "Heater St" & vbTab & "F Run T" & vbTab & "F Run %" & vbTab & "H Run T" & vbTab & "H Run %" & vbTab & "F KWH" & vbTab & "H KWH"
        #     'set values into the string to be printed in the fan output file for the current hour
        #     PrintVar_Fan = Format(FileTemp, "#0.0") & vbTab & Format(FileRH, "#0") & vbTab & Format(FileEMC_wb,
        #                                                                                             "#0.0") & vbTab & Format(
        #         PlenumTemp_C, "#0.0") & vbTab & Format(PlenumRH, "#0") & vbTab & Format(PlenumEMC_wb,
        #                                                                                 "#0.0") & vbTab & FanPrint & vbTab & HeaterPrint & vbTab & Format(
        #         FanRunHours, "0") & vbTab & Format(Per_FanRun, "0.0") & vbTab & Format(HeaterRunHours,
        #                                                                                "0") & vbTab & Format(
        #         Per_HeaterRun, "0.0") & vbTab & Format(FanKWH, "0") & vbTab & Format(HeaterKWH, "0")
        #     FileNumber = 11
        #     Call
        #     PrintToFanFile(PrintVar_Fan, Counter, FileNumber, PrintTime, PrintHeading_Fan)
        # End
        # If
            if weatherData.HourFlag == 1:
                dryingInfo, damageInfo = Drying(Grain, Layer, grainInfo, plenum.PlenumTemp_C, plenum.PlenumRH, airPara.AirVel_S, fanPara, center=True)

                for i in range(Layer.NumberOfLayers):
                    grainInfo.GrainMC_WB_C[i] = dryingInfo.MfW[i]
                    grainInfo.GrainTemp_C_C[i] = dryingInfo.GfC[i]
                    Layer.LayerDepth_C[i] = dryingInfo.dxf[i]
                    damageInfo.LayerDML_C[i] = damageInfo.LayerDML[i]


            #'write the headings for the output file
                textstrim = ""
                for i in range(Layer.NumberOfLayers+3):
                    if i < Layer.NumberOfLayers:
                        varprint = "Layer" + str(i + 1)
                    elif i == Layer.NumberOfLayers:
                        varprint = "Average"
                    elif i == Layer.NumberOfLayers+1:
                        varprint = "Minimun"
                    else:
                        varprint = "Maximum"

                    textstrim = textstrim + varprint + '\t'

                PrintHeading = "Hours" + '\t' + textstrim

                #'print in a file the moisture content of each layers at the center of the bin
                #FileNumber = f3
                output = PrintToFile(grainInfo.GrainMC_WB_C, Layer.NumberOfLayers, Counter, str(Counter+1), PrintHeading)
                f3.write(output)
                #'print in a file the temperature of each layers at the center of the bin
                #FileNumber = f6
                output = PrintToFile(grainInfo.GrainTemp_C_C, Layer.NumberOfLayers , Counter, str(Counter+1), PrintHeading)
                f6.write(output)
                if Grain.ArrGrain[Grain.GrainIndex][14] == '1':
                    #'print in a file the DML of each layers at the center of the bin
                    #FileNumber = f9
                    output = PrintToFile(damageInfo.LayerDML_C, Layer.NumberOfLayers , Counter, str(Counter+1), PrintHeading)
                    f9.write(output)

                #'compute drying for the side of the bin
                dryingInfo, damageInfo = Drying(Grain, Layer, grainInfo, airPara.AirTemp_C, airPara.AirRH, airPara.AirVel_S, fanPara, center=False)

                #GrainMC_WB_S, GrainTemp_C_S, AirTemp_C, AirRH, LayerDepth_S, ArrGrain, GrainIndex, AirVel_S, FanStatus
                #'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the side of the bin

                for i in range(Layer.NumberOfLayers):
                    grainInfo.GrainMC_WB_S[i] = dryingInfo.MfW[i]
                    grainInfo.GrainTemp_C_S[i] = dryingInfo.GfC[i]
                    Layer.LayerDepth_S[i] = dryingInfo.dxf[i]
                    damageInfo.LayerDML_S[i] = damageInfo.LayerDML[i]

                #'print in a file the moisture content of each layers at the side of the bin
                #FileNumber = f4
                output = PrintToFile(grainInfo.GrainMC_WB_S, Layer.NumberOfLayers, Counter, str(Counter+1), PrintHeading)
                f4.write(output)
                #'print in a file the temperature of each layers at the side of the bin
                #FileNumber = f7
                output = PrintToFile(grainInfo.GrainTemp_C_S, Layer.NumberOfLayers, Counter, str(Counter+1), PrintHeading)
                f7.write(output)

                if Grain.ArrGrain[Grain.GrainIndex][14] == "1":
                    #'print in a file the DML of each layers at the side of the bin
                    #FileNumber = f10
                    output = PrintToFile(damageInfo.LayerDML_S, Layer.NumberOfLayers, Counter, str(Counter+1), PrintHeading)
                    f10.write(output)


        f3.close()
        f4.close()
        f6.close()
        f7.close()
        f9.close()
        f10.close()
        f11.close()

        file.close()

    return




# def StopSimulation(MCCriteria, TempCriteria, DMLCriteria, GrainMC_WB_C, GrainMC_WB_S, GrainTemp_C_C, GrainTemp_C_S, LayerDML_C, LayerDML_S, CurrentHours, MaxLayer, MCavg_Stop, MCmax_Stop, Tempavg_Stop, Tempmax_Stop, DMLavg_Stop, Hours_Stop):
# #'this sub check for moisture content, temperature, time and DML criteria to decide if the
# #'simulation was completed or not
#     if MCCriteria == True:
#         for i in range(MaxLayer+1):
#             ArrCriteriaC[i] = GrainMC_WB_C[i]
#             ArrCriteriaS[i] = GrainMC_WB_S[i]
#
#     if TempCriteria == True:
#         for i in range(MaxLayer + 1):
#             ArrCriteriaC[i] = GrainTemp_C_C[i]
#             ArrCriteriaS[i] = GrainTemp_C_S[i]
#
#
#     if DMLCriteria == True:
#         for i in range(MaxLayer + 1):
#             ArrCriteriaC[i] = LayerDML_C[i]
#             ArrCriteriaS[i] = LayerDML_S[i]
#
#
#     for i in range(((MaxLayer+1)*2)+1):
#         if i <= MaxLayer:
#             ArrCriteria[i] = ArrCriteriaC[i]
#         else:
#             ArrCriteria[i] = ArrCriteriaS[i - (MaxLayer + 1)]
#
#
#     if MCCriteria == True:
#         avgvalue = mean(ArrCriteria)
#         maxvalue = max(ArrCriteria)
#         if avgvalue <= MCavg_Stop and maxvalue <= MCmax_Stop:
#             StopSim = True
#         else:
#             StopSim = False
#
#     elif TempCriteria == True:
#         avgvalue = mean(ArrCriteria)
#         maxvalue = max(ArrCriteria)
#         if avgvalue <= Tempavg_Stop and maxvalue <= Tempmax_Stop:
#             StopSim = True
#         else:
#             StopSim = False
#     elif DMLCriteria == True:
#         avgvalue = mean(ArrCriteria)
#         if avgvalue >= DMLavg_Stop:
#             StopSim = True
#         else:
#             StopSim = False
#     else:
#         if CurrentHours >= Hours_Stop:
#             StopSim = True
#         else:
#             StopSim = False


#Private
def Com_Run_Click(Grain, Layer):

    print("Running Simulation")
    #'to update all the settings before running the simulation
    grainInfo, outputFile, binPara, airPara, fanPara, stopPara, stringInfo = Com_UpdateSet(Grain, Layer)

    # global TotalDaysToFile, TotalDaysStartSim, DaysToStart,FirstLineRead, TotalDaysFinishSim,DaysToFinish,LastLineRead,CurrentYear
    # #'set all the weather file variables to 0
    # TotalDaysToFile = 0
    # TotalDaysStartSim = 0
    # DaysToStart = 0
    # FirstLineRead = 0
    # TotalDaysFinishSim = 0
    # DaysToFinish = 0
    # LastLineRead = 0
    # CurrentYear = 0

    #'select drying strategy
    # switch(Co_SelectStrat.ListIndex)
    #     Case Is = 0
    #         Call FixInlet_Strat(AirTemp_C, AirRH)
    #     Case Is = 1
    #         Call CNA_Strat
    #     Case Is = 2
    #         Call ConstHeat_Strat
    #     Case Is = 3
    #         Call VarHeat_Strat
    #     Case Is = 4
    #         Call SAVH_Strat
    # End Select

    FixInlet_Strat(Grain, Layer, grainInfo, outputFile, binPara, airPara, fanPara, stopPara)


    #CNA_Strat(Grain, Layer, grainInfo, outputFile, binPara, airPara, fanPara, stopPara)



    print("Simulation Completed")

    return


def Com_UpdateSet(Grain, Layer):
    #'set the basename for the output files
    #global BaseName, AvgInGrainTemp, AvgInGrainMC, AvgGrainDensity,BinDiameter, BinHeight
    grainInfo = GrainInfo(Layer.NumberOfLayers)
    newFile = OutputFile(Layer.NumberOfLayers)
    newFile.BaseName = 'test'
    #' update change in initial grain MC
    #AvgInGrainMC = Tb_AvgGrainMC
    grainInfo.AvgInGrainMC = 20
    #' update change in initial grain Temp
    #AvgInGrainTemp = Tb_AvgGrainTemp
    grainInfo.AvgInGrainTemp = 20
    #' update change in grain density
    grainInfo.AvgGrainDensity = (ComputeGrainDensity(Grain.GrainIndex, grainInfo.AvgInGrainMC, Grain.ArrGrain)) / 1000

    #'update initial MC and temperature conditions fo all the layers
    for i in range(Layer.NumberOfLayers):
        grainInfo.GrainMC_In_WB_C[i] = grainInfo.AvgInGrainMC
        grainInfo.GrainMC_In_WB_S[i] = grainInfo.AvgInGrainMC
        grainInfo.GrainTemp_In_C_C[i] = grainInfo.AvgInGrainTemp
        grainInfo.GrainTemp_In_C_S[i] = grainInfo.AvgInGrainTemp

    #' update change in bin diameter
    binPara = BIN()
    binPara.BinDiameter = 12
#     if Tb_BinDiam = Empty:
#     BinDiameter = 1  #
#
#
# Else
# If
# Tb_BinDiam <= 0
# Then
# BinDiameter = 0.001
# Else
# BinDiameter = Tb_BinDiam
# End
# If
# End
# If
# Lb_BinDiam_ft.Caption = (Int(BinDiameter / 0.3048 * 100)) / 100 + " ft"

# ' update change in bin height
    binPara.BinHeight = 7
# If
# Tb_BinHeight = Empty
# Then
# BinHeight = 1  #
# Else
# If
# Tb_BinHeight <= 0
# Then
# BinHeight = 0.001
# Else
# BinHeight = Tb_BinHeight
# End
# If
# End
# If
# Lb_BinHeight_ft.Caption = (Int(BinHeight / 0.3048 * 100)) / 100 + " ft"
    #global BinArea, BinCapacity_t, BinCapacity_bu, AirflowRate, AirflowRate_cfm, TotalAirflow, AirVel_C, AirVel_S
# ' compute the area of the bin
    binPara.BinArea = pi * (binPara.BinDiameter * binPara.BinDiameter) / 4
# ' compute bin storage capacity in tonnes
    binPara.BinCapacity_t = binPara.BinArea * binPara.BinHeight * grainInfo.AvgGrainDensity
# 'disply updated change in bin capacity (tonnes)
# Lb_BinCap_t.Caption = (Int(BinCapacity_t * 100)) / 100 + " tonnes"
# 'update change in bin capacity (bushels)
    binPara.BinCapacity_bu = binPara.BinCapacity_t * 40
# 'disply updated change in bin capacity (bushels)
# Lb_BinCap_bu.Caption = (Int(BinCapacity_bu * 100)) / 100 + " bushels"


    #' update change in airflow rate
    #AirflowRate = Tb_Airflow
    airPara = AIR()
    airPara.AirflowRate = 1
    #If AirflowRate = Empty Then AirflowRate = 1
    airPara.AirflowRate_cfm = (int(airPara.AirflowRate / 1.11 * 100)) / 100
    #Lb_AirflowCFM.Caption = AirflowRate_cfm + " cfm/bu"
    #'update change in total airflow rate
    airPara.TotalAirflow = airPara.AirflowRate * binPara.BinCapacity_t

    #'set the air speed at the center and side of the bin
    airPara.AirVel_C = airPara.TotalAirflow / binPara.BinArea / 60 * (1 - airPara.AirflowNonUnif)
    airPara.AirVel_S = airPara.TotalAirflow / binPara.BinArea / 60 * (1 + airPara.AirflowNonUnif)

    #global SingleLayerDepth, SingleLayerDepth1, NumberOfLayers
    #SingleLayerDepth1 = 0.5
    #' update the number of layer in the bin (0.5m each layer), and adjust the layer depth if #of layers is > 32
    #SingleLayerDepth = SingleLayerDepth1 - 0.1
    # NumberOfLayers = 32
    # while not NumberOfLayers <= 31:
    #     SingleLayerDepth = SingleLayerDepth + 0.1
    #     if BinHeight - int(BinHeight) < 0.25:
    #        NumberOfLayers = int(BinHeight / SingleLayerDepth)
    #     else:
    #        NumberOfLayers = int(BinHeight / SingleLayerDepth) + 1


    #Lb_LayerDepth.Caption = "L Depth: " + Format(SingleLayerDepth, "0.00") + " m"
    #Tb_NumberLayers = NumberOfLayers
    #' set grain layer depth
    SetGrainLayerDepth(Layer)
    #' update initial simulation dates
    # InitialSimYear = Co_ISYear.ListIndex + 1961
    # InitialSimMonth = Co_ISMonth.ListIndex + 1
    # InitialSimDay = Co_ISDay.ListIndex + 1
    # ' update final simulation dates
    # FinalSimYear = Co_FSYear.ListIndex + 1961
    # FinalSimMonth = Co_FSMonth.ListIndex + 1
    # FinalSimDay = Co_FSDay.ListIndex + 1
    # 'update multiple year dates

    # IMultYear = Co_IMultYear.ListIndex + 1961
    # If Opt_MultYears = True Then
    #     FMultYear = Co_FMultYear.ListIndex + 1961
    # Else
    #     FMultYear = Co_ISYear
    # End If
    #global AirTemp_C,AirRH,Hours_Stop
    #'update emc calculation based on the fixed inlet air conditios
    airPara.AirTemp_C = 20
    airPara.AirRH = 70
    #Lb_Strat6.Caption = Format(Mdb_Mwb(CF_EMC_D(AirTemp_C, AirRH / 100, ArrGrain, GrainIndex) / 100) * 100, "#0.##")
    #Lb_Strat7.Caption = Format(Mdb_Mwb(CF_EMC_R(AirTemp_C, AirRH / 100, ArrGrain, GrainIndex) / 100) * 100, "#0.##")

    #'end of simulation criteria
    # If Opt_Date = True Then
    #     #'update number of hours for the end of the simulation
    #     Hours_Stop = Tb_FinalAvg
    #     StStopSimulation = "Stop Sim.: Date (" + FinalSimMonth + "/" + FinalSimDay + ")"
    # End If
    # If Opt_MC = True Then
    #     'update MC limit for ending criteria
    #     MCavg_Stop = Tb_FinalAvg
    #     MCmax_Stop = Tb_FinalMax
    #     StStopSimulation = "Stop Sim.: MC (" + MCavg_Stop + "-" + MCmax_Stop + ")"
    #
    # End If
    # If Opt_Temp = True Then
    #     'update temp limit for ending criteria
    #     Tempavg_Stop = Tb_FinalAvg
    #     Tempmax_Stop = Tb_FinalMax
    #     StStopSimulation = "Stop Sim.: Temp.(" + Tempavg_Stop + "-" + Tempmax_Stop + ")"
    # End If
    # If Opt_DML = True Then
    #     'update DML limit for ending criteria
    #     DMLavg_Stop = Tb_FinalAvg
    #     StStopSimulation = "Stop Sim.: DML(" + dmlpavg_Stop + ")"
    # End If
    stopPara = StopCriteria(Layer.NumberOfLayers)
    stopPara.Hours_Stop = 15

    #     'update selecting criteria values
    #     If
    #     Co_SelectStrat.ListIndex >= 1
    #     Then
    #     If
    #     Co_SelectStrat.ListIndex = 1
    #     Then
    #     MinTempSelect = Tb_CrMinTemp
    #     MaxTempSelect = Tb_CrMaxTemp
    #     MinRHSelect = Tb_CrMinRH
    #     MaxRHSelect = Tb_CrMaxRH
    #     MinEMCSelect = Tb_CrMinEMC
    #     MaxEMCSelect = Tb_CrMaxEMC
    #
    #
    # ElseIf
    # Co_SelectStrat.ListIndex = 2
    # Then
    # MinRHSelect = Tb_CrMinRH
    # MaxRHSelect = Tb_CrMaxRH
    # MinEMCSelect = Tb_CrMinEMC
    # MaxEMCSelect = Tb_CrMaxEMC
    # ElseIf
    # Co_SelectStrat.ListIndex = 3
    # Then
    # MinEMCSelect = Tb_CrMinEMC
    # MaxEMCSelect = Tb_CrMaxEMC
    # Else
    # 'update the desired final moiture content for the savh strategy
    # SAVH_FinalMC = Tb_SAVHmc
    # End
    # If
    # End
    # If
    #global AirfResistance, AirfResistance_Wa, FanPower_HP, FanPower_KW, Tinc_C, StBin, StAirflow, StGrain, StStartDate, StGrainEMC,StRunInfo
    #'update the static pressure of the system, Pa
    airPara.AirfResistance = AirFlowResistance(Grain.GrainIndex, Grain.ArrGrain, airPara.TotalAirflow, binPara.BinArea, binPara.PackingFactor) * binPara.BinHeight
    #'update the static pressure of the system, inches of water
    airPara.AirfResistance_Wa = airPara.AirfResistance * PaToInWa
    fanPara = FAN()
    #'update values for the fan power
    fanPara.FanPower_HP = int(CompFanPower(airPara.TotalAirflow, airPara.AirfResistance_Wa, fanPara.FanEfficiency) * 100) / 100
    #' update the power requirement for the fan operation, KW
    fanPara.FanPower_KW = int(fanPara.FanPower_HP * HPtoKW * 100) / 100
    #'update the temperature increse value demanded to the burner
    #Tinc_C = Tb_TempIncr

    stringInfo = StringInfo()
    simTime = SimTime()
    #'set the bin dimensions for the run information string
    stringInfo.StBin = "Height: " + str(binPara.BinHeight) + ", Diameter: " + str(binPara.BinDiameter)
    #'set the airflow for the run information string
    stringInfo.StAirflow = str(airPara.AirflowRate)
    #'set the initial grain temperature and MC for the run information string
    stringInfo.StGrain = "Initial Temp.: " + str(grainInfo.AvgInGrainTemp) + ", Initial MC: " + str(grainInfo.AvgInGrainMC)
    #'set the initial month and day of the simulation for the run information string
    stringInfo.StStartDate = str(simTime.InitialSimMonth) + "/" + str(simTime.InitialSimDay)
    #'set the selected grain EMC for the run information string
    stringInfo.StGrainEMC = Grain.ArrGrain[Grain.GrainIndex][0]
    #'select drying strategy settings for the run information string
    # Select Case Co_SelectStrat.ListIndex
    #     Case Is = 0
    #         StStrategy = "Fix Inlet (" + AirTemp_C + "ºC" + ", " + AirRH + "%)"
    #     Case Is = 1
    #         StStrategy = "CNA (EMC: " + MinEMCSelect + "-" + MaxEMCSelect + "%, Temp: " + MinTempSelect + "-" + MaxTempSelect + "ºC, RH: " + MinRHSelect + "-" + MaxRHSelect + "%)"
    #     Case Is = 2
    #         StStrategy = "Const. Heat(EMC: " + MinEMCSelect + "-" + MaxEMCSelect + "%, RH: " + MinRHSelect + "-" + MaxRHSelect + "%, Temp. Inc.: " + Tinc_C + "ºC)"
    #     Case Is = 3
    #         StStrategy = "Var. Heat(EMC: " + MinEMCSelect + "-" + MaxEMCSelect + "%)"
    #     Case Is = 4
    #         StStrategy = "SAVH (Target MC: " + SAVH_FinalMC + "%)"
    # End Select

    stringInfo.StRunInfo = "This run was made with the following settings: Weather File: " + stringInfo.StWeatherFile + "; Strategy: " + stringInfo.StStrategy + "; Grain EMC: " + stringInfo.StGrainEMC + "; Bin Dimensions: " + stringInfo.StBin + "; Grain: " + stringInfo.StGrain + "; Airflow: " + stringInfo.StAirflow + " m3/min/t; Starting Date: " + stringInfo.StStartDate + "; " + stringInfo.StStopSimulation

    return grainInfo, newFile, binPara, airPara, fanPara, stopPara, stringInfo


# def StopSimulation(MCCriteria, TempCriteria, DMLCriteria, GrainMC_WB_C, GrainMC_WB_S, GrainTemp_C_C, GrainTemp_C_S, LayerDML_C, LayerDML_S, CurrentHours, MaxLayer, MCavg_Stop, MCmax_Stop, Tempavg_Stop, Tempmax_Stop, DMLavg_Stop, Hours_Stop):
# #'this sub check for moisture content, temperature, time and DML criteria to decide if the
# #'simulation was completed or not
#     if MCCriteria == True:
#         i = 0
#         For i = 0 To MaxLayer
#             ArrCriteriaC(i) = GrainMC_WB_C(i)
#             ArrCriteriaS(i) = GrainMC_WB_S(i)
#         Next i
#     End If
#     If TempCriteria = True Then
#         i = 0
#         For i = 0 To MaxLayer
#             ArrCriteriaC(i) = GrainTemp_C_C(i)
#             ArrCriteriaS(i) = GrainTemp_C_S(i)
#         Next i
#     End If
#     If DMLCriteria = True Then
#         i = 0
#         For i = 0 To MaxLayer
#             ArrCriteriaC(i) = LayerDML_C(i)
#             ArrCriteriaS(i) = LayerDML_S(i)
#         Next i
#     End If
#
#     i = 0
#     For i = 0 To ((MaxLayer * 2) + 1)
#         If i <= MaxLayer Then
#             ArrCriteria(i) = ArrCriteriaC(i)
#         Else
#             ArrCriteria(i) = ArrCriteriaS(i - (MaxLayer + 1))
#         End If
#     Next i
#
#     If MCCriteria = True Then
#         avgvalue = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2) + 1)
#         maxvalue = ArrayMax(ArrCriteria, 0, (MaxLayer * 2) + 1)
#         If avgvalue <= MCavg_Stop And maxvalue <= MCmax_Stop Then
#             StopSim = True
#         Else
#             StopSim = False
#         End If
#     ElseIf TempCriteria = True Then
#         avgvalue = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2) + 1)
#         maxvalue = ArrayMax(ArrCriteria, 0, (MaxLayer * 2) + 1)
#         If avgvalue <= Tempavg_Stop And maxvalue <= Tempmax_Stop Then
#             StopSim = True
#         Else
#             StopSim = False
#         End If
#     ElseIf DMLCriteria = True Then
#         avgvalue = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2) + 1)
#         If avgvalue >= DMLavg_Stop Then
#             StopSim = True
#         Else
#             StopSim = False
#         End If
#     Else
#         If CurrentHours >= Hours_Stop Then
#             StopSim = True
#         Else
#             StopSim = False


# def Form_Load():
    #CurrentDir = CurDir
    #' load the in-bin drying strategy list
    # Co_SelectStrat.AddItem
    # "Fix Inlet Conditions"
    # Co_SelectStrat.AddItem
    # "Natural Air"
    # Co_SelectStrat.AddItem
    # "Constant Heat"
    # Co_SelectStrat.AddItem
    # "Variable Heat"
    # Co_SelectStrat.AddItem
    # "Self Adapting Variable Heat"

    #'location of the file with the list of grains available and it parameters
    FileName = "graininfo.dry"
    #ReadGrainList1(FileName)
    #'set initial option for the grain list to "0"
    #Co_SelectGrain.ListIndex = 0

    # 'set defoult values for fix inlet air conditions
    # Tb_Strat1 = 20
    # Tb_Strat2 = 70
    #
    # 'set initial option for the strategy list to "0"
    # Co_SelectStrat.ListIndex = 0
    # 'set the defoult values for initial simulation date
    # InitialSimYear = 1961
    # InitialSimMonth = 10
    # InitialSimDay = 1
    # 'set the defoult values for final simulation date
    # FinalSimYear = 1961
    # FinalSimMonth = 12
    # FinalSimDay = 31
    # 'set the defoult values for the multiple years simulation
    # IMultYear = 1961
    # FMultYear = 2000
    #
    # 'set defoult values for the temp, RH and EMC windows selecting criteria for the CNA strategy
    # MinTempSelect = -50
    # MaxTempSelect = 50
    # MinRHSelect = 0
    # MaxRHSelect = 100
    # MinEMCSelect = 0
    # MaxEMCSelect = 50
    #
    # 'set average initial values for temp and MC
    # AvgInGrainMC = Tb_AvgGrainMC.Text
    # AvgInGrainTemp = Tb_AvgGrainTemp.Text

    #'set initial MC and temperature conditions fo all the layers
    # for i in range(32):
    #     GrainMC_In_WB_C[i] = AvgInGrainMC
    #     GrainMC_In_WB_S[i] = AvgInGrainMC
    #     GrainTemp_In_C_C[i] = AvgInGrainTemp
    #     GrainTemp_In_C_S[i] = AvgInGrainTemp


if __name__ == "__main__":
    grainSet  = ReadGrainLIst("graininfo.txt")
    numberofLayer = 14
    SingleLayerDepth = 0.5
    layerPara = LAYER(numberofLayer, SingleLayerDepth)


    Com_Run_Click(grainSet, layerPara)
