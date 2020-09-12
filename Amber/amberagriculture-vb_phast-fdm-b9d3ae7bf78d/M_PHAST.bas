Attribute VB_Name = "M_PHAST"
'declaration of constants
Public Const ca As Single = 1.005          ' specific heat of dry air; Kj/(kgºK)
Public Const cv As Single = 1.85        ' specific heat of water vapour; Kj/(kgºK)
Public Const cw As Single = 4.186       ' specific heat of water; Kj/(kgºK)
Public Const hv As Single = 2500.8         ' latent heat of vapourization of free water; Kj/(kgºK)
Public Const Pa As Single = 101300        ' atmospheric pressure; ??

'properties of moist air (ASAE 1997)
Public Const R1 As Single = 22105649.25
Public Const a1 As Single = -27405.526
Public Const b1 As Single = 97.5413
Public Const c1 As Single = -0.146244
Public Const d1 As Single = 0.00012558
Public Const e1 As Single = -0.000000048502
Public Const f1 As Single = 4.34903
Public Const g1 As Single = 0.0039381


'declare constants
Public Const pi As Double = 3.14159265358979 ' this is the pi number to compute are of the bin
Public Const PaToInWa As Double = 0.0040146 ' this is the conversion of 1 pa to inches of water
Public Const cfmTom3min As Double = 0.0283168 ' this is the conversion from 1 cfm to m3/min of airflow
Public Const HPtoKW As Single = 0.746 ' this is the conversion from 1 HP to KW of power
Public Const KjToKWH As Single = 0.000278 ' this is the conversion of 1 kJ to KWH
Public Const KWHtoBTU As Single = 3412.141633 ' this is the conversion of 1 KWH to Btu (IT)

'declare variables
Public ArrGrain(30, 14) As Variant
Public LineIndex As Integer
Public GrainIndex As Integer
Public FileName As String
Public TablePath As String
' Form main
Public AirflowRate As Single ' this is the average airflow rate, m3/min/t
Public TotalAirflow As Single 'this is the total airflow rate in the bin, m3/min
Public AirflowRate_cfm As Single ' this is the average airflow rate, cfm/bu
Public AirflowNonUnif As Single 'this is the non-uniformity factor to determine the center and side airflow rates, dec
Public AirflowCenter As Single 'this is the airflow rate at the center of the bin, m3/min/t
Public AirflowSide As Single 'this is the airflow rate at the side of the bin, m3/min/t
Public AirfResistance As Single ' this is the resistance of the airflow, Pa
Public AirfResistance_Wa As Single ' this is the resistance of the airflow, inches of water
Public FanPreWarming_C As Single ' this is the estimated fan prewarming, ºC
Public PackingFactor As Single ' packing factor for the Shedd's curves

Public AirVel_C As Single   'this is the air velocity at the center of the bin, m/s
Public AirVel_S As Single   'this is the air velocity at the side of the bin, m/s

Public FanPower_HP As Single ' this is the estimated fan power, HP
Public FanPower_KW As Single ' this is the estimated fan power, KW
Public HeatingEnergy_KW As Single ' this is the energy required to heat the air, KWH
Public HeatingEnergy_BTU As Single ' this is the energy required to heat the air, BTU
Public BurnerEnergy_KW As Single 'this is the estimated burner power, KW
Public BurnerEnergy_BTU As Single 'this is the estimated burner power, BTU
Public FanEfficiency As Single ' this is the fan efficiency, %
Public BurnerEfficiency As Single ' this is the burner efficiency, %
Public Tinc_C As Single ' this is the temperature increase demanded to the burner, ºC
Public HeaterRHflag As Integer 'this flag indicates if the heater is on or off due to the ambient RH (0+off, 1=on)
Public HeaterEMCflag As Integer 'this flag indicates if the heater is on or off due to the ambient EMC (0+off, 1=on)
Public HeaterFlag As Integer 'this flag indicates if the heater is on or off due to the ambient coditions (0+off, 1=on)

Public BinDiameter As Single ' this is the diameter of the bin, m
Public BinHeight As Single 'this is the height of the bin, m
Public BinArea As Single ' this is the area of the bin, m2
Public BinCapacity_t As Single
Public BinCapacity_bu As Single


Public SingleLayerDepth1 As Single ' initial estimate of layer depth, m
Public SingleLayerDepth As Single ' layer depth, m
Public NumberOfLayers As Integer ' number of layers in the bin
Public LayerDepth(31) As Single ' detpth of each layer at begining of simulation (default value is 0.5m when simulatin starts), meter
Public LayerDepth_C(31) As Single 'depth of each grain layer at the center of the bin anytime after simulation started, m
Public LayerDepth_S(31) As Single 'depth of each grain layer at the side of the bin anytime after simulation started, m

Public AvgInGrainMC As Single
Public AvgInGrainTemp As Single
Public GrainMC_In_WB_C(31) As Single ' MC of the grain in each layer at the center of the bin, %, wb
Public GrainMC_In_WB_S(31) As Single ' MC of the grain in each layer at the side of the bin, %, wb
Public GrainTemp_In_C_C(31) As Single ' Temp of the grain in each layer at the center of the bin, ºC
Public GrainTemp_In_C_S(31) As Single ' Temp of the grain in each layer at the side of the bin, ºC
Public GrainMC_WB_C(31) As Single 'this is the MC of the grain in each layer at the center of the bin at any time, %, wb
Public GrainMC_WB_S(31) As Single 'this is the MC of the grain in each layer at the side of the bin at any time, %, wb
Public GrainTemp_C_C(31) As Single 'this is the temperature of the grain in each layer at the center of the bin at any time, ºC
Public GrainTemp_C_S(31) As Single 'this is the temperature of the grain in each layer at the side of the bin at any time, ºC

Public AvgGrainDensity As Single ' average initial density of the grain, kg/m3 (used to compute airflow resistance)
Public GrainDensity_C(31) As Single ' Desity of the grain in the layer at the center of the bin, kg/m3
Public GrainDensity_S(31) As Single ' Desity of the grain in the layer at the side of the bin, kg/m3

Public Ps As Single ' this is the saturated vapor pressure of the air
Public Pv As Single 'this is the vapor pressure of the air
Public TC As Single 'this is the temperature of the air in ºC
Public TK As Single 'this is the temperature of the air in ºK
Public RH As Single ' this is the RH of the air, %
Public AbsHum As Single ' this is the absolute humidity (H) of the air, g of water/kg of dry air

Public DML_Mult_Damage As Single 'this is the DML multiplier for damage kernel
Public DML_Mult_Fungicide As Single ' this is the DML multiplier for fungicide application
Public DML_Mult_Genetics As Single 'this is the DML multiplier for hybrid
Public DML_Mult_Temp As Single 'this is the DML multiplier for temperature
Public DML_Mult_MC As Single 'this is the DML multiplier for MC
Public GT As Single 'this is the temperature of the grain for DML computation, ºC
Public GMC As Single 'this is the MC of the grain  for DML computation, %, wb
Public GDamage As Single ' this is the percentage of damage grain for DML computation, %
Public tr As Single ' this is the equivalent storage time for the DML computation, hours
Public DML_GramsPerKg As Single 'this is the predicted DML for the timestep, grams of CO2 produced per Kg of Dry Matter
Public DML_BTUPerKg As Single 'this is the amount of BUT produced when 1 gram of CO2 per KG of dry matter is realesed
Public DML_TempIncr_C As Single 'this is the temperature increase of the grain mass due to the heat generated by respiration (DML generation), ºC
Public DMLFlag As String 'this flag indicates if the grain selected is corn, if DMLFlag is 1=true, then the grain is corn and DML will be computed
Public LayerDML(31) As Single 'this is the DML (%) of the layer of corn
Public LayerDML_C(31) As Single 'this is the DML (%) of the layer of corn at the center of the bin
Public LayerDML_S(31) As Single 'this is the DML (%) of the layer of corn at the side of the bin

Public AirTemp_C As Single 'this is the temperature of the drying air, ºC
Public AirRH As Single 'this is the relative humidity of the drying air, %
Public FanStatus As Boolean 'this is the fan status, true when fan is "on" and false when fan is "off"
Public HeaterStatus As Boolean 'this is the burner status, true when burner is "on" and false when burner is "off"
Public SimStatus As Boolean   'this is the status for the simulation, when it is true, simulation will run for next timestep, if it is false it will stop at current time step

Public dx(31) As Single        ' depth of the thin layer; m
Public dxf(31) As Single        ' depth of the thin layer at the end of the timestep
Public MfW(31) As Single     ' final grain moisture content; %, wb
Public GfC(31) As Single     ' final grain temperature, ºC
Dim va As Single             ' velocity of the drying air, m/s
Dim M0W(31) As Single     ' initial grain moisture content; %, wb
Dim G0C(31) As Single     ' initial grain temperature; ºC

Public FileName1 As String ' this is the filename of the file in which the output will be written
Public FileNumber As Integer 'this is the filenumber indicator of a given open file, represents the number in the open file FileName as #1
Public PrintVar(31) As Single 'this is the variable to print in the file
Public TimeStamp As String ' this is the time stamp (hour, i.e.; 1:00) for the outputs files
Public DateStamp As String 'this is the date stamp (month-day-year) for the output files
Public PrintTime As String 'this is the time stamp and date stamp combined for the outputs files
Public PrintHeading As String 'this is the heading string for the output MC, Temp and DML files
Public PrintHeading_Fan As String 'this is the heading string for the output Fan file
Public PrintHeading_End As String 'this is the heading string for the output Summary file
Public PrintVar_Fan As String ' this is the string that carries all the variables for each hour to be printed in the fan output file
Public PrintVar_End As String ' this is the string that carries all the variables for each year to be printed in the Summary output file
Public PrintYear As String 'this is the Year stamp for the Summary outputs files

'variabled associated to the subroutine StopSimulation
Public MCCriteria As Boolean ' criteria to consider simulation completed, when MCCriteria = true, then MC is the selected criteria
Public TempCriteria As Boolean ' criteria to consider simulation completed, when TempCriteria = true, then temperature is the selected criteria
Public DMLCriteria As Boolean ' criteria to consider simulation completed, when DMLCriteria = true, then DML is the selected criteria
Public DateCriteria As Boolean ' criteria to consider simulation completed, when DateCriteria = true, then Date is the selected criteria
Public ArrCriteriaC(31) As Single ' this is the array that carries the information from the center of the bin to decide if simulation was completed or not
Public ArrCriteriaS(31) As Single ' this is the array that carries the information from the side of the bin to decide if simulation was completed or not
Public ArrCriteria(62) As Single ' this is the array that carries the information to decide if simulation was completed or not
Public CurrentHours As Integer ' this is the variable that carries the information about the current number of hours since simulation started
Public StopSim As Boolean 'this is the output of the sub, when true simulation was completed
Public MaxLayer As Single 'this is the number of layers considered in the simulation, used to determine the upper bound of the array
Public MCavg_Stop As Single 'this is the desired average final moisture content value to consider completed the simulation
Public MCmax_Stop As Single ' this is the desired maximum final moisture content value to consider completed the simulation
Public Tempavg_Stop As Single 'this is the desired average final temperature value to consider completed the simulation
Public Tempmax_Stop As Single ' this is the desired maximum final temperature value to consider completed the simulation
Public DMLavg_Stop As Single 'this is the desired average final DML value to consider completed the simulation
Public Hours_Stop As Integer 'this is the desired number of hours to consider completed the simulation

'variables associated with the simulation time period
Public InitialSimYear As Single 'this is the year at which the first simulation will start
Public InitialSimMonth As Single 'this is the month at which the first simulation will start
Public InitialSimDay As Single 'this is the day at which the first simulation will start
Public FinalSimYear As Single 'this is the year at which the first simulation will finish
Public FinalSimMonth As Single 'this is the month at which the first simulation will finish
Public FinalSimDay As Single 'this is the day at which the first simulation will finish

'variables associated with the multiple years simulation period
Public IMultYear As Single 'this is the first year of the multiple years simulation
Public FMultYear As Single 'this is the last year of the multiple years simulation
Public MultYears As Boolean ' this variable indicates if mulatiple years simulation is requested

'varaibles asociated with the weather file
Public TempColumn As Integer 'this is the number of the temperature column in the weather file
Public RHColumn As Integer 'this is the number of the RH column in the weather file
Public InitialFileYear As Integer 'this is the first year in the weather file
Public InitialFileMonth As Integer 'this is the first month in the weather file
Public InitialFileDay As Integer 'this is the first day in the weather file
Public FinalFileYear As Integer 'this is the last year in the weather file
Public FinalFileMonth As Integer 'this is the last month in the weather file
Public FinalFileDay As Integer 'this is the last day in the weather file
Public sItems() As String 'thi  is the string that will contain the information read from each line of the weather file

'variables associated with the location of the line to read in the weather file
Public TotalDaysToFile As Long ' number of days from January 1 of 1960 to the first year month and day of the weather file
Public TotalDaysStartSim As Long 'number of days from January 1 of 1960 to the first year month and day of the analysis
Public DaysToStart As Long 'number of days since the beginin of the weather data to the begining of the analysis
Public FirstLineRead As Long 'number of hours (lines) between the begining of the weather file and the first simulation hour
Public TotalDaysFinishSim As Long 'days from January 1 of 1960 to the last year month and day of the analysis
Public DaysToFinish As Long 'number of days since the beginin of the weather data to the last day of the analysis
Public LastLineRead As Long 'number of hours (lines) between the begining of the weather file and the last simulation hour
Public CurrentYear As Long 'this is the current year of the simulation
Public FileNameW As String 'this is the filename and path of the weather file used for the simulation
Public LineFromFile As String 'this string is the line readed from the weather file

'varaibles associated with the selecting criteria of useful hours for the CNA strategy
Public MinTempSelect As Single 'this is the minimum temperature limit for the temperature window of the CNA strategy, ºC
Public MaxTempSelect As Single 'this is the maximum temperature limit for the temperature window for the CNA strategy, ºC
Public MinRHSelect As Single 'this is the minimum RH limit for the RH window of the CNA strategy, %
Public MaxRHSelect As Single 'this is the maximum RH limit for the RH window of the CAN strategy, %
Public MinEMCSelect As Single 'this is the minimum drying EMC limit for the EMC window of the CNA strategy, %, wb
Public MaxEMCSelect As Single 'this is the maximum drying EMC limit for the EMC window of the CNA strategy, %, wb

'variables associated with weather data
Public RHflag As Integer ' this is the flag that indicates that the RH read from the file fits into the selected window of fan operation, 1 = true, 0= false
Public Tempflag As Integer ' this is the flag that indicates that the temperature read from the file fits into the selected window of fan operation, 1 = true, 0= false
Public EMCflag As Integer ' this is the flag that indicates that the EMC computed from Temp and RH read from the file fits into the selected window of fan operation, 1 = true, 0= false
Public BadDataflag As Integer ' this is the flag that indicates that the temperature and RH data read from the file are good (tempe >-29ºC and <40ºC and RH >1 and <=100%), 1 = true, 0= false
Public HourFlag As Integer 'this is the flag that indicates that the temperature,RH , EMC and BadData flags are true, 1=true, 0=false
Public FileTemp As Single 'this is the ambient temperature data read from the weather file, ºC
Public FileRH As Single 'this is the ambient RH data read from the weather file, %
Public FileEMC_db As Single 'this is the ambient EMC data computed from the ambient T and RH read from the weather file, %, db
Public FileEMC_wb As Single 'this is the ambient EMC data computed from the ambient T and RH read from the weather file, %, wb

'variables associated with the drying air condition at the plenum of the bin
Public PlenumTemp_C As Single ' this is the plenum air drying temperature, ºC
Public PlenumRH As Single ' this is the plenum air RH, %
Public PlenumRHfan As Single ' this is the rh in the plenum after the fan prewarming, %
Public PsAmbient As Single ' this is the Vapor Pressure at Saturation of the ambient air
Public PvAmbient As Single ' this is the Vapor Pressure of the ambient air
Public PsPlenum As Single ' this is the Vapor Pressure at Saturation of the drying air at the plenum
Public PvPlenum As Single ' this is the Vapor Pressure of the drying air at the plenum
Public PlenumEMC_db As Single 'this is the plenum EMC data computed from the plenum T and RH, %, db
Public PlenumEMC_wb As Single 'this is the plenum EMC data computed from the plenum T and RH, %, wb

'variables associated with the fan output file
Public FanRunHours As Single 'this is the total fan run hours of the current year simulation, hours
Public HeaterRunHours As Single 'this is the total heater run hours of the current year simulation, hours
Public Per_FanRun As Single 'this is the % of fan run time of the current year simulation, %
Public Per_HeaterRun As Single 'this is the % of heater run time of the current year simulation, %
Public FanKWH As Single 'this is the cumulative KWH (energy consumption) of the fan during the current simulation year, kwh
Public HeaterKWH As Single 'this is the cumulative KWH (energy consumption) of the heater during the current simulation year, kw
Public FanPrint As String ' this is the string with information about the fan status (F On or F Off) to be printed in the fan oputput file
Public HeaterPrint As String ' this is the string with information about the heater status (H On or H Off) to be printed in the fan oputput file

'variables related to the simulation of a set of years
Public S_Moisture_Avg(50) As Single 'this array constain the final average moisture content for each one of the years of the simulation, %, wb
Public S_Moisture_Min(50) As Single 'this array constain the final minimum moisture content for each one of the years of the simulation, %, wb
Public S_Moisture_Max(50) As Single 'this array constain the final maximum moisture content for each one of the years of the simulation, %, wb
Public S_Temperature_Avg(50) As Single 'this array constain the final average temperature for each one of the years of the simulation,ºC
Public S_Temperature_Min(50) As Single 'this array constain the final minimum temperature for each one of the years of the simulation,ºC
Public S_Temperature_Max(50) As Single 'this array constain the final maximum temperature for each one of the years of the simulation,ºC
Public S_DML_Avg(50) As Single 'this array constain the final average DML for each one of the years of the simulation, %
Public S_DML_Max(50) As Single 'this array constain the final maximum DML for each one of the years of the simulation, %
Public S_DryingHs(50) As Single 'this array constain the final drying hours for each one of the years of the simulation, hours
Public S_FanHs(50) As Single 'this array constain the final fan run hours for each one of the years of the simulation, hours
Public S_PerFanHs(50) As Single 'this array constain the final percentaje of fan run hours for each one of the years of the simulation, %
Public S_HeaterHs(50) As Single 'this array constain the final heater run hours for each one of the years of the simulation, hours
Public S_PerHeaterHs(50) As Single 'this array constain the final percentaje of heater run hours for each one of the years of the simulation, %
Public S_FanKWH(50) As Single 'this array constain the final energy consumption of the fan for each one of the years of the simulation, KWH
Public S_HeaterKWH(50) As Single 'this array constain the final energy consumption of the heater for each one of the years of the simulation, KWH
Public S_OutputSummary As Boolean 'this variable indicates if only Summary results are required, or all results are requires (true=summary, false=all)
Public S_TotDryingCost(50) As Single 'this is the drying cost for each simulation run, $/tonne

Public CurrentDir As String ' this string carries the information about the location of the VB-PHAST program
Public BaseName As String ' this is the basename of the output files

Public SAVH_FinalMC As Single ' this is the desired final moisture content for the savh strategy, %, wb
Public EstDryingTime As Single ' this is the estimated drying time for the savh strategy, % drying time = 15*50/cfm/bu
Public MC_Highest As Single ' this is the highest MC of the first grain layer at any given hour, %, wb
Public MC_Lowest As Single ' this is the lowest MC of the first grain layer at any given hour, %, wb
Public MC_HighLimit As Single ' this is the high MC limit at any given hour, %, wb
Public MC_LowLimit As Single ' this is the low MC limit at any given hour, %, wb

'variables related to the computation of the drying cost
Public GrainPrice As Single 'this is the price of the grain, in $/tonne
Public ElectricityCost As Single 'this is the cost of electricity, in $/kwh
Public PropaneCost As Single 'this is the cost of propane (for the burner), in $/gallon
Public EnergyCost As Single 'this is the cost of the energy required by the Fan and the Heater, $
Public ShrinkCost As Single 'this is the cost of overdrying (from the desired final MC) and DML,$
Public DesiredFinMC As Single 'this is the desired final moisture content (%, wb) to conpute the shrink cost
Public FanCost As Single 'this is the total cost of the fan, $
Public HeaterGalons As Single 'this is the total gallons of propane consummed (3414 BTU/KW and 92000 BTU/Gallon), gallons
Public HeaterCost As Single 'this is the total cost of the heater, $
Public DMLCost As Single 'this is te cost related to DML, $
Public DMLBin As Single 'this is the average dml of the bin used for cost calculations, %
Public OverdryingCost As Single ' this is the cost related to the overdrying of the grain, $
Public HeaterType As Boolean 'this is the type of heater considered, true=propane gas and false=electrical
Public FinalTonGrain As Single 'this is the final tons of grains at the desired final MC
Public InTonnes As Single 'thi is the initial number of tonnes of grain in the bin, used for drying cost computations
Public AvgFinMC As Single 'this is the average final mc of the grain, used for drying cost computation

Public StWeatherFile As String 'this string carries the information about the weather file used in the simulation
Public StStrategy As String 'this string carries the information about the strategy selected for the simulation
Public StBin As String 'this string carries the information about the dimensions of the bin used for the simulation
Public StGrainEMC As String 'this string carries the information about the EMC parameters used in the simulation
Public StAirflow As String 'this string carries the information af the airflow rate of the simulation
Public StGrain As String 'this string carries the information aboput the initial grain condition (MC and T)
Public StStartDate As String 'this string carries the information about the start date of the simulation (month and day)
Public StStopSimulation As String 'this string carries the information about the stop criteria for the simulation
Public StRunInfo As String ' this string combines the information of the 8 strings above





Sub LoadList(LineIndex, ArrGrain)

' to populate the list of possible grains to select
m = 0
For m = 0 To (LineIndex - 1)
    Combo1.AddItem ArrGrain(m, 0)
Next m
End Sub



Sub ReadGrainList(FileName)
' read the grain list and load all the parameters in an doble dimension array
Dim ArrComponents() As String
Dim GrainID As Integer
Dim LineFromFile As String

LineIndex = 0
j = 0
Open FileName For Input As #1
Do Until EOF(1)
Line Input #1, LineFromFile
ArrComponents() = Split(LineFromFile, vbTab)
For j = 0 To 13
    ArrGrain(LineIndex, j) = ArrComponents(j)
Next j
LineIndex = LineIndex + 1
Loop

Close #1

' to populate the list of possible grains to select
m = 0
For m = 0 To (LineIndex - 1)
    F_GrainEdit.Combo1.AddItem ArrGrain(m, 0)
Next m

End Sub

Sub ReadGrainList1(FileName)
' read the grain list and load all the parameters in an doble dimension array
Dim ArrComponents() As String
Dim GrainID As Integer
Dim LineFromFile As String

LineIndex = 0
j = 0
Open FileName For Input As #1
Do Until EOF(1)
Line Input #1, LineFromFile
ArrComponents() = Split(LineFromFile, vbTab)
For j = 0 To 14
    ArrGrain(LineIndex, j) = ArrComponents(j)
Next j
LineIndex = LineIndex + 1
Loop

Close #1
' to populate the list of possible grains to select
m = 0
For m = 0 To (LineIndex - 1)
    F_Main.Co_SelectGrain.AddItem ArrGrain(m, 0)
Next m

End Sub

Sub DisplayGrainInfo(ArrGrain, GrainIndex)
' print the information corresponding to the selected grain into the text boxes
F_GrainEdit.Tb_GrainName = ArrGrain(GrainIndex, 0)
F_GrainEdit.Tb_EMC_D_A = ArrGrain(GrainIndex, 1)
F_GrainEdit.Tb_EMC_D_B = ArrGrain(GrainIndex, 2)
F_GrainEdit.Tb_EMC_D_C = ArrGrain(GrainIndex, 3)
F_GrainEdit.Tb_EMC_R_A = ArrGrain(GrainIndex, 4)
F_GrainEdit.Tb_EMC_R_B = ArrGrain(GrainIndex, 5)
F_GrainEdit.Tb_EMC_R_C = ArrGrain(GrainIndex, 6)
F_GrainEdit.Tb_Dens_A = ArrGrain(GrainIndex, 7)
F_GrainEdit.Tb_Dens_B = ArrGrain(GrainIndex, 8)
F_GrainEdit.Tb_Dens_C = ArrGrain(GrainIndex, 9)
F_GrainEdit.Tb_SpHeat_A = ArrGrain(GrainIndex, 10)
F_GrainEdit.Tb_SpHeat_B = ArrGrain(GrainIndex, 11)
F_GrainEdit.Tb_AirfRes_A = ArrGrain(GrainIndex, 12)
F_GrainEdit.Tb_AirfRes_B = ArrGrain(GrainIndex, 13)
If ArrGrain(GrainIndex, 14) = "0" Then
    F_GrainEdit.Opt_No = True
    F_GrainEdit.Opt_Yes = False
Else
    F_GrainEdit.Opt_No = False
    F_GrainEdit.Opt_Yes = True
End If

End Sub

Sub UpdateF_GrainEdit()
If F_GrainChoice.Option1 = True Then
    F_GrainEdit.Label1.Caption = "Add Parameters for a New Grain"
    F_GrainEdit.Combo1.Visible = False
    F_GrainEdit.Label23.Visible = False
    F_GrainEdit.Command1.Visible = True
    F_GrainEdit.Command1.Caption = "Add New Grain"
End If
If F_GrainChoice.Option2 = True Then
    F_GrainEdit.Label1.Caption = "Edit Parameters of an Existing Grain"
    F_GrainEdit.Combo1.Visible = True
    F_GrainEdit.Label23.Visible = True
    F_GrainEdit.Command1.Caption = "Edit Existing Grain"
    F_GrainEdit.Command1.Visible = True
End If
If F_GrainChoice.Option3 = True Then
    F_GrainEdit.Label1.Caption = "View Parameters of an Existing Grain"
    F_GrainEdit.Combo1.Visible = True
    F_GrainEdit.Label23.Visible = True
    F_GrainEdit.Command1.Visible = False
End If

End Sub

Public Function ErrorF_GrainEdit() As Boolean
'this fuction checks for the information provided to add a new grain or edit an existing grain
'of F_GrainEdit. Returns a true value if all the requires text box are not empty, and a false value if any of
'the required textboxes is empty

    If F_GrainEdit.Tb_GrainName = "" Then
        ErrorF_GrainEdit = False
    Else
        If F_GrainEdit.Tb_EMC_D_A = "" Then
            ErrorF_GrainEdit = False
        Else
            If F_GrainEdit.Tb_EMC_D_B = "" Then
                ErrorF_GrainEdit = False
            Else
                If F_GrainEdit.Tb_EMC_D_C = "" Then
                    ErrorF_GrainEdit = False
                Else
                    If F_GrainEdit.Tb_EMC_R_A = "" Then
                        ErrorF_GrainEdit = False
                    Else
                        If F_GrainEdit.Tb_EMC_R_B = "" Then
                            ErrorF_GrainEdit = False
                        Else
                            If F_GrainEdit.Tb_EMC_D_C = "" Then
                                ErrorF_GrainEdit = False
                            Else
                                If F_GrainEdit.Tb_Dens_A = "" Then
                                    ErrorF_GrainEdit = False
                                Else
                                    If F_GrainEdit.Tb_Dens_B = "" Then
                                        ErrorF_GrainEdit = False
                                    Else
                                        If F_GrainEdit.Tb_Dens_C = "" Then
                                            ErrorF_GrainEdit = False
                                        Else
                                            If F_GrainEdit.Tb_SpHeat_A = "" Then
                                                ErrorF_GrainEdit = False
                                            Else
                                                If F_GrainEdit.Tb_SpHeat_B = "" Then
                                                    ErrorF_GrainEdit = False
                                                Else
                                                    If F_GrainEdit.Tb_AirfRes_A = "" Then
                                                        ErrorF_GrainEdit = False
                                                    Else
                                                        If F_GrainEdit.Tb_AirfRes_B = "" Then
                                                            ErrorF_GrainEdit = False
                                                        Else
                                                             ErrorF_GrainEdit = True
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If


End Function
' this sub is to set the initial grain layer depth
Sub SetGrainLayerDepth(Layers, Height, SingleLayerDepth)
i = 0
For i = 0 To (Layers - 1)
    LayerDepth(i) = SingleLayerDepth
Next i
LayerDepth(Layers) = Height - Int(Height)
End Sub

Public Function ComputeGrainDensity(GrainIndex, GrainMC1, ArrGrain)
' this function computes the density of the grain (as function of grainMC) based on the ASAED241.4 standard
'Grain MC is the moisture content of the grin, dec, w.b
' Grain density is the density of the grain in kg/m3
    ComputeGrainDensity = ArrGrain(GrainIndex, 7) - ArrGrain(GrainIndex, 8) * (GrainMC1 / 100) + ArrGrain(GrainIndex, 9) * (GrainMC1 / 100) * (GrainMC1 / 100)
End Function

Public Function AirFlowResistance(GrainIndex, ArrGrain, TotalAirflow, BinArea, PackingFactor) As Double
'this function computes the pressure drop per meter according to a given airflow rate (m3/m2/sec)
' for a specific grain, Pa/m
AirFlowResistance = (ArrGrain(GrainIndex, 12) * ((TotalAirflow / 60 / BinArea) * (TotalAirflow / 60 / BinArea)) / Log(1 + ArrGrain(GrainIndex, 13) * (TotalAirflow / 60 / BinArea))) * PackingFactor
End Function


Public Function CompFanPreWarm(AirfResistance_Wa)
' this function computes the fan prewarming based on airflow resistance in inches of water
' fan prewarming is in ºC, and assumes that 0.28ºC of temp increase for each inch of water of
' static pressure
CompFanPreWarm = 0.28 * AirfResistance_Wa
End Function

Public Function CompFanPower(TotalAirflow, AirfResistance_Wa, FanEfficiency)
'this function computes the fan power required for the given airflow rate (cfm) and the given
'static pressure (inch of water) for the given fan efficiency.
'the output is in HP
CompFanPower = (TotalAirflow / cfmTom3min * AirfResistance_Wa) / (63.46 * FanEfficiency)
End Function

'declare functions
' to convert ºC to ºK
' TC is temperature in ºC
' TK is temperature in ºK
Public Function Kelvin(TC) As Single
    Kelvin = TC + 273.15
End Function
' compute saturated vapour pressure
' TK is temperature, ºK
' Ps is saturated vapour pressure
Public Function Sat_press(TK) As Single
    Sat_press = R1 * (Exp((a1 + b1 * TK + c1 * ((TK) ^ 2) + d1 * ((TK) ^ 3) + e1 * ((TK) ^ 4)) / (f1 * TK - g1 * ((TK) ^ 2))))      ' compute Pse
End Function
'compute the vapor pressure of the air
'Ps is the saturated vapor pressure of the air
'RH is the relative humidity of the air, %
Public Function CompVaporPress(Ps, RH)
    CompVaporPress = Ps * RH / 100
End Function

' compute air density in kg/m3
'PV is the vapor pressure of the air
'TK is the temperature of the air, ºK
Public Function CompAirDens(Pv, TK)
    CompAirDens = (Pa - Pv) / (287 * TK)
End Function
'compute asbsolute humidity (H) of the air, g of water/kg of dry air
'Pv is the vapor pressure of the air
'Pa is the atmospheric pressure of the air
Public Function CompAbsHum(Pv)
CompAbsHum = 0.6219 * Pv / (Pa - Pv)
End Function

'compute the specific heat of the grain based on the grain moisture content according to ASAE D243.4
' grain specific heat is in kJ/(kg*ºK)
' grain MC is in %, w.b
Public Function CompGrainSpHeat(ArrGrain, GrainIndex, GrainMC)
CompGrainSpHeat = ArrGrain(GrainIndex, 10) + ArrGrain(GrainIndex, 11) * GrainMC
End Function

'compute the power of the heater based on the temperature increase and the airflow
'TC is the temperature of the air, ºC
Public Function CompBurnerPower(Tinc_C, TC, RH, TotalAirflow)
    TK = Kelvin(TC)
    Ps = Sat_press(TK)
    Pv = CompVaporPress(Ps, RH)
    AbsHum = CompAbsHum(Pv)
    CompBurnerPower = ((ca + cv * AbsHum) * (Tinc_C) * (TotalAirflow * 60 * CompAirDens(Pv, TK))) * KjToKWH
End Function

'compute the Temperature Multiplier for the DML equation
'the temperature multiplier is computed based on the ASAE Standard (X535) in revision (12-22-04)
'GT is grain temperature, ºC
'GMC is grain moisture content, %, wb
Public Function CompTempMult(GT, GMC)

    If GT < 15.6 Then
        CompTempMult = 128.389 * Exp(-4.86 * (1.8 * GT + 32) / 60)
    ElseIf GT < 26.7 Then
        If GMC < 19 Then
            CompTempMult = 32.3 * Exp(-3.48 * (1.8 * GT + 32) / 60)
        ElseIf GMC < 28 Then
            CompTempMult = Exp(-0.00493277 + (0.05 * (1.8 * GT + 32) - 3) * (Log(0.0795012 + 0.012315 * GMC)))
        Else
            CompTempMult = Exp(2.56683 - 0.0428628 * (1.8 * GT + 32))
        End If
    Else
        If GMC < 19 Then
            CompTempMult = 32.3 * Exp(-3.48 * (1.8 * GT + 32) / 60)
        ElseIf GMC < 28 Then
            CompTempMult = 32.3 * Exp(-3.48 * (1.8 * GT + 32) / 60 + ((GMC - 19) / 100 * Exp(0.61 * (1.8 * GT - 28) / 60)))
        Else
            CompTempMult = 32.3 * Exp(-3.48 * (1.8 * GT + 32) / 60) + 0.09 * Exp(0.61 * (1.8 * GT - 28) / 60)
        End If
    End If
End Function

'compute the Moisture Multiplier for the DML equation
'the moisture multiplier is computed based on the ASAE Standard (X535) in revision (12-22-04)
'GMC is grain moisture content, %, wb
Public Function CompMoistMult(GMC)
CompMoistMult = 0.103 * (Exp(455 / ((GMC * 100 / (100 - GMC)) ^ 1.53)) - (0.00845 * (GMC * 100 / (100 - GMC))) + 1.558)
End Function

'compute the Damage Multiplier for the DML equation
'the damage multiplier is computed based on the ASAE Standard (X535) in revision (12-22-04)
'GDamage is the percentage of damaged grain, %
Public Function CompDamageMult(GDamage)
CompDamageMult = 2.08 * Exp(-0.0239 * GDamage)
End Function
' compute the equivalent storage time for the DML equation
'the equivalent storage time is computed on the basis of the Temperature, MC, grain Damage, grain Gentetics and Fungicide Multipliers
'the results is in hours
Public Function CompEqStorageTime(GT, GMC, GDamage, DML_Mult_Fungicide, DML_Mult_Genetics)
DML_Mult_Temp = CompTempMult(GT, GMC)
DML_Mult_MC = CompMoistMult(GMC)
DML_Mult_Damage = CompDamageMult(GDamage)
CompEqStorageTime = 1 / (DML_Mult_Temp * DML_Mult_MC * DML_Mult_Damage * DML_Mult_Fungicide * DML_Mult_Genetics)
End Function

'compute the CO2 production of the stored corn
'compute the grain DML in grams of CO2 produced per kilogram of initial dry matter
' CompCO2Prod is computed according to the procedure proposed by Saul and Steel (1966)
Public Function CompCO2Prod(GT, GMC, GDamage, DML_Mult_Fungicide, DML_Mult_Genetics)
tr = CompEqStorageTime(GT, GMC, GDamage, DML_Mult_Fungicide, DML_Mult_Genetics)
CompCO2Prod = 1.3 * (Exp(0.006 * tr) - 1) + 0.015 * tr
End Function

' compute the temperature increase (ºC) in the grain mass due to the heat generated during respiration
'the computing is based on the following relationship:
'respiration of 1 mol of glucose (180 g) produces: 264 g of CO2; 108 g of water; and 2816 kJoules
'based on this relationship, per gram of glucose respired: 1.47 g of CO2; 0.6 g of water; and 15.64 kJoules

Public Function CompDML_TempInc(DML_GramsPerKg, ArrGrain, GrainIndex, MC_WB)
CompDML_TempInc = (15.64 * DML_GramsPerKg) / (ArrGrain(GrainIndex, 10) + ArrGrain(GrainIndex, 11) * MC_WB)
End Function


' compute the desorption (drying) EMC with the modified Chung-Pfost equation
' TC is temperature in ºC
' RH is the air relative humidity, decimal
' EMC is the equilibrium moisture content, %, db
Public Function CF_EMC_D(TC, RH, ArrGrain, GrainIndex) As Single
    If TC < -30 Then TC = -30 'chung-pfost model does not work with temperaures below -30ºC
    CF_EMC_D = (Log(Log(RH) * -1 * ((TC + ArrGrain(GrainIndex, 3)) / ArrGrain(GrainIndex, 1)))) / -ArrGrain(GrainIndex, 2)
End Function
' compute the adsorption (re-wetting) EMC with the modified Chung-Pfost equation
' TC is temperature in ºC
' RH is the air relative humidity, decimal
' EMC is the equilibrium moisture content, %, db
Public Function CF_EMC_R(TC, RH, ArrGrain, GrainIndex) As Single
    If TC < -30 Then TC = -30 'chung-pfost model does not work with temperaures below -30ºC
    CF_EMC_R = (Log(Log(RH) * -1 * ((TC + ArrGrain(GrainIndex, 6)) / ArrGrain(GrainIndex, 4)))) / -ArrGrain(GrainIndex, 5)
End Function


' compute the desorption (drying) ERH with the modified Chung-Pfost equation
' TC is temperature in ºC
' RH is the equilibrium relative humidity, decimal
' MC is the grain moisture content, %, db
Public Function CF_ERH_D(TC, MC, ArrGrain, GrainIndex) As Single
    If TC < -30 Then TC = -30 'chung-pfost model does not work with temperaures below -30ºC
    CF_ERH_D = Exp(-(ArrGrain(GrainIndex, 1) / (TC + ArrGrain(GrainIndex, 3))) * Exp(-ArrGrain(GrainIndex, 2) * MC))
End Function

' compute the adsorption (re-wetting) ERH with the modified Chung-Pfost equation
' TC is temperature in ºC
' RH is the equilibrium relative humidity, decimal
' MC is the grain moisture content, %, db
Public Function CF_ERH_R(TC, MC, ArrGrain, GrainIndex) As Single
    If TC < -30 Then TC = -30 'chung-pfost model does not work with temperaures below -30ºC
    CF_ERH_R = Exp(-(ArrGrain(GrainIndex, 4) / (TC + ArrGrain(GrainIndex, 6))) * Exp(-ArrGrain(GrainIndex, 5) * MC))
End Function

' convert moisture content from wet basis to dry basis, decimal
' Mwb: moisture content wet basis, decimal
' Mdb: moisture content dry basis, decimal
Public Function Mwb_Mdb(Mwb) As Single
    Mwb_Mdb = Mwb / (1 - Mwb)
End Function

' convert moisture content from dry basis to wet basis, decimal
' Mwb: moisture content wet basis, decimal
' Mdb: moisture content dry basis, decimal
Public Function Mdb_Mwb(Mdb) As Single
    Mdb_Mwb = Mdb / (1 + Mdb)
End Function
' this function is to adjust the grain layer depth according to the change in moisture content during the timestep
Public Function CompLayerDepthF(GrainIndex, MCInitial, MCFinal, ArrGrain, LayerDepth0)
Dim DensityI As Single 'this is the density of the cor at the begining of the time step
Dim CornMassI As Single 'this is the mass of corn at MC initial
Dim DryMassCorn As Single 'this is the dy mass of corn in the layer
Dim DensityF As Single 'this is the density of the corn at the end of the time step (at MC final)
Dim CorMassF As Single 'this is the mass of corn at MC final
'CompLayerDepth is the depth of the corn layer at the end of the time step, adjusted by change in MC and grain density
DensityI = ArrGrain(GrainIndex, 7) - ArrGrain(GrainIndex, 8) * (MCInitial / 100) + ArrGrain(GrainIndex, 9) * (MCInitial / 100) * (MCInitial / 100)
CornMassI = DensityI * LayerDepth0
DryMassCorn = CornMassI * (1 - MCInitial / 100)
DensityF = ArrGrain(GrainIndex, 7) - ArrGrain(GrainIndex, 8) * (MCFinal / 100) + ArrGrain(GrainIndex, 9) * (MCFinal / 100) * (MCFinal / 100)
CornMassF = DryMassCorn / (1 - MCFinal / 100)
CompLayerDepthF = CornMassF / DensityF
End Function
Sub Drying(M0W, G0C, AirTemp_C, AirRH, dx, ArrGrain, GrainIndex, va, FanStatus)

' Comments at 8/5/2004 by REB
' This drying model was develop on the basis of the Thompson et al., 1972 equilibrium model. Adapted from Romualdo Martinez (2001)
' thesis: Modelling and Simulation of the Two Stage Rice Drying System in the Philippines,
' Hohenheim, 2001.




ReDim T0C(31) As Single     ' initial air temperature; ºC
ReDim T0K(31) As Single     ' initial air temperature; ºK
ReDim RH0(31) As Single     ' initial air relative humidity; %
ReDim Ps0(31) As Single     ' initial saturated vapour pressure; ??
ReDim Pv0(31) As Single     ' initial vapour pressure at given air condition; ??
ReDim H0(31) As Single      ' initial absolute humidity of the air; kg/kg

ReDim M0d(31) As Single     ' initial grain moisture content; %, db
ReDim M0d1(31) As Single    ' grain moisture content at the beginning of the simulation; %, db

ReDim G0K(31) As Single     ' initial grain temperature; ºK
ReDim R(31) As Single       ' dry mater to dry air ratio; kg/kg
ReDim GD(31) As Single      ' grain density; kg/m3
ReDim ad(31) As Single      ' air density; kg/m3
ReDim Cg1(31) As Single     ' specific heat of corn; kJ/(kgºK)
ReDim Cg(31) As Single      ' specific heat of corn in relation to the dry mater to dry air ratio; kJ/(kgºK)
ReDim Hf1(50) As Single     ' absolute humidity used to find feasible final RH conditions
ReDim Mfd1(50) As Single
ReDim TfC1(50) As Single
ReDim TfK1(50) As Single
ReDim Psf1(50) As Single
ReDim Pvf1(50) As Single
ReDim RHf1A(50) As Single
ReDim RHf1B(50) As Single
ReDim dRH1(50) As Single

ReDim Mfd(31) As Single     ' final grain moisture content; %, db
ReDim dH(31) As Single      ' change in absolute humidity of the air; kg/kg
ReDim Hf(31) As Single      ' final absolute humidity of the air; kg/kg
ReDim Rgas(31) As Single    ' R constant of water vapour at final conditions; kJ/(kgºK)
ReDim dL(31) As Single      ' Latent heat of vaporization of water; KJ/(kgºK)
ReDim TfC(31) As Single     ' final air temperature; ºC
ReDim TfK(31) As Single     ' final air temperature; ºK
ReDim Psf(31) As Single     ' final Ps
ReDim Pvf(31) As Single     ' final Pv
ReDim RHf(31) As Single     ' final RH, %

Dim Process As Integer      'indicates if this is a drying or rewetting process: 0= drying, 1=rewetting, 2= average between drying and rewetting
Dim TimeStepEMC_D As Single   'this is the computed drying EMC for the timested, used to determine if during the current timestep a drying or rewetting equation should be used
Dim TimeStepEMC_R As Single   'this is the computed rewetting EMC for the timested, used to determine if during the current timestep a drying or rewetting equation should be used
t = 1 'timestep is 1 hour
      
    i = 0       ' set array index to 1 to start in the first thin layer
    For i = 0 To (NumberOfLayers - 1)
        
        If i = 0 Then           ' to set the T0 to ambient for thin layer 1, or to Tf of the layer before
            T0C(0) = AirTemp_C
            RH0(0) = AirRH
        Else
            T0C(i) = TfC(i - 1)
            RH0(i) = RHf(i - 1)
        End If
        ' to compute moisture and temperature change only if fan is on
        If FanStatus = True Then
            M0d(i) = Mwb_Mdb(M0W(i) / 100) * 100
           T0K(i) = Kelvin(T0C(i))
            Ps0(i) = Sat_press(T0K(i))
            Pv0(i) = CompVaporPress(Ps0(i), RH0(i))      ' compute Pvs
            H0(i) = CompAbsHum(Pv0(i))     ' compute absolute humidity 0
            GD(i) = ArrGrain(GrainIndex, 7) - ArrGrain(GrainIndex, 8) * (M0W(i) / 100) + ArrGrain(GrainIndex, 9) * (M0W(i) / 100) * (M0W(i) / 100)  ' compute grain density
            ad(i) = CompAirDens(Pv0(i), T0K(i))      ' compute air density
            R(i) = (GD(i) * dx(i) * (1 - M0W(i) / 100)) / (va * t * 3600 * ad(i))  'compute R: dry matter to dry air ratio, kg/kg
            Cg1(i) = CompGrainSpHeat(ArrGrain, GrainIndex, M0W(i))        ' compute specific heat of grain
            Cg(i) = R(i) * Cg1(i)                   ' compute specific heat of grain, converted to Kj/(kg ar K)
        
            'detremine if the current step if drying or rewetting
            'a third possibility is also considered, this is when the grain MC of the layer is in between the drying EMC and the rewetting EMC
            'in this case the average of the drying and re-wetting curves is used
            Process = 0
            TimeStepEMC_D = (Mdb_Mwb(CF_EMC_D(T0C(i), RH0(i) / 100, ArrGrain, GrainIndex) / 100)) * 100
            TimeStepEMC_R = (Mdb_Mwb(CF_EMC_R(T0C(i), RH0(i) / 100, ArrGrain, GrainIndex) / 100)) * 100
            If TimeStepEMC_D = TimeStepEMC_R Then
                Process = 0 'in case that drying and rewetting parameters are the same
            Else
                If M0W(i) >= TimeStepEMC_D Then
                    Process = 0
                ElseIf M0W(i) >= TimeStepEMC_R Then
                    Process = 1
                Else:
                    Process = 2
                End If
            End If
        
            flag = 1
            n = 1
            Do While flag = 1
                If n = 1 Then
                    Hf1(n) = H0(i) * 0.99
                ElseIf n = 2 Then
                    Hf1(n) = H0(i) * 1.01
                Else
                    Hf1(n) = Hf1(n - 1) - dRH1(n - 1) * ((Hf1(n - 2) - Hf1(n - 1)) / (dRH1(n - 2) - dRH1(n - 1)))
                End If
                Mfd1(n) = M0d(i) - 100 * (Hf1(n) - H0(i)) / R(i)
                TfC1(n) = ((ca + cv * H0(i)) * T0C(i) - (Hf1(n) - H0(i)) * (hv - cw * G0C(i)) + Cg(i) * G0C(i)) / (ca + cv * Hf1(n) + Cg(i))
                TfK1(n) = Kelvin(TfC1(n))
                If Process = 0 Then
                    RHf1A(n) = CF_ERH_D(TfC1(n), Mfd1(n), ArrGrain, GrainIndex) * 100
                ElseIf Process = 1 Then
                    RHf1A(n) = CF_ERH_R(TfC1(n), Mfd1(n), ArrGrain, GrainIndex) * 100
                Else
                    RHf1A(n) = ((CF_ERH_D(TfC1(n), Mfd1(n), ArrGrain, GrainIndex) * 100) + (CF_ERH_R(TfC1(n), Mfd1(n), ArrGrain, GrainIndex) * 100)) / 2
                End If
                Psf1(n) = Sat_press(TfK1(n))
                Pvf1(n) = Hf1(n) * Pa / (0.6219 + Hf1(n))
                RHf1B(n) = Pvf1(n) / Psf1(n) * 100
                dRH1(n) = RHf1A(n) - RHf1B(n)
                If Abs(dRH1(n)) < 0.001 Then
                    flag = 0
                    Hf(i) = Hf1(n)
                    Mfd(i) = Mfd1(n)
                    MfW(i) = Mdb_Mwb(Mfd(i) / 100) * 100
                    TfC(i) = TfC1(n)
                    GfC(i) = TfC(i)
                    RHf(i) = RHf1B(n)
                    If RHf(i) >= 100 Then RHf(i) = 99
                Else
                    flag = 1
                End If
            n = n + 1
            Loop
    
            'update layer depth at the end of the timestep
            dxf(i) = CompLayerDepthF(GrainIndex, M0W(i), MfW(i), ArrGrain, dx(i))
        End If
    'compute DML of the layer if grain is corn
    If ArrGrain(GrainIndex, 14) = "1" Then
        GT = GfC(i)
        GMC = MfW(i)
        DML_GramsPerKg = CompCO2Prod(GT, GMC, GDamage, DML_Mult_Fungicide, DML_Mult_Genetics) / 14.7
        LayerDML(i) = LayerDML(i) + DML_GramsPerKg
        DML_TempIncr_C = CompDML_TempInc(DML_GramsPerKg, ArrGrain, GrainIndex, GMC)
        TfC(i) = TfC(i) + DML_TempIncr_C
    End If
    Next i

End Sub

Sub PrintToEndFile(PrintVar_End, j, FileNumber, PrintYear, PrintHeading_End)
    
    If j = 0 Then       ' to print the headings of the table
        Print #FileNumber, PrintHeading_End
    End If
    
    Print #FileNumber, PrintYear; vbTab; PrintVar_End

End Sub


Sub PrintToFanFile(PrintVar_Fan, j, FileNumber, PrintTime, PrintHeading)
    
    If j = 0 Then       ' to print the headings of the table
        Print #FileNumber, PrintHeading
    End If
    
    Print #FileNumber, PrintTime; vbTab; PrintVar_Fan

End Sub


Sub PrintToFile(PrintVar, m, j, FileNumber, PrintTime, PrintHeading)
    
    If j = 0 Then       ' to print the headings of the table
        Print #FileNumber, PrintHeading
    End If
    
    i = 0           ' to print the values on the table
    Valuestrim = ""
    For i = 0 To (m + 3)
        If i <= m Then
            VarValue = Format(PrintVar(i), "##0.00")
        ElseIf i = (m + 1) Then
            VarValue = Format(ArrayAvg(PrintVar, 0, m), "##0.00")
        ElseIf i = (m + 2) Then
            VarValue = Format(ArrayMin(PrintVar, 0, m), "##0.00")
        Else
            VarValue = Format(ArrayMax(PrintVar, 0, m), "##0.00")
        End If
        Valuestrim = Valuestrim & VarValue & vbTab
    Next i
    Print #FileNumber, PrintTime; vbTab; Valuestrim

End Sub
Sub FixInlet_Strat(AirTemp_C, AirRH)
Counter = 0
'set the initial temperature and MC conditions for each grain layer
i = 0
For i = 0 To (NumberOfLayers - 1)
    GrainMC_WB_C(i) = GrainMC_In_WB_C(i)
    GrainMC_WB_S(i) = GrainMC_In_WB_S(i)
    GrainTemp_C_C(i) = GrainTemp_In_C_C(i)
    GrainTemp_C_S(i) = GrainTemp_In_C_S(i)
Next i
'set the initial grain layer depth
i = 0
For i = 0 To (NumberOfLayers - 1)
    LayerDepth_C(i) = LayerDepth(i)
    LayerDepth_S(i) = LayerDepth(i)
Next i

'open all the output files
Open CurrentDir & "\output files\" & BaseName & "_mccenter.txt" For Output As #3
Open CurrentDir & "\output files\" & BaseName & "_mcside.txt" For Output As #4
Open CurrentDir & "\output files\" & BaseName & "_tempcenter.txt" For Output As #6
Open CurrentDir & "\output files\" & BaseName & "_tempside.txt" For Output As #7
Open CurrentDir & "\output files\" & BaseName & "_dmlcenter.txt" For Output As #9
Open CurrentDir & "\output files\" & BaseName & "_dmlside.txt" For Output As #10
StopSim = False
Do While StopSim = False
'compute drying for the center of the bin
Call Drying(GrainMC_WB_C, GrainTemp_C_C, AirTemp_C, AirRH, LayerDepth_C, ArrGrain, GrainIndex, AirVel_C, FanStatus)
'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the center of the bin
i = 0
For i = 0 To (NumberOfLayers - 1)
    GrainMC_WB_C(i) = MfW(i)
    GrainTemp_C_C(i) = GfC(i)
    LayerDepth_C(i) = dxf(i)
    LayerDML_C(i) = LayerDML(i)
Next i
'write the headings for the output file
i = 0
textstrim = ""
m = NumberOfLayers - 1
For i = 0 To (m + 3)
    If i <= m Then
        varprint = "Layer" & (i + 1)
    ElseIf i = (m + 1) Then
        varprint = "Average"
    ElseIf i = (m + 2) Then
        varprint = "Minimun"
    Else
        varprint = "Maximum"
    End If
    textstrim = textstrim & varprint & vbTab
Next i
PrintHeading = "Hours" & vbTab & textstrim

'print in a file the moisture content of each layers at the center of the bin
FileNumber = 3
Call PrintToFile(GrainMC_WB_C, (NumberOfLayers - 1), Counter, FileNumber, Counter + 1, PrintHeading)
'print in a file the temperature of each layers at the center of the bin
FileNumber = 6
Call PrintToFile(GrainTemp_C_C, (NumberOfLayers - 1), Counter, FileNumber, Counter + 1, PrintHeading)
If ArrGrain(GrainIndex, 14) = "1" Then
    'print in a file the DML of each layers at the center of the bin
    FileNumber = 9
    Call PrintToFile(LayerDML_C, (NumberOfLayers - 1), Counter, FileNumber, Counter + 1, PrintHeading)
End If

'compute drying for the side of the bin
Call Drying(GrainMC_WB_S, GrainTemp_C_S, AirTemp_C, AirRH, LayerDepth_S, ArrGrain, GrainIndex, AirVel_S, FanStatus)
'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the side of the bin
i = 0
For i = 0 To (NumberOfLayers - 1)
    GrainMC_WB_S(i) = MfW(i)
    GrainTemp_C_S(i) = GfC(i)
    LayerDepth_S(i) = dxf(i)
    LayerDML_S(i) = LayerDML(i)
Next i
'print in a file the moisture content of each layers at the side of the bin
FileNumber = 4
Call PrintToFile(GrainMC_WB_S, (NumberOfLayers - 1), Counter, FileNumber, Counter + 1, PrintHeading)
'print in a file the temperature of each layers at the side of the bin
FileNumber = 7
Call PrintToFile(GrainTemp_C_S, (NumberOfLayers - 1), Counter, FileNumber, Counter + 1, PrintHeading)
If ArrGrain(GrainIndex, 14) = "1" Then
    'print in a file the DML of each layers at the side of the bin
    FileNumber = 10
    Call PrintToFile(LayerDML_S, (NumberOfLayers - 1), Counter, FileNumber, Counter + 1, PrintHeading)
End If

Call StopSimulation(MCCriteria, TempCriteria, DMLCriteria, GrainMC_WB_C, GrainMC_WB_S, GrainTemp_C_C, GrainTemp_C_S, LayerDML_C, LayerDML_S, Counter + 1, NumberOfLayers - 1, MCavg_Stop, MCmax_Stop, Tempavg_Stop, Tempmax_Stop, DMLavg_Stop, Hours_Stop)

Counter = Counter + 1
Loop

'close all open files
Close #3 'MC center
Close #4 'MC side
Close #6 'temp center
Close #7 'temp side
Close #9 'dml center
Close #10 'dml side

End Sub
Sub CNA_Strat()

'set variables to initial value at the begining of the simulation
HeaterStatus = False
'to set the temp increase demanded to the burnesr+0
Tinc_C = 0
' set all the arrays for the set of years simulation to "0"
If MultYears = True Then
    i = 0
    For i = 0 To UBound(S_Moisture_Avg)
        S_Moisture_Avg(i) = 0
        S_Moisture_Min(i) = 0
        S_Moisture_Max(i) = 0
        S_Temperature_Avg(i) = 0
        S_Temperature_Min(i) = 0
        S_Temperature_Max(i) = 0
        S_DML_Avg(i) = 0
        S_DML_Max(i) = 0
        S_DryingHs(i) = 0
        S_FanHs(i) = 0
        S_PerFanHs(i) = 0
        S_HeaterHs(i) = 0
        S_PerHeaterHs(i) = 0
        S_FanKWH(i) = 0
        S_HeaterKWH(i) = 0
        S_TotDryingCost(i) = 0
    Next i
End If

' set the year index (yi) to 0
yi = 0

' compute number of days from January 1 of 1960 to the first year month and day of the weather file
TotalDaysToFile = NumberOfDays(InitialFileYear, InitialFileMonth, InitialFileDay)

' compute number of days from January 1 of 1960 to the first year month and day of the analysis
TotalDaysStartSim = NumberOfDays(InitialSimYear, InitialSimMonth, InitialSimDay)

' compute the number of days since the beginin of the weather data to the begining of the analysis
DaysToStart = TotalDaysStartSim - TotalDaysToFile

' coumpute the number of hours (lines) between the begining of the weather file and the first simulation hour
FirstLineRead = (DaysToStart& * 24)

' compute number of days from January 1 of 1960 to the last year month and day of the analysis
TotalDaysFinishSim = NumberOfDays(FinalSimYear, FinalSimMonth, FinalSimDay)

' compute the position of the Last line to read in the file for the first year
DaysToFinish = (TotalDaysFinishSim - TotalDaysToFile) + 1

LastLineRead = (DaysToFinish * 24)

CurrentYear = InitialSimYear
il = 0
Open FileNameW For Input As #1

For CurrentYear = InitialSimYear To FMultYear 'this is the loop for the multiple simulation years

'these lines are for locating the first and last line to read in the file for the subsequent years
If CurrentYear = InitialSimYear Then
    FirstLineRead = FirstLineRead
    LastLineRead = LastLineRead
Else
    FirstLineRead = FirstLineRead + ((365 + Bisiesto1(CurrentYear)) * 24)
    LastLineRead = LastLineRead + ((365 + Bisiesto1(CurrentYear)) * 24)
End If
'these lines are to read all the previous lines in the file up to the first line used for the simulation
Do Until il = FirstLineRead
    Line Input #1, LineFromFile 'read this line from file
      
    il = il + 1
Loop 'loop to locate the first line to read

'check if only summary outouts were required
If S_OutputSummary = False Then
'open all the output files
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_mccenter" & ".txt" For Output As #3
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_mcside" & ".txt" For Output As #4
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_tempcenter" & ".txt" For Output As #6
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_tempside" & ".txt" For Output As #7
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_dmlcenter" & ".txt" For Output As #9
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_dmlside" & ".txt" For Output As #10
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_Fan" & ".txt" For Output As #11
End If
'initialize variables for the current year
FanRunHours = 0
HeaterRunHours = 0
Per_FanRun = 0
Per_HeaterRun = 0
FanKWH = 0
HeaterKWH = 0

'set the initial temperature and MC conditions for each grain layer
i = 0
For i = 0 To (NumberOfLayers - 1)
    GrainMC_WB_C(i) = GrainMC_In_WB_C(i)
    GrainMC_WB_S(i) = GrainMC_In_WB_S(i)
    GrainTemp_C_C(i) = GrainTemp_In_C_C(i)
    GrainTemp_C_S(i) = GrainTemp_In_C_S(i)
    LayerDML(i) = 0
    LayerDML_C(i) = 0
    LayerDML_S(i) = 0
Next i
'set the initial grain layer depth
i = 0
For i = 0 To (NumberOfLayers - 1)
    LayerDepth_C(i) = LayerDepth(i)
    LayerDepth_S(i) = LayerDepth(i)
Next i

'set counter to 0
Counter = 0
'set hous_stop to 0 for ending of the simulation based on final date
StopSim = False
Do While StopSim = False
    Line Input #1, LineFromFile 'read this line from file

'the string is converted into an array. Temperature corresponds to posicion 5, and RH
'to posicion 6 in the array

    
sItems() = Split(LineFromFile, vbTab)
    BadDataflag = 1
    FileTemp = sItems(TempColumn - 1)
    If FileTemp < -29.9 Or FileTemp > 40 Then BadDataflag = 0
    FileRH = sItems(RHColumn - 1)
    If FileRH = 100 Then FileRH = 99
    If FileRH < 1 Or FileRH > 100 Then BadDataflag = 0
    DateStamp = Format(sItems(1), "00") & "-" & Format(sItems(2), "00") & "-" & sItems(0)
    TimeStamp = Format(sItems(3), "00") & ":00"
    PrintTime = DateStamp & " " & TimeStamp
    
    If BadDataflag = 1 Then
        FileEMC_db = CF_EMC_D(FileTemp, FileRH / 100, ArrGrain, GrainIndex)
        FileEMC_wb = 100 * FileEMC_db / (100 + FileEMC_db)
        
        Select Case (FileTemp)
         Case MinTempSelect To MaxTempSelect
                Tempflag = 1
            Case Else
            Tempflag = 0
        End Select
    
        Select Case (FileRH)
            Case MinRHSelect To MaxRHSelect
                RHflag = 1
            Case Else
            RHflag = 0
        End Select
    
        Select Case (FileEMC_wb)
            Case MinEMCSelect To MaxEMCSelect
                EMCflag = 1
            Case Else
            EMCflag = 0
        End Select
    End If
        HourFlag = Tempflag * RHflag * EMCflag * BadDataflag
    If HourFlag = 1 Then
        FanStatus = True
        PlenumTemp_C = FileTemp + FanPreWarming_C
        PsAmbient = Sat_press(Kelvin(FileTemp))
        PvAmbient = CompVaporPress(PsAmbient, FileRH)
        PvPlenum = PvAmbient
        PsPlenum = Sat_press(Kelvin(PlenumTemp_C))
        PlenumRH = PvPlenum / PsPlenum * 100
        PlenumEMC_db = CF_EMC_D(PlenumTemp_C, PlenumRH / 100, ArrGrain, GrainIndex)
        PlenumEMC_wb = 100 * PlenumEMC_db / (100 + PlenumEMC_db)
    Else: FanStatus = False
    End If
    
    FanPrint = "Fan Off"
    HeaterPrint = "Heater Off"
    ' compute the fan runtime and percentaje of fan runtime
    If FanStatus = True Then
        FanRunHours = FanRunHours + 1
        FanPrint = "Fan On"
    End If
    Per_FanRun = Format(FanRunHours / (Counter + 1) * 100, "00.0")
    'compute the total energy consumtion of the fan (kwh) since the begining of the simulation for the current year
    FanKWH = FanPower_KW * FanRunHours
    ' compute the heater runtime and percentaje of heater runtime
    If HeaterStatus = True Then
        HeaterRunHours = HeaterRunHours + 1
        HeaterPrint = "Heater On"
    End If
    Per_HeaterRun = Format(HeaterRunHours / (Counter + 1) * 100, "00.0")
    'compute the total energy consumtion of the heater (kwh) since the begining of the simulation for the current year
    HeaterKWH = BurnerEnergy_KW * HeaterRunHours
    
    If S_OutputSummary = False Then
        'print the heading for the fan output file
        PrintHeading_Fan = "Date  /   Time" & vbTab & vbTab & "A Temp" & vbTab & "A RH" & vbTab & "A EMC" & vbTab & "Pl Temp" & vbTab & "Pl RH" & vbTab & "Pl EMC" & vbTab & "Fan St" & vbTab & "Heater St" & vbTab & "F Run T" & vbTab & "F Run %" & vbTab & "H Run T" & vbTab & "H Run %" & vbTab & "F KWH" & vbTab & "H KWH"
        'set values into the string to be printed in the fan output file for the current hour
        PrintVar_Fan = Format(FileTemp, "#0.0") & vbTab & Format(FileRH, "#0") & vbTab & Format(FileEMC_wb, "#0.0") & vbTab & Format(PlenumTemp_C, "#0.0") & vbTab & Format(PlenumRH, "#0") & vbTab & Format(PlenumEMC_wb, "#0.0") & vbTab & FanPrint & vbTab & HeaterPrint & vbTab & Format(FanRunHours, "0") & vbTab & Format(Per_FanRun, "0.0") & vbTab & Format(HeaterRunHours, "0") & vbTab & Format(Per_HeaterRun, "0.0") & vbTab & Format(FanKWH, "0") & vbTab & Format(HeaterKWH, "0")
        FileNumber = 11
        Call PrintToFanFile(PrintVar_Fan, Counter, FileNumber, PrintTime, PrintHeading_Fan)
    End If
    
    If HourFlag = 1 Then    'if all the temp, RH and EMC data are satisfied for the current hour, then run the drying sub for each layer
        'compute drying for the center of the bin
        Call Drying(GrainMC_WB_C, GrainTemp_C_C, PlenumTemp_C, PlenumRH, LayerDepth_C, ArrGrain, GrainIndex, AirVel_C, FanStatus)
        'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the center of the bin
        i = 0
        For i = 0 To (NumberOfLayers - 1)
            GrainMC_WB_C(i) = MfW(i)
            GrainTemp_C_C(i) = GfC(i)
            LayerDepth_C(i) = dxf(i)
            LayerDML_C(i) = LayerDML(i)
        Next i
        
        If S_OutputSummary = False Then
            'write the headings for the output file
            i = 0
            textstrim = ""
            m = NumberOfLayers - 1
            For i = 0 To (m + 3)
                If i <= m Then
                    varprint = "Layer" & (i + 1)
                ElseIf i = (m + 1) Then
                    varprint = "Average"
                ElseIf i = (m + 2) Then
                    varprint = "Minimun"
                Else
                    varprint = "Maximum"
                End If
                textstrim = textstrim & varprint & vbTab
            Next i
            PrintHeading = "Date  /  Time" & vbTab & vbTab & textstrim

            'print in a file the moisture content of each layers at the center of the bin
            FileNumber = 3
            Call PrintToFile(GrainMC_WB_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            'print in a file the temperature of each layers at the center of the bin
            FileNumber = 6
            Call PrintToFile(GrainTemp_C_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            If ArrGrain(GrainIndex, 14) = "1" Then
            'print in a file the DML of each layers at the center of the bin
                FileNumber = 9
                Call PrintToFile(LayerDML_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            End If
        End If
        'compute drying for the side of the bin
        Call Drying(GrainMC_WB_S, GrainTemp_C_S, PlenumTemp_C, PlenumRH, LayerDepth_S, ArrGrain, GrainIndex, AirVel_S, FanStatus)
        'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the side of the bin
        i = 0
        For i = 0 To (NumberOfLayers - 1)
            GrainMC_WB_S(i) = MfW(i)
            GrainTemp_C_S(i) = GfC(i)
            LayerDepth_S(i) = dxf(i)
            LayerDML_S(i) = LayerDML(i)
        Next i
        
        If S_OutputSummary = False Then
            'print in a file the moisture content of each layers at the side of the bin
            FileNumber = 4
            Call PrintToFile(GrainMC_WB_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            'print in a file the temperature of each layers at the side of the bin
            FileNumber = 7
            Call PrintToFile(GrainTemp_C_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            If ArrGrain(GrainIndex, 14) = "1" Then
            'print in a file the DML of each layers at the side of the bin
                FileNumber = 10
                Call PrintToFile(LayerDML_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            End If
        End If
    End If ' for hour flag


' to make the simulation for the current year stop based on final date
If DateCriteria = True Then
    If il >= LastLineRead - 1 Then
        Hours_Stop = 0
        CurrentHours = 1
    Else
        Hours_Stop = 1
        CurrentHours = 0
    End If
End If

Call StopSimulation(MCCriteria, TempCriteria, DMLCriteria, GrainMC_WB_C, GrainMC_WB_S, GrainTemp_C_C, GrainTemp_C_S, LayerDML_C, LayerDML_S, CurrentHours, NumberOfLayers - 1, MCavg_Stop, MCmax_Stop, Tempavg_Stop, Tempmax_Stop, DMLavg_Stop, Hours_Stop)

Counter = Counter + 1
il = il + 1
Loop 'loop for the year simulation

If S_OutputSummary = False Then
    'close all open files
    Close #3 'MC center
    Close #4 'MC side
    Close #6 'temp center
    Close #7 'temp side
    Close #9 'dml center
    Close #10 'dml side
    Close #11 'fan file
End If

' create the arrays for the multiple year simulation
If MultYears = True Then
MaxLayer = NumberOfLayers - 1
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = GrainMC_WB_C(i)
    Else
        ArrCriteria(i) = GrainMC_WB_S(i - (MaxLayer + 1))
    End If
Next i
        S_Moisture_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Moisture_Min(yi) = ArrayMin(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Moisture_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = GrainTemp_C_C(i)
    Else
        ArrCriteria(i) = GrainTemp_C_S(i - (MaxLayer + 1))
    End If
Next i
        S_Temperature_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Temperature_Min(yi) = ArrayMin(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Temperature_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = LayerDML_C(i)
    Else
        ArrCriteria(i) = LayerDML_S(i - (MaxLayer + 1))
    End If
Next i
        S_DML_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_DML_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
        
        S_DryingHs(yi) = Counter
        S_FanHs(yi) = FanRunHours
        S_PerFanHs(yi) = Per_FanRun
        S_HeaterHs(yi) = HeaterRunHours
        S_PerHeaterHs(yi) = Per_HeaterRun
        S_FanKWH(yi) = FanKWH
        S_HeaterKWH(yi) = HeaterKWH
        
        DMLBin = S_DML_Avg(yi)
        AvgFinMC = S_Moisture_Avg(yi)
        S_TotDryingCost(yi) = DryingCost(FanKWH, HeaterKWH, BinCapacity_t, DMLBin, AvgFinMC)
    End If

yi = yi + 1
Next CurrentYear 'loop for the multiple years simulation

If MultYears = True Then
Open CurrentDir & "\output files\" & BaseName & "_" & "_Summary.txt" For Output As #12

' print the summary of the simulation run settings in the end file
Print #12, StRunInfo
Print #12, ""
i = 0
For i = 0 To yi + 3
    If i < yi Then
        PrintHeading_End = "S Year" & vbTab & "$/ton" & vbTab & "c/bu" & vbTab & "MC Avg" & vbTab & "MC Min" & vbTab & "MC Max" & vbTab & "T Avg" & vbTab & "T Min" & vbTab & "T Max" & vbTab & "DML Avg" & vbTab & "DML Max" & vbTab & "Dr hs" & vbTab & "F hs" & vbTab & "F hs %" & vbTab & "H hs" & vbTab & "H hs %" & vbTab & "F KWH" & vbTab & "H KWH"
        PrintYear = (InitialSimYear + i)
        PrintVar_End = Format(S_TotDryingCost(i), "0.00") & vbTab & Format(S_TotDryingCost(i) / 0.4, "0.00") & vbTab & Format(S_Moisture_Avg(i), "0.00") & vbTab & Format(S_Moisture_Min(i), "0.00") & vbTab & Format(S_Moisture_Max(i), "0.00") & vbTab & Format(S_Temperature_Avg(i), "0.00") & vbTab & Format(S_Temperature_Min(i), "0.00") & vbTab & Format(S_Temperature_Max(i), "0.00") & vbTab & Format(S_DML_Avg(i), "0.00") & vbTab & Format(S_DML_Max(i), "0.00") & vbTab & Format(S_DryingHs(i), "0") & vbTab & Format(S_FanHs(i), "0") & vbTab & Format(S_PerFanHs(i), "0.0") & vbTab & Format(S_HeaterHs(i), "0") & vbTab & Format(S_PerHeaterHs(i), "0.0") & vbTab & Format(S_FanKWH(i), "0") & vbTab & Format(S_HeaterKWH(i), "0")
    ElseIf i = yi Then
        PrintYear = "Avg"
        PrintVar_End = Format(ArrayAvg(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayAvg(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayAvg(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayAvg(S_HeaterKWH, 0, yi - 1), "0")
    ElseIf i = yi + 1 Then
        PrintYear = "Min"
        PrintVar_End = Format(ArrayMin(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayMin(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMin(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMin(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayMin(S_HeaterKWH, 0, yi - 1), "0")
    ElseIf i = yi + 2 Then
        PrintYear = "Max"
        PrintVar_End = Format(ArrayMax(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayMax(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMax(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMax(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayMax(S_HeaterKWH, 0, yi - 1), "0")
        ElseIf i = yi + 2 Then
    ElseIf i = yi + 3 Then
        PrintYear = "StD"
        PrintVar_End = Format(ArrayStdDev(S_TotDryingCost, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_TotDryingCost, True, True, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Min, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Min, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DML_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DML_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DryingHs, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_FanHs, True, True, yi - 1), "0") _
                        & vbTab & Format(ArrayStdDev(S_PerFanHs, True, True, yi - 1), "0.0") & vbTab & Format(ArrayStdDev(S_HeaterHs, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_PerHeaterHs, True, True, yi - 1), "0.0") & vbTab & Format(ArrayStdDev(S_FanKWH, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_HeaterKWH, True, True, yi - 1), "0")
    End If
Call PrintToEndFile(PrintVar_End, i, 12, PrintYear, PrintHeading_End)
Next i
End If

Close #12
Close #1
End Sub

Sub ConstHeat_Strat()

' set all the arrays for the set of years simulation to "0"
If MultYears = True Then
    i = 0
    For i = 0 To UBound(S_Moisture_Avg)
        S_Moisture_Avg(i) = 0
        S_Moisture_Min(i) = 0
        S_Moisture_Max(i) = 0
        S_Temperature_Avg(i) = 0
        S_Temperature_Min(i) = 0
        S_Temperature_Max(i) = 0
        S_DML_Avg(i) = 0
        S_DML_Max(i) = 0
        S_DryingHs(i) = 0
        S_FanHs(i) = 0
        S_PerFanHs(i) = 0
        S_HeaterHs(i) = 0
        S_PerHeaterHs(i) = 0
        S_FanKWH(i) = 0
        S_HeaterKWH(i) = 0
    Next i
End If
' set the year index (yi) to 0
yi = 0

' compute number of days from January 1 of 1960 to the first year month and day of the weather file
TotalDaysToFile = NumberOfDays(InitialFileYear, InitialFileMonth, InitialFileDay)

' compute number of days from January 1 of 1960 to the first year month and day of the analysis
TotalDaysStartSim = NumberOfDays(InitialSimYear, InitialSimMonth, InitialSimDay)

' compute the number of days since the beginin of the weather data to the begining of the analysis
DaysToStart = TotalDaysStartSim - TotalDaysToFile

' coumpute the number of hours (lines) between the begining of the weather file and the first simulation hour
FirstLineRead = (DaysToStart * 24)

' compute number of days from January 1 of 1960 to the last year month and day of the analysis
TotalDaysFinishSim = NumberOfDays(FinalSimYear, FinalSimMonth, FinalSimDay)

' compute the position of the Last line to read in the file for the first year
DaysToFinish = (TotalDaysFinishSim - TotalDaysToFile) + 1

LastLineRead = (DaysToFinish * 24)

CurrentYear = InitialSimYear
il = 0
Open FileNameW For Input As #1

For CurrentYear = InitialSimYear To FMultYear 'this is the loop for the multiple simulation years

'these lines are for locating the first and last line to read in the file for the subsequent years
If CurrentYear = InitialSimYear Then
    FirstLineRead = FirstLineRead
    LastLineRead = LastLineRead
Else
    FirstLineRead = FirstLineRead + ((365 + Bisiesto1(CurrentYear)) * 24)
    LastLineRead = LastLineRead + ((365 + Bisiesto1(CurrentYear)) * 24)
End If
'these lines are to read all the previous lines in the file up to the first line used for the simulation
Do Until il = FirstLineRead
    Line Input #1, LineFromFile 'read this line from file
      
    il = il + 1
Loop 'loop to locate the first line to read

'check if only summary outouts were required
If S_OutputSummary = False Then
'open all the output files
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_mccenter" & ".txt" For Output As #3
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_mcside" & ".txt" For Output As #4
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_tempcenter" & ".txt" For Output As #6
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_tempside" & ".txt" For Output As #7
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_dmlcenter" & ".txt" For Output As #9
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_dmlside" & ".txt" For Output As #10
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_Fan" & ".txt" For Output As #11
End If
'initialize variables for the current year
FanRunHours = 0
HeaterRunHours = 0
Per_FanRun = 0
Per_HeaterRun = 0
FanKWH = 0
HeaterKWH = 0

'set the initial temperature and MC conditions for each grain layer
i = 0
For i = 0 To (NumberOfLayers - 1)
    GrainMC_WB_C(i) = GrainMC_In_WB_C(i)
    GrainMC_WB_S(i) = GrainMC_In_WB_S(i)
    GrainTemp_C_C(i) = GrainTemp_In_C_C(i)
    GrainTemp_C_S(i) = GrainTemp_In_C_S(i)
    LayerDML(i) = 0
    LayerDML_C(i) = 0
    LayerDML_S(i) = 0
Next i
'set the initial grain layer depth
i = 0
For i = 0 To (NumberOfLayers - 1)
    LayerDepth_C(i) = LayerDepth(i)
    LayerDepth_S(i) = LayerDepth(i)
Next i

'set counter to 0
Counter = 0
'set hous_stop to 0 for ending of the simulation based on final date
StopSim = False
Do While StopSim = False
    Line Input #1, LineFromFile 'read this line from file

'the string is converted into an array. Temperature corresponds to posicion 5, and RH
'to posicion 6 in the array

    
sItems() = Split(LineFromFile, vbTab)
    BadDataflag = 1
    FileTemp = sItems(TempColumn - 1)
    If FileTemp < -29.9 Or FileTemp > 40 Then BadDataflag = 0
    FileRH = sItems(RHColumn - 1)
    If FileRH = 100 Then FileRH = 99
    If FileRH < 1 Or FileRH > 100 Then BadDataflag = 0
    DateStamp = Format(sItems(1), "00") & "-" & Format(sItems(2), "00") & "-" & sItems(0)
    TimeStamp = Format(sItems(3), "00") & ":00"
    PrintTime = DateStamp & " " & TimeStamp
    
    If BadDataflag = 1 Then
        FileEMC_db = CF_EMC_D(FileTemp, FileRH / 100, ArrGrain, GrainIndex)
        FileEMC_wb = 100 * FileEMC_db / (100 + FileEMC_db)
        
        If FileRH >= MinRHSelect Then
            RHflag = 1
        Else
            RHflag = 0
        End If
        
        If FileEMC_wb >= MinEMCSelect Then
            EMCflag = 1
        Else
            EMCflag = 0
        End If
        
    End If
        HourFlag = RHflag * EMCflag * BadDataflag
    If HourFlag = 1 Then
        FanStatus = True
        PlenumTemp_C = FileTemp + FanPreWarming_C
        PsAmbient = Sat_press(Kelvin(FileTemp))
        PvAmbient = CompVaporPress(PsAmbient, FileRH)
        PvPlenum = PvAmbient
        PsPlenum = Sat_press(Kelvin(PlenumTemp_C))
        PlenumRHfan = PvPlenum / PsPlenum * 100
        PlenumRH = PlenumRHfan
        If FileRH > MaxRHSelect Then
            HeaterRHflag = 0
        Else
            HeaterRHflag = 1
        End If
        
        If FileEMC_wb > MaxEMCSelect Then
            HeaterEMCflag = 0
        Else
            HeaterEMCflag = 1
        End If
        
        HeaterFlag = HeaterRHflag * HeaterEMCflag
        
        If HeaterFlag = 0 Then
            HeaterStatus = True
        Else
            HeaterStatus = False
        End If
        If Tinc_C <= FanPreWarming_C Then
            HeaterStatus = False
        End If
        BurnerEnergy_KW = 0
        If HeaterStatus = True Then
            PlenumTemp_C = FileTemp + Tinc_C
            PsAmbient = Sat_press(Kelvin(FileTemp))
            PvAmbient = CompVaporPress(PsAmbient, FileRH)
            PvPlenum = PvAmbient
            PsPlenum = Sat_press(Kelvin(PlenumTemp_C))
            PlenumRH = PvPlenum / PsPlenum * 100
            ' compute the estimate power required
            HeatingEnergy_KW = CompBurnerPower(Tinc_C - FanPreWarming_C, FileTemp + FanPreWarming_C, PlenumRHfan, TotalAirflow)
            'compute the energy required by the burner to increase the temperature of the air
            BurnerEnergy_KW = HeatingEnergy_KW / (BurnerEfficiency / 100)
            BurnerEnergy_BTU = BurnerEnergy_KW * KWHtoBTU
        End If
       
        PlenumEMC_db = CF_EMC_D(PlenumTemp_C, PlenumRH / 100, ArrGrain, GrainIndex)
        PlenumEMC_wb = 100 * PlenumEMC_db / (100 + PlenumEMC_db)
        
    Else
        FanStatus = False
        HeaterStatus = False
    End If
    
    FanPrint = "Fan Off"
    HeaterPrint = "Heater Off"
    ' compute the fan runtime and percentaje of fan runtime
    If FanStatus = True Then
        FanRunHours = FanRunHours + 1
        FanPrint = "Fan On"
    End If
    Per_FanRun = Format(FanRunHours / (Counter + 1) * 100, "00.0")
    'compute the total energy consumtion of the fan (kwh) since the begining of the simulation for the current year
    FanKWH = FanPower_KW * FanRunHours
    ' compute the heater runtime and percentaje of heater runtime
    If HeaterStatus = True Then
        HeaterRunHours = HeaterRunHours + 1
        HeaterPrint = "Heater On"
    End If
    Per_HeaterRun = Format(HeaterRunHours / (Counter + 1) * 100, "00.0")
    'compute the total energy consumtion of the heater (kwh) since the begining of the simulation for the current year
    HeaterKWH = HeaterKWH + BurnerEnergy_KW
    
    If S_OutputSummary = False Then
        'print the heading for the fan output file
        PrintHeading_Fan = "Date  /   Time" & vbTab & vbTab & "A Temp" & vbTab & "A RH" & vbTab & "A EMC" & vbTab & "Pl Temp" & vbTab & "Pl RH" & vbTab & "Pl EMC" & vbTab & "Fan St" & vbTab & "Heater St" & vbTab & "F Run T" & vbTab & "F Run %" & vbTab & "H Run T" & vbTab & "H Run %" & vbTab & "F KWH" & vbTab & "H KWH"
        'set values into the string to be printed in the fan output file for the current hour
        PrintVar_Fan = Format(FileTemp, "#0.0") & vbTab & Format(FileRH, "#0") & vbTab & Format(FileEMC_wb, "#0.0") & vbTab & Format(PlenumTemp_C, "#0.0") & vbTab & Format(PlenumRH, "#0") & vbTab & Format(PlenumEMC_wb, "#0.0") & vbTab & FanPrint & vbTab & HeaterPrint & vbTab & Format(FanRunHours, "0") & vbTab & Format(Per_FanRun, "0.0") & vbTab & Format(HeaterRunHours, "0") & vbTab & Format(Per_HeaterRun, "0.0") & vbTab & Format(FanKWH, "0") & vbTab & Format(HeaterKWH, "0")
        FileNumber = 11
        Call PrintToFanFile(PrintVar_Fan, Counter, FileNumber, PrintTime, PrintHeading_Fan)
    End If
    
    If HourFlag = 1 Then    'if all the temp, RH and EMC data are satisfied for the current hour, then run the drying sub for each layer
        'compute drying for the center of the bin
        Call Drying(GrainMC_WB_C, GrainTemp_C_C, PlenumTemp_C, PlenumRH, LayerDepth_C, ArrGrain, GrainIndex, AirVel_C, FanStatus)
        'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the center of the bin
        i = 0
        For i = 0 To (NumberOfLayers - 1)
            GrainMC_WB_C(i) = MfW(i)
            GrainTemp_C_C(i) = GfC(i)
            LayerDepth_C(i) = dxf(i)
            LayerDML_C(i) = LayerDML(i)
        Next i
        
        If S_OutputSummary = False Then
            'write the headings for the output file
            i = 0
            textstrim = ""
            m = NumberOfLayers - 1
            For i = 0 To (m + 3)
                If i <= m Then
                    varprint = "Layer" & (i + 1)
                ElseIf i = (m + 1) Then
                    varprint = "Average"
                ElseIf i = (m + 2) Then
                    varprint = "Minimun"
                Else
                    varprint = "Maximum"
                End If
                textstrim = textstrim & varprint & vbTab
            Next i
            PrintHeading = "Date  /  Time" & vbTab & vbTab & textstrim

            'print in a file the moisture content of each layers at the center of the bin
            FileNumber = 3
            Call PrintToFile(GrainMC_WB_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            'print in a file the temperature of each layers at the center of the bin
            FileNumber = 6
            Call PrintToFile(GrainTemp_C_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            If ArrGrain(GrainIndex, 14) = "1" Then
            'print in a file the DML of each layers at the center of the bin
                FileNumber = 9
                Call PrintToFile(LayerDML_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            End If
        End If
        'compute drying for the side of the bin
        Call Drying(GrainMC_WB_S, GrainTemp_C_S, PlenumTemp_C, PlenumRH, LayerDepth_S, ArrGrain, GrainIndex, AirVel_S, FanStatus)
        'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the side of the bin
        i = 0
        For i = 0 To (NumberOfLayers - 1)
            GrainMC_WB_S(i) = MfW(i)
            GrainTemp_C_S(i) = GfC(i)
            LayerDepth_S(i) = dxf(i)
            LayerDML_S(i) = LayerDML(i)
        Next i
        
        If S_OutputSummary = False Then
            'print in a file the moisture content of each layers at the side of the bin
            FileNumber = 4
            Call PrintToFile(GrainMC_WB_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            'print in a file the temperature of each layers at the side of the bin
            FileNumber = 7
            Call PrintToFile(GrainTemp_C_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            If ArrGrain(GrainIndex, 14) = "1" Then
            'print in a file the DML of each layers at the side of the bin
                FileNumber = 10
                Call PrintToFile(LayerDML_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            End If
        End If
    End If ' for hour flag


' to make the simulation for the current year stop based on final date
If DateCriteria = True Then
    If il >= LastLineRead - 1 Then
        Hours_Stop = 0
        CurrentHours = 1
    Else
        Hours_Stop = 1
        CurrentHours = 0
    End If
End If

Call StopSimulation(MCCriteria, TempCriteria, DMLCriteria, GrainMC_WB_C, GrainMC_WB_S, GrainTemp_C_C, GrainTemp_C_S, LayerDML_C, LayerDML_S, CurrentHours, NumberOfLayers - 1, MCavg_Stop, MCmax_Stop, Tempavg_Stop, Tempmax_Stop, DMLavg_Stop, Hours_Stop)

Counter = Counter + 1
il = il + 1
Loop 'loop for the year simulation

If S_OutputSummary = False Then
    'close all open files
    Close #3 'MC center
    Close #4 'MC side
    Close #6 'temp center
    Close #7 'temp side
    Close #9 'dml center
    Close #10 'dml side
    Close #11 'fan file
End If

' create the arrays for the multiple year simulation
If MultYears = True Then
MaxLayer = NumberOfLayers - 1
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = GrainMC_WB_C(i)
    Else
        ArrCriteria(i) = GrainMC_WB_S(i - (MaxLayer + 1))
    End If
Next i
        S_Moisture_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Moisture_Min(yi) = ArrayMin(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Moisture_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = GrainTemp_C_C(i)
    Else
        ArrCriteria(i) = GrainTemp_C_S(i - (MaxLayer + 1))
    End If
Next i
        S_Temperature_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Temperature_Min(yi) = ArrayMin(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Temperature_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = LayerDML_C(i)
    Else
        ArrCriteria(i) = LayerDML_S(i - (MaxLayer + 1))
    End If
Next i
        S_DML_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_DML_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
        
        S_DryingHs(yi) = Counter
        S_FanHs(yi) = FanRunHours
        S_PerFanHs(yi) = Per_FanRun
        S_HeaterHs(yi) = HeaterRunHours
        S_PerHeaterHs(yi) = Per_HeaterRun
        S_FanKWH(yi) = FanKWH
        S_HeaterKWH(yi) = HeaterKWH
        
        DMLBin = S_DML_Avg(yi)
        AvgFinMC = S_Moisture_Avg(yi)
        S_TotDryingCost(yi) = DryingCost(FanKWH, HeaterKWH, BinCapacity_t, DMLBin, AvgFinMC)
End If

yi = yi + 1
Next CurrentYear 'loop for the multiple years simulation

If MultYears = True Then
Open CurrentDir & "\output files\" & BaseName & "_" & "_Summary.txt" For Output As #12
' print the summary of the simulation run settings in the end file
Print #12, StRunInfo
Print #12, ""

i = 0
For i = 0 To yi + 3
    If i < yi Then
        PrintHeading_End = "S Year" & vbTab & "$/ton" & vbTab & "c/bu" & vbTab & "MC Avg" & vbTab & "MC Min" & vbTab & "MC Max" & vbTab & "T Avg" & vbTab & "T Min" & vbTab & "T Max" & vbTab & "DML Avg" & vbTab & "DML Max" & vbTab & "Dr hs" & vbTab & "F hs" & vbTab & "F hs %" & vbTab & "H hs" & vbTab & "H hs %" & vbTab & "F KWH" & vbTab & "H KWH"
        PrintYear = (InitialSimYear + i)
        PrintVar_End = Format(S_TotDryingCost(i), "0.00") & vbTab & Format(S_TotDryingCost(i) / 0.4, "0.00") & vbTab & Format(S_Moisture_Avg(i), "0.00") & vbTab & Format(S_Moisture_Min(i), "0.00") & vbTab & Format(S_Moisture_Max(i), "0.00") & vbTab & Format(S_Temperature_Avg(i), "0.00") & vbTab & Format(S_Temperature_Min(i), "0.00") & vbTab & Format(S_Temperature_Max(i), "0.00") & vbTab & Format(S_DML_Avg(i), "0.00") & vbTab & Format(S_DML_Max(i), "0.00") & vbTab & Format(S_DryingHs(i), "0") & vbTab & Format(S_FanHs(i), "0") & vbTab & Format(S_PerFanHs(i), "0.0") & vbTab & Format(S_HeaterHs(i), "0") & vbTab & Format(S_PerHeaterHs(i), "0.0") & vbTab & Format(S_FanKWH(i), "0") & vbTab & Format(S_HeaterKWH(i), "0")
    ElseIf i = yi Then
        PrintYear = "Avg"
        PrintVar_End = Format(ArrayAvg(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayAvg(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayAvg(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayAvg(S_HeaterKWH, 0, yi - 1), "0")
    ElseIf i = yi + 1 Then
        PrintYear = "Min"
        PrintVar_End = Format(ArrayMin(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayMin(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMin(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMin(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayMin(S_HeaterKWH, 0, yi - 1), "0")
    ElseIf i = yi + 2 Then
        PrintYear = "Max"
        PrintVar_End = Format(ArrayMax(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayMax(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMax(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMax(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayMax(S_HeaterKWH, 0, yi - 1), "0")
        ElseIf i = yi + 2 Then
    ElseIf i = yi + 3 Then
        PrintYear = "StD"
        PrintVar_End = Format(ArrayStdDev(S_TotDryingCost, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_TotDryingCost, True, True, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Min, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Min, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DML_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DML_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DryingHs, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_FanHs, True, True, yi - 1), "0") _
                        & vbTab & Format(ArrayStdDev(S_PerFanHs, True, True, yi - 1), "0.0") & vbTab & Format(ArrayStdDev(S_HeaterHs, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_PerHeaterHs, True, True, yi - 1), "0.0") & vbTab & Format(ArrayStdDev(S_FanKWH, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_HeaterKWH, True, True, yi - 1), "0")
    End If
Call PrintToEndFile(PrintVar_End, i, 12, PrintYear, PrintHeading_End)
Next i
End If

Close #12
Close #1
End Sub

Sub VarHeat_Strat()

' set all the arrays for the set of years simulation to "0"
If MultYears = True Then
    i = 0
    For i = 0 To UBound(S_Moisture_Avg)
        S_Moisture_Avg(i) = 0
        S_Moisture_Min(i) = 0
        S_Moisture_Max(i) = 0
        S_Temperature_Avg(i) = 0
        S_Temperature_Min(i) = 0
        S_Temperature_Max(i) = 0
        S_DML_Avg(i) = 0
        S_DML_Max(i) = 0
        S_DryingHs(i) = 0
        S_FanHs(i) = 0
        S_PerFanHs(i) = 0
        S_HeaterHs(i) = 0
        S_PerHeaterHs(i) = 0
        S_FanKWH(i) = 0
        S_HeaterKWH(i) = 0
    Next i
End If
' set the year index (yi) to 0
yi = 0

' compute number of days from January 1 of 1960 to the first year month and day of the weather file
TotalDaysToFile = NumberOfDays(InitialFileYear, InitialFileMonth, InitialFileDay)

' compute number of days from January 1 of 1960 to the first year month and day of the analysis
TotalDaysStartSim = NumberOfDays(InitialSimYear, InitialSimMonth, InitialSimDay)

' compute the number of days since the beginin of the weather data to the begining of the analysis
DaysToStart = TotalDaysStartSim - TotalDaysToFile

' coumpute the number of hours (lines) between the begining of the weather file and the first simulation hour
FirstLineRead = (DaysToStart * 24)

' compute number of days from January 1 of 1960 to the last year month and day of the analysis
TotalDaysFinishSim = NumberOfDays(FinalSimYear, FinalSimMonth, FinalSimDay)

' compute the position of the Last line to read in the file for the first year
DaysToFinish = (TotalDaysFinishSim - TotalDaysToFile) + 1

LastLineRead = (DaysToFinish * 24)

CurrentYear = InitialSimYear
il = 0
Open FileNameW For Input As #1

For CurrentYear = InitialSimYear To FMultYear 'this is the loop for the multiple simulation years

'these lines are for locating the first and last line to read in the file for the subsequent years
If CurrentYear = InitialSimYear Then
    FirstLineRead = FirstLineRead
    LastLineRead = LastLineRead
Else
    FirstLineRead = FirstLineRead + ((365 + Bisiesto1(CurrentYear)) * 24)
    LastLineRead = LastLineRead + ((365 + Bisiesto1(CurrentYear)) * 24)
End If
'these lines are to read all the previous lines in the file up to the first line used for the simulation
Do Until il = FirstLineRead
    Line Input #1, LineFromFile 'read this line from file
      
    il = il + 1
Loop 'loop to locate the first line to read

'check if only summary outouts were required
If S_OutputSummary = False Then
'open all the output files
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_mccenter" & ".txt" For Output As #3
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_mcside" & ".txt" For Output As #4
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_tempcenter" & ".txt" For Output As #6
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_tempside" & ".txt" For Output As #7
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_dmlcenter" & ".txt" For Output As #9
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_dmlside" & ".txt" For Output As #10
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_Fan" & ".txt" For Output As #11
End If
'initialize variables for the current year
FanRunHours = 0
HeaterRunHours = 0
Per_FanRun = 0
Per_HeaterRun = 0
FanKWH = 0
HeaterKWH = 0

'set the initial temperature and MC conditions for each grain layer
i = 0
For i = 0 To (NumberOfLayers - 1)
    GrainMC_WB_C(i) = GrainMC_In_WB_C(i)
    GrainMC_WB_S(i) = GrainMC_In_WB_S(i)
    GrainTemp_C_C(i) = GrainTemp_In_C_C(i)
    GrainTemp_C_S(i) = GrainTemp_In_C_S(i)
    LayerDML(i) = 0
    LayerDML_C(i) = 0
    LayerDML_S(i) = 0
Next i
'set the initial grain layer depth
i = 0
For i = 0 To (NumberOfLayers - 1)
    LayerDepth_C(i) = LayerDepth(i)
    LayerDepth_S(i) = LayerDepth(i)
Next i

'set counter to 0
Counter = 0
'set hous_stop to 0 for ending of the simulation based on final date
StopSim = False
Do While StopSim = False
    Line Input #1, LineFromFile 'read this line from file

'the string is converted into an array. Temperature corresponds to posicion 5, and RH
'to posicion 6 in the array

    
sItems() = Split(LineFromFile, vbTab)
    BadDataflag = 1
    FileTemp = sItems(TempColumn - 1)
    If FileTemp < -29.9 Or FileTemp > 40 Then BadDataflag = 0
    FileRH = sItems(RHColumn - 1)
    If FileRH = 100 Then FileRH = 99
    If FileRH < 1 Or FileRH > 100 Then BadDataflag = 0
    DateStamp = Format(sItems(1), "00") & "-" & Format(sItems(2), "00") & "-" & sItems(0)
    TimeStamp = Format(sItems(3), "00") & ":00"
    PrintTime = DateStamp & " " & TimeStamp
    
    If BadDataflag = 1 Then
        FileEMC_db = CF_EMC_D(FileTemp, FileRH / 100, ArrGrain, GrainIndex)
        FileEMC_wb = 100 * FileEMC_db / (100 + FileEMC_db)
        
             
        If FileEMC_wb >= MinEMCSelect Then
            EMCflag = 1
        Else
            EMCflag = 0
        End If
        
    End If
        HourFlag = EMCflag * BadDataflag
    If HourFlag = 1 Then
        FanStatus = True
        PlenumTemp_C = FileTemp + FanPreWarming_C
        PsAmbient = Sat_press(Kelvin(FileTemp))
        PvAmbient = CompVaporPress(PsAmbient, FileRH)
        PvPlenum = PvAmbient
        PsPlenum = Sat_press(Kelvin(PlenumTemp_C))
        PlenumRHfan = PvPlenum / PsPlenum * 100
        PlenumRH = PlenumRHfan
               
        If FileEMC_wb > MaxEMCSelect Then
            HeaterEMCflag = 0
        Else
            HeaterEMCflag = 1
        End If
        
        HeaterFlag = HeaterEMCflag
        
        If HeaterFlag = 0 Then
            HeaterStatus = True
        Else
            HeaterStatus = False
        End If
        
        'compute the plenum EMC after the fan prewarming
        PlenumEMC_db = CF_EMC_D(PlenumTemp_C, PlenumRHfan / 100, ArrGrain, GrainIndex)
        PlenumEMC_wb = 100 * PlenumEMC_db / (100 + PlenumEMC_db)
        
        If PlenumEMC_wb < MaxEMCSelect Then
            HeaterStatus = False
        End If
        BurnerEnergy_KW = 0
        
        If HeaterStatus = True Then
        ' compute the temperature increase in the plenum to reduce the abient emc to the upper emc limit
        i = 0
        Do While PlenumEMC_wb > MaxEMCSelect
            PlenumTemp_C = FileTemp + i
            PsAmbient = Sat_press(Kelvin(FileTemp))
            PvAmbient = CompVaporPress(PsAmbient, FileRH)
            PvPlenum = PvAmbient
            PsPlenum = Sat_press(Kelvin(PlenumTemp_C))
            PlenumRH = PvPlenum / PsPlenum * 100
            PlenumEMC_db = CF_EMC_D(PlenumTemp_C, PlenumRH / 100, ArrGrain, GrainIndex)
            PlenumEMC_wb = 100 * PlenumEMC_db / (100 + PlenumEMC_db)
            i = i + 0.01
        Loop
        Tinc_C = i
        ' compute the estimate power required
            HeatingEnergy_KW = CompBurnerPower(Tinc_C - FanPreWarming_C, FileTemp + FanPreWarming_C, PlenumRHfan, TotalAirflow)
            'compute the energy required by the burner to increase the temperature of the air
            BurnerEnergy_KW = HeatingEnergy_KW / (BurnerEfficiency / 100)
            BurnerEnergy_BTU = BurnerEnergy_KW * KWHtoBTU
        End If
       
        PlenumEMC_db = CF_EMC_D(PlenumTemp_C, PlenumRH / 100, ArrGrain, GrainIndex)
        PlenumEMC_wb = 100 * PlenumEMC_db / (100 + PlenumEMC_db)
        
    Else
        FanStatus = False
        HeaterStatus = False
    End If
    
    FanPrint = "Fan Off"
    HeaterPrint = "Heater Off"
    ' compute the fan runtime and percentaje of fan runtime
    If FanStatus = True Then
        FanRunHours = FanRunHours + 1
        FanPrint = "Fan On"
    End If
    Per_FanRun = Format(FanRunHours / (Counter + 1) * 100, "00.0")
    'compute the total energy consumtion of the fan (kwh) since the begining of the simulation for the current year
    FanKWH = FanPower_KW * FanRunHours
    ' compute the heater runtime and percentaje of heater runtime
    If HeaterStatus = True Then
        HeaterRunHours = HeaterRunHours + 1
        HeaterPrint = "Heater On"
    End If
    Per_HeaterRun = Format(HeaterRunHours / (Counter + 1) * 100, "00.0")
    'compute the total energy consumtion of the heater (kwh) since the begining of the simulation for the current year
    HeaterKWH = HeaterKWH + BurnerEnergy_KW
    
    If S_OutputSummary = False Then
        'print the heading for the fan output file
        PrintHeading_Fan = "Date  /   Time" & vbTab & vbTab & "A Temp" & vbTab & "A RH" & vbTab & "A EMC" & vbTab & "Pl Temp" & vbTab & "Pl RH" & vbTab & "Pl EMC" & vbTab & "Fan St" & vbTab & "Heater St" & vbTab & "F Run T" & vbTab & "F Run %" & vbTab & "H Run T" & vbTab & "H Run %" & vbTab & "F KWH" & vbTab & "H KWH"
        'set values into the string to be printed in the fan output file for the current hour
        PrintVar_Fan = Format(FileTemp, "#0.0") & vbTab & Format(FileRH, "#0") & vbTab & Format(FileEMC_wb, "#0.0") & vbTab & Format(PlenumTemp_C, "#0.0") & vbTab & Format(PlenumRH, "#0") & vbTab & Format(PlenumEMC_wb, "#0.0") & vbTab & FanPrint & vbTab & HeaterPrint & vbTab & Format(FanRunHours, "0") & vbTab & Format(Per_FanRun, "0.0") & vbTab & Format(HeaterRunHours, "0") & vbTab & Format(Per_HeaterRun, "0.0") & vbTab & Format(FanKWH, "0") & vbTab & Format(HeaterKWH, "0")
        FileNumber = 11
        Call PrintToFanFile(PrintVar_Fan, Counter, FileNumber, PrintTime, PrintHeading_Fan)
    End If
    
    If HourFlag = 1 Then    'if all the temp, RH and EMC data are satisfied for the current hour, then run the drying sub for each layer
        'compute drying for the center of the bin
        Call Drying(GrainMC_WB_C, GrainTemp_C_C, PlenumTemp_C, PlenumRH, LayerDepth_C, ArrGrain, GrainIndex, AirVel_C, FanStatus)
        'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the center of the bin
        i = 0
        For i = 0 To (NumberOfLayers - 1)
            GrainMC_WB_C(i) = MfW(i)
            GrainTemp_C_C(i) = GfC(i)
            LayerDepth_C(i) = dxf(i)
            LayerDML_C(i) = LayerDML(i)
        Next i
        
        If S_OutputSummary = False Then
            'write the headings for the output file
            i = 0
            textstrim = ""
            m = NumberOfLayers - 1
            For i = 0 To (m + 3)
                If i <= m Then
                    varprint = "Layer" & (i + 1)
                ElseIf i = (m + 1) Then
                    varprint = "Average"
                ElseIf i = (m + 2) Then
                    varprint = "Minimun"
                Else
                    varprint = "Maximum"
                End If
                textstrim = textstrim & varprint & vbTab
            Next i
            PrintHeading = "Date  /  Time" & vbTab & vbTab & textstrim

            'print in a file the moisture content of each layers at the center of the bin
            FileNumber = 3
            Call PrintToFile(GrainMC_WB_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            'print in a file the temperature of each layers at the center of the bin
            FileNumber = 6
            Call PrintToFile(GrainTemp_C_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            If ArrGrain(GrainIndex, 14) = "1" Then
            'print in a file the DML of each layers at the center of the bin
                FileNumber = 9
                Call PrintToFile(LayerDML_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            End If
        End If
        'compute drying for the side of the bin
        Call Drying(GrainMC_WB_S, GrainTemp_C_S, PlenumTemp_C, PlenumRH, LayerDepth_S, ArrGrain, GrainIndex, AirVel_S, FanStatus)
        'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the side of the bin
        i = 0
        For i = 0 To (NumberOfLayers - 1)
            GrainMC_WB_S(i) = MfW(i)
            GrainTemp_C_S(i) = GfC(i)
            LayerDepth_S(i) = dxf(i)
            LayerDML_S(i) = LayerDML(i)
        Next i
        
        If S_OutputSummary = False Then
            'print in a file the moisture content of each layers at the side of the bin
            FileNumber = 4
            Call PrintToFile(GrainMC_WB_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            'print in a file the temperature of each layers at the side of the bin
            FileNumber = 7
            Call PrintToFile(GrainTemp_C_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            If ArrGrain(GrainIndex, 14) = "1" Then
            'print in a file the DML of each layers at the side of the bin
                FileNumber = 10
                Call PrintToFile(LayerDML_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            End If
        End If
    End If ' for hour flag


' to make the simulation for the current year stop based on final date
If DateCriteria = True Then
    If il >= LastLineRead - 1 Then
        Hours_Stop = 0
        CurrentHours = 1
    Else
        Hours_Stop = 1
        CurrentHours = 0
    End If
End If

Call StopSimulation(MCCriteria, TempCriteria, DMLCriteria, GrainMC_WB_C, GrainMC_WB_S, GrainTemp_C_C, GrainTemp_C_S, LayerDML_C, LayerDML_S, CurrentHours, NumberOfLayers - 1, MCavg_Stop, MCmax_Stop, Tempavg_Stop, Tempmax_Stop, DMLavg_Stop, Hours_Stop)

Counter = Counter + 1
il = il + 1
Loop 'loop for the year simulation

If S_OutputSummary = False Then
    'close all open files
    Close #3 'MC center
    Close #4 'MC side
    Close #6 'temp center
    Close #7 'temp side
    Close #9 'dml center
    Close #10 'dml side
    Close #11 'fan file
End If

' create the arrays for the multiple year simulation
If MultYears = True Then
MaxLayer = NumberOfLayers - 1
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = GrainMC_WB_C(i)
    Else
        ArrCriteria(i) = GrainMC_WB_S(i - (MaxLayer + 1))
    End If
Next i
        S_Moisture_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Moisture_Min(yi) = ArrayMin(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Moisture_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = GrainTemp_C_C(i)
    Else
        ArrCriteria(i) = GrainTemp_C_S(i - (MaxLayer + 1))
    End If
Next i
        S_Temperature_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Temperature_Min(yi) = ArrayMin(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Temperature_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = LayerDML_C(i)
    Else
        ArrCriteria(i) = LayerDML_S(i - (MaxLayer + 1))
    End If
Next i
        S_DML_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_DML_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
        
        S_DryingHs(yi) = Counter
        S_FanHs(yi) = FanRunHours
        S_PerFanHs(yi) = Per_FanRun
        S_HeaterHs(yi) = HeaterRunHours
        S_PerHeaterHs(yi) = Per_HeaterRun
        S_FanKWH(yi) = FanKWH
        S_HeaterKWH(yi) = HeaterKWH
        
        DMLBin = S_DML_Avg(yi)
        AvgFinMC = S_Moisture_Avg(yi)
        S_TotDryingCost(yi) = DryingCost(FanKWH, HeaterKWH, BinCapacity_t, DMLBin, AvgFinMC)
End If

yi = yi + 1
Next CurrentYear 'loop for the multiple years simulation

If MultYears = True Then
Open CurrentDir & "\output files\" & BaseName & "_" & "_Summary.txt" For Output As #12
' print the summary of the simulation run settings in the end file
Print #12, StRunInfo
Print #12, ""
i = 0
For i = 0 To yi + 3
    If i < yi Then
        PrintHeading_End = "S Year" & vbTab & "$/ton" & vbTab & "c/bu" & vbTab & "MC Avg" & vbTab & "MC Min" & vbTab & "MC Max" & vbTab & "T Avg" & vbTab & "T Min" & vbTab & "T Max" & vbTab & "DML Avg" & vbTab & "DML Max" & vbTab & "Dr hs" & vbTab & "F hs" & vbTab & "F hs %" & vbTab & "H hs" & vbTab & "H hs %" & vbTab & "F KWH" & vbTab & "H KWH"
        PrintYear = (InitialSimYear + i)
        PrintVar_End = Format(S_TotDryingCost(i), "0.00") & vbTab & Format(S_TotDryingCost(i) / 0.4, "0.00") & vbTab & Format(S_Moisture_Avg(i), "0.00") & vbTab & Format(S_Moisture_Min(i), "0.00") & vbTab & Format(S_Moisture_Max(i), "0.00") & vbTab & Format(S_Temperature_Avg(i), "0.00") & vbTab & Format(S_Temperature_Min(i), "0.00") & vbTab & Format(S_Temperature_Max(i), "0.00") & vbTab & Format(S_DML_Avg(i), "0.00") & vbTab & Format(S_DML_Max(i), "0.00") & vbTab & Format(S_DryingHs(i), "0") & vbTab & Format(S_FanHs(i), "0") & vbTab & Format(S_PerFanHs(i), "0.0") & vbTab & Format(S_HeaterHs(i), "0") & vbTab & Format(S_PerHeaterHs(i), "0.0") & vbTab & Format(S_FanKWH(i), "0") & vbTab & Format(S_HeaterKWH(i), "0")
    ElseIf i = yi Then
        PrintYear = "Avg"
        PrintVar_End = Format(ArrayAvg(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayAvg(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayAvg(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayAvg(S_HeaterKWH, 0, yi - 1), "0")
    ElseIf i = yi + 1 Then
        PrintYear = "Min"
        PrintVar_End = Format(ArrayMin(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayMin(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMin(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMin(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayMin(S_HeaterKWH, 0, yi - 1), "0")
    ElseIf i = yi + 2 Then
        PrintYear = "Max"
        PrintVar_End = Format(ArrayMax(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayMax(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMax(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMax(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayMax(S_HeaterKWH, 0, yi - 1), "0")
        ElseIf i = yi + 2 Then
    ElseIf i = yi + 3 Then
        PrintYear = "StD"
        PrintVar_End = Format(ArrayStdDev(S_TotDryingCost, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_TotDryingCost, True, True, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Min, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Min, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DML_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DML_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DryingHs, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_FanHs, True, True, yi - 1), "0") _
                        & vbTab & Format(ArrayStdDev(S_PerFanHs, True, True, yi - 1), "0.0") & vbTab & Format(ArrayStdDev(S_HeaterHs, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_PerHeaterHs, True, True, yi - 1), "0.0") & vbTab & Format(ArrayStdDev(S_FanKWH, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_HeaterKWH, True, True, yi - 1), "0")
    End If
Call PrintToEndFile(PrintVar_End, i, 12, PrintYear, PrintHeading_End)
Next i
End If

Close #12
Close #1
End Sub


Sub SAVH_Strat()

' set all the arrays for the set of years simulation to "0"
If MultYears = True Then
    i = 0
    For i = 0 To UBound(S_Moisture_Avg)
        S_Moisture_Avg(i) = 0
        S_Moisture_Min(i) = 0
        S_Moisture_Max(i) = 0
        S_Temperature_Avg(i) = 0
        S_Temperature_Min(i) = 0
        S_Temperature_Max(i) = 0
        S_DML_Avg(i) = 0
        S_DML_Max(i) = 0
        S_DryingHs(i) = 0
        S_FanHs(i) = 0
        S_PerFanHs(i) = 0
        S_HeaterHs(i) = 0
        S_PerHeaterHs(i) = 0
        S_FanKWH(i) = 0
        S_HeaterKWH(i) = 0
    Next i
End If
' set the year index (yi) to 0
yi = 0
'estimate the number of drying hours according to the average airflow
EstDryingTime = 600 / AirflowRate_cfm
' compute number of days from January 1 of 1960 to the first year month and day of the weather file
TotalDaysToFile = NumberOfDays(InitialFileYear, InitialFileMonth, InitialFileDay)

' compute number of days from January 1 of 1960 to the first year month and day of the analysis
TotalDaysStartSim = NumberOfDays(InitialSimYear, InitialSimMonth, InitialSimDay)

' compute the number of days since the beginin of the weather data to the begining of the analysis
DaysToStart = TotalDaysStartSim - TotalDaysToFile

' coumpute the number of hours (lines) between the begining of the weather file and the first simulation hour
FirstLineRead = (DaysToStart * 24)

' compute number of days from January 1 of 1960 to the last year month and day of the analysis
TotalDaysFinishSim = NumberOfDays(FinalSimYear, FinalSimMonth, FinalSimDay)

' compute the position of the Last line to read in the file for the first year
DaysToFinish = (TotalDaysFinishSim - TotalDaysToFile) + 1

LastLineRead = (DaysToFinish * 24)

CurrentYear = InitialSimYear
il = 0
Open FileNameW For Input As #1

For CurrentYear = InitialSimYear To FMultYear 'this is the loop for the multiple simulation years

'these lines are for locating the first and last line to read in the file for the subsequent years
If CurrentYear = InitialSimYear Then
    FirstLineRead = FirstLineRead
    LastLineRead = LastLineRead
Else
    FirstLineRead = FirstLineRead + ((365 + Bisiesto1(CurrentYear)) * 24)
    LastLineRead = LastLineRead + ((365 + Bisiesto1(CurrentYear)) * 24)
End If
'these lines are to read all the previous lines in the file up to the first line used for the simulation
Do Until il = FirstLineRead
    Line Input #1, LineFromFile 'read this line from file
      
    il = il + 1
Loop 'loop to locate the first line to read

'check if only summary outouts were required
If S_OutputSummary = False Then
'open all the output files
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_mccenter" & ".txt" For Output As #3
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_mcside" & ".txt" For Output As #4
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_tempcenter" & ".txt" For Output As #6
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_tempside" & ".txt" For Output As #7
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_dmlcenter" & ".txt" For Output As #9
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_dmlside" & ".txt" For Output As #10
Open CurrentDir & "\output files\" & BaseName & "_" & CurrentYear & "_Fan" & ".txt" For Output As #11
End If
'initialize variables for the current year
FanRunHours = 0
HeaterRunHours = 0
Per_FanRun = 0
Per_HeaterRun = 0
FanKWH = 0
HeaterKWH = 0

'set the initial temperature and MC conditions for each grain layer
i = 0
For i = 0 To (NumberOfLayers - 1)
    GrainMC_WB_C(i) = GrainMC_In_WB_C(i)
    GrainMC_WB_S(i) = GrainMC_In_WB_S(i)
    GrainTemp_C_C(i) = GrainTemp_In_C_C(i)
    GrainTemp_C_S(i) = GrainTemp_In_C_S(i)
    LayerDML(i) = 0
    LayerDML_C(i) = 0
    LayerDML_S(i) = 0
Next i
'set the initial grain layer depth
i = 0
For i = 0 To (NumberOfLayers - 1)
    LayerDepth_C(i) = LayerDepth(i)
    LayerDepth_S(i) = LayerDepth(i)
Next i

'set counter to 0
Counter = 0
'set hous_stop to 0 for ending of the simulation based on final date
StopSim = False
Do While StopSim = False
    Line Input #1, LineFromFile 'read this line from file

'the string is converted into an array. Temperature corresponds to posicion 5, and RH
'to posicion 6 in the array

    
sItems() = Split(LineFromFile, vbTab)
    BadDataflag = 1
    FileTemp = sItems(TempColumn - 1)
    If FileTemp < -29.9 Or FileTemp > 40 Then BadDataflag = 0
    FileRH = sItems(RHColumn - 1)
    If FileRH = 100 Then FileRH = 99
    If FileRH < 1 Or FileRH > 100 Then BadDataflag = 0
    DateStamp = Format(sItems(1), "00") & "-" & Format(sItems(2), "00") & "-" & sItems(0)
    TimeStamp = Format(sItems(3), "00") & ":00"
    PrintTime = DateStamp & " " & TimeStamp
    
    If BadDataflag = 1 Then
        FileEMC_db = CF_EMC_D(FileTemp, FileRH / 100, ArrGrain, GrainIndex)
        FileEMC_wb = 100 * FileEMC_db / (100 + FileEMC_db)
        'find the lowest and highest mc in the first grain layer
        If GrainMC_WB_C(0) >= GrainMC_WB_S(0) Then
            MC_Highest = GrainMC_WB_C(0)
            MC_Lowest = GrainMC_WB_S(0)
        Else
            MC_Highest = GrainMC_WB_S(0)
            MC_Lowest = GrainMC_WB_C(0)
        End If
        'set the lower and upper mc limits
        MC_HighLimit = SAVH_FinalMC
        If FanRunHours / EstDryingTime < 0.5 Then
            MC_LowLimit = SAVH_FinalMC - 3
        ElseIf FanRunHours / EstDryingTime < 0.65 Then
            MC_LowLimit = SAVH_FinalMC - 1.5
        ElseIf FanRunHours / EstDryingTime < 0.8 Then
            MC_LowLimit = SAVH_FinalMC - 1
        ElseIf FanRunHours / EstDryingTime < 0.9 Then
            MC_LowLimit = SAVH_FinalMC - 0.75
        Else
            MC_LowLimit = SAVH_FinalMC - 0.5
        End If
        
    
    End If
    
    HourFlag = BadDataflag
    
    If HourFlag = 1 Then
        PlenumTemp_C = FileTemp + FanPreWarming_C
        PsAmbient = Sat_press(Kelvin(FileTemp))
        PvAmbient = CompVaporPress(PsAmbient, FileRH)
        PvPlenum = PvAmbient
        PsPlenum = Sat_press(Kelvin(PlenumTemp_C))
        PlenumRHfan = PvPlenum / PsPlenum * 100
        PlenumRH = PlenumRHfan
        'compute the plenum EMC after the fan prewarming
        PlenumEMC_db = CF_EMC_D(PlenumTemp_C, PlenumRHfan / 100, ArrGrain, GrainIndex)
        PlenumEMC_wb = 100 * PlenumEMC_db / (100 + PlenumEMC_db)
               
    'determine the fan and heater status
    If MC_Lowest < MC_LowLimit Then
        If FileEMC_wb <= MC_Lowest Then
            FanStatus = False
            HeaterStatus = False
        Else
            FanStatus = True
            HeaterStatus = False
        End If
    Else
        FanStatus = True
        If MC_Highest > MC_HighLimit Then
            If FileEMC_wb > MC_HighLimit Then
                HeaterStatus = True
            Else
                HeaterStatus = False
            End If
        End If
    End If
        
        
       BurnerEnergy_KW = 0
        
        If HeaterStatus = True Then
        ' compute the temperature increase in the plenum to reduce the ambient emc to the upper emc limit
        i = 0
        Do While PlenumEMC_wb > MC_HighLimit
            PlenumTemp_C = FileTemp + i
            PsAmbient = Sat_press(Kelvin(FileTemp))
            PvAmbient = CompVaporPress(PsAmbient, FileRH)
            PvPlenum = PvAmbient
            PsPlenum = Sat_press(Kelvin(PlenumTemp_C))
            PlenumRH = PvPlenum / PsPlenum * 100
            PlenumEMC_db = CF_EMC_D(PlenumTemp_C, PlenumRH / 100, ArrGrain, GrainIndex)
            PlenumEMC_wb = 100 * PlenumEMC_db / (100 + PlenumEMC_db)
            i = i + 0.01
        Loop
        Tinc_C = i
        ' compute the estimate power required
            HeatingEnergy_KW = CompBurnerPower(Tinc_C - FanPreWarming_C, FileTemp + FanPreWarming_C, PlenumRHfan, TotalAirflow)
            'compute the energy required by the burner to increase the temperature of the air
            BurnerEnergy_KW = HeatingEnergy_KW / (BurnerEfficiency / 100)
            BurnerEnergy_BTU = BurnerEnergy_KW * KWHtoBTU
        End If
       
        PlenumEMC_db = CF_EMC_D(PlenumTemp_C, PlenumRH / 100, ArrGrain, GrainIndex)
        PlenumEMC_wb = 100 * PlenumEMC_db / (100 + PlenumEMC_db)
        
    Else
        FanStatus = False
        HeaterStatus = False
    End If
    
    FanPrint = "Fan Off"
    HeaterPrint = "Heater Off"
    ' compute the fan runtime and percentaje of fan runtime
    If FanStatus = True Then
        FanRunHours = FanRunHours + 1
        FanPrint = "Fan On"
    End If
    Per_FanRun = Format(FanRunHours / (Counter + 1) * 100, "00.0")
    'compute the total energy consumtion of the fan (kwh) since the begining of the simulation for the current year
    FanKWH = FanPower_KW * FanRunHours
    ' compute the heater runtime and percentaje of heater runtime
    If HeaterStatus = True Then
        HeaterRunHours = HeaterRunHours + 1
        HeaterPrint = "Heater On"
    End If
    Per_HeaterRun = Format(HeaterRunHours / (Counter + 1) * 100, "00.0")
    'compute the total energy consumtion of the heater (kwh) since the begining of the simulation for the current year
    HeaterKWH = HeaterKWH + BurnerEnergy_KW
    
    If S_OutputSummary = False Then
        'print the heading for the fan output file
        PrintHeading_Fan = "Date  /   Time" & vbTab & vbTab & "A Temp" & vbTab & "A RH" & vbTab & "A EMC" & vbTab & "Pl Temp" & vbTab & "Pl RH" & vbTab & "Pl EMC" & vbTab & "Fan St" & vbTab & "Heater St" & vbTab & "F Run T" & vbTab & "F Run %" & vbTab & "H Run T" & vbTab & "H Run %" & vbTab & "F KWH" & vbTab & "H KWH"
        'set values into the string to be printed in the fan output file for the current hour
        PrintVar_Fan = Format(FileTemp, "#0.0") & vbTab & Format(FileRH, "#0") & vbTab & Format(FileEMC_wb, "#0.0") & vbTab & Format(PlenumTemp_C, "#0.0") & vbTab & Format(PlenumRH, "#0") & vbTab & Format(PlenumEMC_wb, "#0.0") & vbTab & FanPrint & vbTab & HeaterPrint & vbTab & Format(FanRunHours, "0") & vbTab & Format(Per_FanRun, "0.0") & vbTab & Format(HeaterRunHours, "0") & vbTab & Format(Per_HeaterRun, "0.0") & vbTab & Format(FanKWH, "0") & vbTab & Format(HeaterKWH, "0")
        FileNumber = 11
        Call PrintToFanFile(PrintVar_Fan, Counter, FileNumber, PrintTime, PrintHeading_Fan)
    End If
    
    If HourFlag = 1 Then    'if all the temp, RH and EMC data are satisfied for the current hour, then run the drying sub for each layer
        'compute drying for the center of the bin
        Call Drying(GrainMC_WB_C, GrainTemp_C_C, PlenumTemp_C, PlenumRH, LayerDepth_C, ArrGrain, GrainIndex, AirVel_C, FanStatus)
        'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the center of the bin
        i = 0
        For i = 0 To (NumberOfLayers - 1)
            GrainMC_WB_C(i) = MfW(i)
            GrainTemp_C_C(i) = GfC(i)
            LayerDepth_C(i) = dxf(i)
            LayerDML_C(i) = LayerDML(i)
        Next i
        
        If S_OutputSummary = False Then
            'write the headings for the output file
            i = 0
            textstrim = ""
            m = NumberOfLayers - 1
            For i = 0 To (m + 3)
                If i <= m Then
                    varprint = "Layer" & (i + 1)
                ElseIf i = (m + 1) Then
                    varprint = "Average"
                ElseIf i = (m + 2) Then
                    varprint = "Minimun"
                Else
                    varprint = "Maximum"
                End If
                textstrim = textstrim & varprint & vbTab
            Next i
            PrintHeading = "Date  /  Time" & vbTab & vbTab & textstrim

            'print in a file the moisture content of each layers at the center of the bin
            FileNumber = 3
            Call PrintToFile(GrainMC_WB_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            'print in a file the temperature of each layers at the center of the bin
            FileNumber = 6
            Call PrintToFile(GrainTemp_C_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            If ArrGrain(GrainIndex, 14) = "1" Then
            'print in a file the DML of each layers at the center of the bin
                FileNumber = 9
                Call PrintToFile(LayerDML_C, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            End If
        End If
        'compute drying for the side of the bin
        Call Drying(GrainMC_WB_S, GrainTemp_C_S, PlenumTemp_C, PlenumRH, LayerDepth_S, ArrGrain, GrainIndex, AirVel_S, FanStatus)
        'set the grain temperature and MC and layer depth values at the end of the time step to the temp and MC variables for the side of the bin
        i = 0
        For i = 0 To (NumberOfLayers - 1)
            GrainMC_WB_S(i) = MfW(i)
            GrainTemp_C_S(i) = GfC(i)
            LayerDepth_S(i) = dxf(i)
            LayerDML_S(i) = LayerDML(i)
        Next i
        
        If S_OutputSummary = False Then
            'print in a file the moisture content of each layers at the side of the bin
            FileNumber = 4
            Call PrintToFile(GrainMC_WB_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            'print in a file the temperature of each layers at the side of the bin
            FileNumber = 7
            Call PrintToFile(GrainTemp_C_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            If ArrGrain(GrainIndex, 14) = "1" Then
            'print in a file the DML of each layers at the side of the bin
                FileNumber = 10
                Call PrintToFile(LayerDML_S, (NumberOfLayers - 1), Counter, FileNumber, PrintTime, PrintHeading)
            End If
        End If
    End If ' for hour flag


' to make the simulation for the current year stop based on final date
If DateCriteria = True Then
    If il >= LastLineRead - 1 Then
        Hours_Stop = 0
        CurrentHours = 1
    Else
        Hours_Stop = 1
        CurrentHours = 0
    End If
End If

Call StopSimulation(MCCriteria, TempCriteria, DMLCriteria, GrainMC_WB_C, GrainMC_WB_S, GrainTemp_C_C, GrainTemp_C_S, LayerDML_C, LayerDML_S, CurrentHours, NumberOfLayers - 1, MCavg_Stop, MCmax_Stop, Tempavg_Stop, Tempmax_Stop, DMLavg_Stop, Hours_Stop)

Counter = Counter + 1
il = il + 1
Loop 'loop for the year simulation

If S_OutputSummary = False Then
    'close all open files
    Close #3 'MC center
    Close #4 'MC side
    Close #6 'temp center
    Close #7 'temp side
    Close #9 'dml center
    Close #10 'dml side
    Close #11 'fan file
End If

' create the arrays for the multiple year simulation
If MultYears = True Then
MaxLayer = NumberOfLayers - 1
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = GrainMC_WB_C(i)
    Else
        ArrCriteria(i) = GrainMC_WB_S(i - (MaxLayer + 1))
    End If
Next i
        S_Moisture_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Moisture_Min(yi) = ArrayMin(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Moisture_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = GrainTemp_C_C(i)
    Else
        ArrCriteria(i) = GrainTemp_C_S(i - (MaxLayer + 1))
    End If
Next i
        S_Temperature_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Temperature_Min(yi) = ArrayMin(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_Temperature_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = LayerDML_C(i)
    Else
        ArrCriteria(i) = LayerDML_S(i - (MaxLayer + 1))
    End If
Next i
        S_DML_Avg(yi) = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2 + 1))
        S_DML_Max(yi) = ArrayMax(ArrCriteria, 0, (MaxLayer * 2 + 1))
        
        S_DryingHs(yi) = Counter
        S_FanHs(yi) = FanRunHours
        S_PerFanHs(yi) = Per_FanRun
        S_HeaterHs(yi) = HeaterRunHours
        S_PerHeaterHs(yi) = Per_HeaterRun
        S_FanKWH(yi) = FanKWH
        S_HeaterKWH(yi) = HeaterKWH
        
        DMLBin = S_DML_Avg(yi)
        AvgFinMC = S_Moisture_Avg(yi)
        S_TotDryingCost(yi) = DryingCost(FanKWH, HeaterKWH, BinCapacity_t, DMLBin, AvgFinMC)
End If

yi = yi + 1
Next CurrentYear 'loop for the multiple years simulation

If MultYears = True Then
Open CurrentDir & "\output files\" & BaseName & "_" & "_Summary.txt" For Output As #12
' print the summary of the simulation run settings in the end file
Print #12, StRunInfo
Print #12, ""
i = 0
For i = 0 To yi + 3
    If i < yi Then
        PrintHeading_End = "S Year" & vbTab & "$/ton" & vbTab & "c/bu" & vbTab & "MC Avg" & vbTab & "MC Min" & vbTab & "MC Max" & vbTab & "T Avg" & vbTab & "T Min" & vbTab & "T Max" & vbTab & "DML Avg" & vbTab & "DML Max" & vbTab & "Dr hs" & vbTab & "F hs" & vbTab & "F hs %" & vbTab & "H hs" & vbTab & "H hs %" & vbTab & "F KWH" & vbTab & "H KWH"
        PrintYear = (InitialSimYear + i)
        PrintVar_End = Format(S_TotDryingCost(i), "0.00") & vbTab & Format(S_TotDryingCost(i) / 0.4, "0.00") & vbTab & Format(S_Moisture_Avg(i), "0.00") & vbTab & Format(S_Moisture_Min(i), "0.00") & vbTab & Format(S_Moisture_Max(i), "0.00") & vbTab & Format(S_Temperature_Avg(i), "0.00") & vbTab & Format(S_Temperature_Min(i), "0.00") & vbTab & Format(S_Temperature_Max(i), "0.00") & vbTab & Format(S_DML_Avg(i), "0.00") & vbTab & Format(S_DML_Max(i), "0.00") & vbTab & Format(S_DryingHs(i), "0") & vbTab & Format(S_FanHs(i), "0") & vbTab & Format(S_PerFanHs(i), "0.0") & vbTab & Format(S_HeaterHs(i), "0") & vbTab & Format(S_PerHeaterHs(i), "0.0") & vbTab & Format(S_FanKWH(i), "0") & vbTab & Format(S_HeaterKWH(i), "0")
    ElseIf i = yi Then
        PrintYear = "Avg"
        PrintVar_End = Format(ArrayAvg(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayAvg(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayAvg(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayAvg(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayAvg(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayAvg(S_HeaterKWH, 0, yi - 1), "0")
    ElseIf i = yi + 1 Then
        PrintYear = "Min"
        PrintVar_End = Format(ArrayMin(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayMin(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMin(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMin(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayMin(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMin(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayMin(S_HeaterKWH, 0, yi - 1), "0")
    ElseIf i = yi + 2 Then
        PrintYear = "Max"
        PrintVar_End = Format(ArrayMax(S_TotDryingCost, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_TotDryingCost, 0, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayMax(S_Moisture_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Moisture_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Moisture_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Min, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_Temperature_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DML_Avg, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DML_Max, 0, yi - 1), "0.00") & vbTab & Format(ArrayMax(S_DryingHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_FanHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_PerFanHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMax(S_HeaterHs, 0, yi - 1), "0") & vbTab & Format(ArrayMax(S_PerHeaterHs, 0, yi - 1), "0.0") & vbTab & Format(ArrayMax(S_FanKWH, 0, yi - 1), "0") _
        & vbTab & Format(ArrayMax(S_HeaterKWH, 0, yi - 1), "0")
        ElseIf i = yi + 2 Then
    ElseIf i = yi + 3 Then
        PrintYear = "StD"
        PrintVar_End = Format(ArrayStdDev(S_TotDryingCost, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_TotDryingCost, True, True, yi - 1) / 0.4, "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Min, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Moisture_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Min, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_Temperature_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DML_Avg, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DML_Max, True, True, yi - 1), "0.00") & vbTab & Format(ArrayStdDev(S_DryingHs, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_FanHs, True, True, yi - 1), "0") _
                        & vbTab & Format(ArrayStdDev(S_PerFanHs, True, True, yi - 1), "0.0") & vbTab & Format(ArrayStdDev(S_HeaterHs, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_PerHeaterHs, True, True, yi - 1), "0.0") & vbTab & Format(ArrayStdDev(S_FanKWH, True, True, yi - 1), "0") & vbTab & Format(ArrayStdDev(S_HeaterKWH, True, True, yi - 1), "0")
    End If
Call PrintToEndFile(PrintVar_End, i, 12, PrintYear, PrintHeading_End)
Next i
End If

Close #12
Close #1
End Sub



' The average of an array of any type
'
' FIRST and LAST indicate which portion of the array
' should be considered; they default to the first
' and last element, respectively
' if IGNOREEMPTY argument is True or omitted,
' Empty values aren't accounted for

Function ArrayAvg(arr As Variant, Optional First As Variant, _
    Optional Last As Variant, Optional IgnoreEmpty As Boolean = True) As Variant
    Dim Index As Long
    Dim sum As Variant
    Dim count As Long

    If IsMissing(First) Then First = LBound(arr)
    If IsMissing(Last) Then Last = UBound(arr)
    
    ' if arr isn't an array, the following statement raises an error
    For Index = First To Last
        If IgnoreEmpty = False Or Not IsEmpty(arr(Index)) Then
            sum = sum + arr(Index)
            count = count + 1
        End If
    Next
    
    ' return the average
    ArrayAvg = sum / count

End Function

' The standard deviation of an array of any type
'
' if the second argument is True or omitted,
' it evaluates the standard deviation of a sample,
' if it is False it evaluates the standard deviation of a population
'
' if the third argument is True or omitted, Empty values aren't accounted for


Function ArrayStdDev(arr As Variant, Optional SampleStdDev As Boolean = True, _
    Optional IgnoreEmpty As Boolean = True, Optional ArrUpLimit As Integer) As Double
    Dim sum As Double
    Dim sumSquare As Double
    Dim value As Double
    Dim count As Long
    Dim Index As Long
    

    ' evaluate sum of values
    ' if arr isn't an array, the following statement raises an error
    For Index = LBound(arr) To ArrUpLimit
        value = arr(Index)
        ' skip over non-numeric values
        If IsNumeric(value) Then
            ' skip over empty values, if requested
            If Not (IgnoreEmpty And IsEmpty(value)) Then
                ' add to the running total
                count = count + 1
                sum = sum + value
                sumSquare = sumSquare + value * value
            End If
         End If
    Next
    
    If count < 3 Then
        ArrayStdDev = 0
    Else
        ' evaluate the result
        ' use (Count-1) if evaluating the standard deviation of a sample
        If SampleStdDev Then
            ArrayStdDev = Sqr((sumSquare - (sum * sum / count)) / (count - 1))
        Else
            ArrayStdDev = Sqr((sumSquare - (sum * sum / count)) / count)
        End If
    End If

End Function

' Return the minimum value in an array of any type
'
' FIRST and LAST indicate which portion of the array
' should be considered; they default to the first
' and last element, respectively
' If MININDEX is passed, it receives the index of the
' minimum element in the array

Function ArrayMin(arr As Variant, Optional ByVal First As Variant, _
    Optional ByVal Last As Variant, Optional MinIndex As Long) As Variant
    Dim Index As Long
    
    If IsMissing(First) Then First = LBound(arr)
    If IsMissing(Last) Then Last = UBound(arr)

    MinIndex = First
    ArrayMin = arr(MinIndex)

    For Index = First + 1 To Last
        If ArrayMin > arr(Index) Then
            MinIndex = Index
            ArrayMin = arr(MinIndex)
        End If
    Next
End Function

' Return the maximum value in an array of any type
'
' FIRST and LAST indicate which portion of the array
' should be considered; they default to the first
' and last element, respectively
' If MAXINDEX is passed, it receives the index of the
' maximum element in the array

Function ArrayMax(arr As Variant, Optional ByVal First As Variant, _
    Optional ByVal Last As Variant, Optional MaxIndex As Long) As Variant
    Dim Index As Long
    
    If IsMissing(First) Then First = LBound(arr)
    If IsMissing(Last) Then Last = UBound(arr)

    MaxIndex = First
    ArrayMax = arr(MaxIndex)

    For Index = First + 1 To Last
        If ArrayMax < arr(Index) Then
            MaxIndex = Index
            ArrayMax = arr(MaxIndex)
        End If
    Next
End Function

Sub StopSimulation(MCCriteria, TempCriteria, DMLCriteria, GrainMC_WB_C, GrainMC_WB_S, GrainTemp_C_C, GrainTemp_C_S, LayerDML_C, LayerDML_S, CurrentHours, MaxLayer, MCavg_Stop, MCmax_Stop, Tempavg_Stop, Tempmax_Stop, DMLavg_Stop, Hours_Stop)
'this sub check for moisture content, temperature, time and DML criteria to decide if the
'simulation was completed or not
If MCCriteria = True Then
    i = 0
    For i = 0 To MaxLayer
        ArrCriteriaC(i) = GrainMC_WB_C(i)
        ArrCriteriaS(i) = GrainMC_WB_S(i)
    Next i
End If
If TempCriteria = True Then
    i = 0
    For i = 0 To MaxLayer
        ArrCriteriaC(i) = GrainTemp_C_C(i)
        ArrCriteriaS(i) = GrainTemp_C_S(i)
    Next i
End If
If DMLCriteria = True Then
    i = 0
    For i = 0 To MaxLayer
        ArrCriteriaC(i) = LayerDML_C(i)
        ArrCriteriaS(i) = LayerDML_S(i)
    Next i
End If

i = 0
For i = 0 To ((MaxLayer * 2) + 1)
    If i <= MaxLayer Then
        ArrCriteria(i) = ArrCriteriaC(i)
    Else
        ArrCriteria(i) = ArrCriteriaS(i - (MaxLayer + 1))
    End If
Next i

If MCCriteria = True Then
    avgvalue = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2) + 1)
    maxvalue = ArrayMax(ArrCriteria, 0, (MaxLayer * 2) + 1)
    If avgvalue <= MCavg_Stop And maxvalue <= MCmax_Stop Then
        StopSim = True
    Else
        StopSim = False
    End If
ElseIf TempCriteria = True Then
    avgvalue = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2) + 1)
    maxvalue = ArrayMax(ArrCriteria, 0, (MaxLayer * 2) + 1)
    If avgvalue <= Tempavg_Stop And maxvalue <= Tempmax_Stop Then
        StopSim = True
    Else
        StopSim = False
    End If
ElseIf DMLCriteria = True Then
    avgvalue = ArrayAvg(ArrCriteria, 0, (MaxLayer * 2) + 1)
    If avgvalue >= DMLavg_Stop Then
        StopSim = True
    Else
        StopSim = False
    End If
Else
    If CurrentHours >= Hours_Stop Then
        StopSim = True
    Else
        StopSim = False
    End If
End If
    
    
End Sub



'declare procedure
' To determine if a year is "bisiesto"
' Year is the year that questioned
' If Year is Bisiesto, then Bisiesto1 = 1
Public Function Bisiesto1(Year)
Dim InternalRef As Integer
    If (Int(Year / 4) = (Year / 4)) Then
        InternalRef = 1
    Else: InternalRef = 0
    End If
Bisiesto1 = InternalRef
End Function


' compute the number of days from January 1 1960 to January 1 of the first year in file
Public Function DaysToFirstYear(StartYear)
Dim NumberDays As Integer
Dim NumberDays1 As Integer
Dim YearIndex As Integer
YearIndex = 1960
NumberDays = 0
NumberDays1 = 0
If StartYear > 1960 Then
    Do Until YearIndex = StartYear
        If Bisiesto1(YearIndex) = 1 Then
            NumberDays1 = NumberDays1 + 366
        Else: NumberDays1 = NumberDays1 + 365
        End If
    YearIndex = YearIndex + 1
    Loop
Else: NumberDays1 = 0
End If
DaysToFirstYear = NumberDays1
End Function


' compute number of days from January 1 of the first year on data file to first day on data file
Public Function DaysOnYear(StartYear, StartMonth)
Dim totaldays As Integer
Select Case StartMonth
    Case Is = 1
        totaldays = 0
    Case Is = 2
        totaldays = 31
    Case Is = 3
        totaldays = 59
    Case Is = 4
        totaldays = 90
    Case Is = 5
        totaldays = 120
    Case Is = 6
        totaldays = 151
    Case Is = 7
        totaldays = 181
    Case Is = 8
        totaldays = 212
    Case Is = 9
        totaldays = 243
    Case Is = 10
        totaldays = 273
    Case Is = 11
        totaldays = 304
    Case Is = 12
        totaldays = 334
End Select

If StartMonth > 2 Then
    If Bisiesto1(StartYear) = 1 Then
        totaldays = totaldays + 1
    End If
End If
DaysOnYear = totaldays

End Function

Public Function NumberOfDays(InitialYear, InitialMonth, InitialDay)
' this function computes the total number of days since January 1 of 1960 to the given date
Dim totaldays, DaysPrevYears As Integer

' compute the number of days from January 1 1960 to January 1 of the first year in file
DaysPrevYears = DaysToFirstYear(InitialYear)

' compute number of days from January 1 of the first year on data file to first day on data file
totaldays = DaysOnYear(InitialYear, InitialMonth)
totaldays = DaysPrevYears + totaldays + (InitialDay - 1)
NumberOfDays = totaldays
End Function

Public Function ReadSamson(LineFromFile)
'this piece of code read the line from the weather file, and coverts it into a string separated by "TABS"

    Dim s1 As String, sItems1() As String
    s1 = LineFromFile
    sItems1() = Split(s1, " ") ' split the sting in subs trings (year, month, day, hour, tempºC, RH)
    p1 = 0
    
     Dim NewString As String
     NewString = ""
     
    For p1 = 0 To UBound(sItems1)
        Trim (sItems1(p1))
        If sItems1(p1) <> Empty Then
            NewString = NewString + sItems1(p1) + vbTab
        End If
    Next
    ReadSamson = NewString
End Function

Public Function DryingCost(FanKWH, HeaterKWH, InTonnes, DMLBin, AvgFinMC)
'this sub is to compute the total drying cost for the current yer (Ricardo Bartosik Master thesis)
'drying cost=energy cost + shrink cost


FanCost = FanKWH * ElectricityCost
If HeaterType = True Then
    HeaterCost = HeaterKWH * 3414 / 92000 * PropaneCost
Else
    HeaterCost = HeaterKWH * ElectricityCost
End If
EnergyCost = FanCost + HeaterCost

DMLCost = InTonnes * DMLBin / 100 * GrainPrice
FinalTonGrain = InTonnes * (1 - AvgInGrainMC / 100) / (1 - DesiredFinMC / 100)
If AvgFinMC < DesiredFinMC Then
    OverdryingCost = (FinalTonGrain - (InTonnes * (1 - AvgInGrainMC / 100) / (1 - AvgFinMC / 100))) * GrainPrice
Else
    OverdryingCost = 0
End If

ShrinkCost = DMLCost + OverdryingCost


DryingCost = (EnergyCost + ShrinkCost) / FinalTonGrain

End Function
