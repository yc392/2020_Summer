Inputs:
The program takes two file as input, initial.json and Weather.txt. 
The initial.json includes all the initial values of parameters, the information of different types of grains, and the required simulation date for CNA strategy. 
The Weather.txt includes weather forcast information, not neccessary after including the database querying.

How To Run:
You can set all the parameters in the initial.json to initialize your model.
The program has two modes, FixInlet and CNA_strat. You can switch the mode by changing the parameter "Model"->"Mode" from "FixInlet" to "CNA".
Under FixInlet mode, you can set the parameter "Model"->"Hours_Stop" to desired simulation time(hours).
Under CNA_strat mode, you can set all parameters in "Date" for your desired simulation time(date).

Expected Outputs:
The expected outputs of FixInlet mode are six .txt files. FixInlet_mccenter.txt, FixInlet_mcside.txt, FixInlet_tempcenter.txt, FixInlet_tempside.txt, FixInlet_dmlcenter.txt and FixInlet_dmlside.txt 
seperately include the moisture content of each layers at the center of the bin, the moisture content of each layers at the side of the bin, the temperature of each layers at the center of the bin,
the temperature of each layers at the side of the bin, the DML of each layers at the center of the bin and the temperature of each layers at the side of the bin.
The expected ouputs of CNA_strat mode include all files listed above and another _Fan.txt which describes the working of the fan.