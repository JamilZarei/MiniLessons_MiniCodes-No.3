#general library
import numpy as np
#Excel management library
import xlwings as xl
#Thermo library for flash calculatiobn
from thermo import *
from chemicals import *
from thermo.interaction_parameters import IPDB
#Load Excel inputs by openpyxl library
wb=xl.Book(r"C:\Users\RayanPartov\Desktop\Python_VBA_Flasher.xlsm")
ws=wb.sheets[0]
#loading components
zs_list=ws.range('D3:D12').value
# print(ws.range('D3:D15').value)
T_K=(ws.range('E9').value-32)/1.8+273.5
P_PSI=ws.range('E6').value*6894.76
# print(T_K,P_PSI)
#Flash Calculation
try:
    pure_constants = ChemicalConstantsPackage.from_IDs(['nitrogen','carbon dioxide','hydrogen sulfide','methane', 'ethane',
                                                               'propane','butane','pentane','hexane'
                                                               ])
    #C7+ component propetises & constants
    pseudos = ChemicalConstantsPackage(names=['C7+'],Tcs=[ws.range('D19').value], Pcs=[ws.range('D18').value*(10**6)],
                                   omegas=[ws.range('D20').value], MWs=[ws.range('D16').value])
    # print(pure_constants)
    # print('--------------------')
    # print(pseudos)
    constants = pure_constants[0] + pseudos
  
    properties = PropertyCorrelationsPackage(constants=constants)
    kijs = np.zeros((constants.N, constants.N)).tolist() # kijs left as zero in this example
    eos_kwargs = {'Pcs': constants.Pcs, 'Tcs': constants.Tcs, 'omegas': constants.omegas, 'kijs': kijs}
    gas = CEOSGas(PRMIX, eos_kwargs=eos_kwargs, HeatCapacityGases=properties.HeatCapacityGases)
    liquid = CEOSLiquid(PRMIX, eos_kwargs=eos_kwargs, HeatCapacityGases=properties.HeatCapacityGases)
    flasher = FlashVL(constants, properties, liquid=liquid, gas=gas)
    zs = zs_list
    PT = flasher.flash(T=T_K, P=P_PSI, zs=zs)
    #Save flash values in excel
    for i in range(3,13):
        ws.cells(i,6).value=PT.gas.zs[i-3]
        ws.cells(i,7).value=PT.liquid0.zs[i-3]
    ws.range('H3').value=PT.VF
except:
    for i in range(3,13):
        ws.cells(i,6).value=0
        ws.cells(i,7).value=PT.liquid0.zs[i-3]
    ws.range('H3').value=PT.VF

