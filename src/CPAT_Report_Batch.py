# -*- coding: utf-8 -*-
"""
Created on Thu Oct 10 14:21:20 2022

@author: JRPilan
"""

import sys
import csv
import os
import cympy
import cympy.eq
import cympy.db
import urllib.request
from xml.etree import ElementTree as ET
import re
import datetime as dt
import operator

# output folder for reports
report_folder = r'V:/z_JessePiland/CPAT/Reports'

# studies to load
study_folder = r'V:/z_JessePiland/CPAT/Studies/Base_test'

# reference folder, currently just for source impedances
ref_folder = r'V:/z_JessePiland/CPAT/References'
source_impedance_file = os.path.join(ref_folder, 'Bank_Impedances.csv')

op_region = 'DEC'

# if both false, will leave dg as is
connect_dg = False
disconnect_dg = True

print(f'Reports will be saved in the following location: {report_folder}')
if connect_dg:
    resultlocation = report_folder+'\CPAT_Output_DG_Connected\\'
elif disconnect_dg:
    resultlocation = report_folder+'\CPAT_Output_DG_Disconnected\\'
else:
    resultlocation = report_folder+'\CPAT_Output_DG_As_Is\\'
os.makedirs(resultlocation, exist_ok=True)



inverter_eqids = ['SOLAR','PHOTOVOLTAIC','BATTERY']

relay_settings = {
    'SEL-351S-6' : [
        'RID',
        'CTR',
        'E79',
        '51P1P',
        '51P1C',
        '51P1TD',
        '51G1P',
        '51G1C',
        '51G1TD',
        '50P1P',
        '50P2P',
        '50P3P',
        '50G1P',
        '50G2P',
        '50G3P',
        '67P1D',
        '67P2D',
        '67P3D',
        '67G1D',
        '67G2D',
        '67G3D',
        ],
    
    'SEL-351S-7' : [
        'RID',
        'CTR',
        'E79',
        '51P1P',
        '51P1C',
        '51P1TD',
        '51G1P',
        '51G1C',
        '51G1TD',
        '50P1P',
        '50P2P',
        '50P3P',
        '50G1P',
        '50G2P',
        '50G3P',
        '67P1D',
        '67P2D',
        '67P3D',
        '67G1D',
        '67G2D',
        '67G3D',
        ],
    
    'SEL-351-6-V0' : [
        'RID',
        'CTR',
        'E79',
        '51PP',
        '51PC',
        '51PTD',
        '51GP',
        '51GC',
        '51GTD',
        '50P1P',
        '50P4P',
        '50G1P',
        '50G3P',
        '67P1D',
        '67P4D',
        '67G1D',
        '67G3D',
        ],

    'SEL-351-7-V0' : [
        'RID',
        'CTR',
        'E79',
        '51PP',
        '51PC',
        '51PTD',
        '51GP',
        '51GC',
        '51GTD',
        '50P1P',
        '50P4P',
        '50G1P',
        '50G3P',
        '67P1D',
        '67P4D',
        '67G1D',
        '67G3D',
        ],

    'SEL-351' : [
        'RID',
        'CTR',
        'CTRN',
        'E79',
        '51PP',
        '51PC',
        '51PTD',
        '51NP',
        '51NC',
        '51NTD',
        '50P1P',
        '50P2P',
        '50P3P',
        '50N1P',
        '50N2P',
        '50N3P',
        '67P1D',
        '67P2D',
        '67P3D',
        '67N1D',
        '67N2D',
        '67N3D',
        ],

    'SEL-351R-2' : [
        'RID',
        'CTR',
        'E79',
        '51P1P',
        '51P1C',
        '51P1TD',
        '51G1P',
        '51G1C',
        '51G1TD',
        '50P2P',
        '67P2D',
        '50G2P',
        '67G2D',
        ],
    
    'SEL-351R-4' : [
        'RID',
        'CTR',
        'E79',
        '51P1P',
        '51P1C',
        '51P1TD',
        '51G1P',
        '51G1C',
        '51G1TD',
        '50P2P',
        '67P2D',
        '50G2P',
        '67G2D',
        ],
}


other_relays =  [
    'IAC51',
    'IAC53',
    'CO-8',
    # 'SEL-351R-4',
    # 'SEL-351S-7',
    # 'SEL-351-7-V0',
]

# dictionaries for objects
circuits = {}
banks = {}
stations = {}
relays = {}


voltage_classes = {
    # 'DEP' : [12, 23, 34.5],
    'DEC': [4.16, 12.47, 23.9, 24.94, 34.5],
}

pdts = [
    cympy.enums.DeviceType.Breaker,
    cympy.enums.DeviceType.Recloser,
    cympy.enums.DeviceType.Fuse,
]
sdts = [
    cympy.enums.DeviceType.Switch,
    cympy.enums.DeviceType.Sectionalizer,
]

pds = [] #protective devices
sds = [] #switching devices

class Station(object):
    def __init__(self, staid, name, aspenid, size):
        self.ID = staid
        self.Name = name
        self.Aspen = aspenid
        self.Size = size
        self.Relays = {}
    
    def __repr__(self):
        return str(self.Name) + ' - ID: ' + str(self.ID)

class Relay(object):
    def __init__(self, circuitnum = '0000', breakername = '0000', relayid = 0, relaytype = 'Unknown', ctr = 0, hltph = 0, hltg = 0):
        self.ID = relayid
        self.CircuitNumber = circuitnum
        self.BreakerName = breakername
        self.Type = relaytype
        # self.CTR = ctr
        # self.HLT_ph = hltph
        # self.HLT_g = hltg
    
    def __repr__(self):
        return str(self.CircuitNumber) + ' ID: ' + str(self.ID)


class Circuit(object):
    def __init__(self, name, region = op_region, sections = []):
        self.name = name
        self.Sections = sections

        sources = cympy.study.ListDevices(cympy.enums.DeviceType.Source, name)
        if sources:
            source = sources[0]
            source_eq = cympy.eq.GetEquipment(source.EquipmentID, cympy.enums.EquipmentType.Substation)
            self.EqID = source.EquipmentID
            if source.EquipmentID != 'DEFAULT':
                self.StationID = source.EquipmentID.split('_')[-2] # e.g. DEC_3002_1
                self.Bank = source.EquipmentID.split('_')[-1]
            else:
                self.StationID = 'Unknown'
                self.Bank = '0'
            self.kV = tryfloat(source_eq.GetValue('NominalKVLL'))
            self.kV_class = min(voltage_classes[region], key = lambda x:abs(x-self.kV))
            self.SourceNode = cympy.study.ListNodes(cympy.enums.NodeType.SourceNode,net)[0]
        else: # should handle Equivalent Sources
            self.EqID = 'Unknown'
            self.StationID = '0000'
            self.Bank = '0'
            try:
                self.kV = tryfloat(cympy.study.GetValueTopo('Sources[0].EquivalentSourceModels['+str(cympy.study.GetActiveLoadModel().ID - 1)+'].EquivalentSource.KVLL',name))
                self.kV_class = min(voltage_classes[region], key = lambda x:abs(x-self.kV))
                self.SourceNode = cympy.study.ListNodes(cympy.enums.NodeType.SourceNode,net)[0]
            except:
                self.kV = None
                self.kV_class = None
                self.SourceNode = None

        self.State = cympy.study.GetValueTopo('Group3',name).replace(' ','_')
        self.Substation = cympy.study.GetValueTopo('Group1',name).replace(' ','_')
        self.Area = 'Unknown'



class Bank(object):
    '''Contains substation transformer bank information gathered from references'''
    def __init__(self, eqID, name, kv=12.47, r1=0, x1=0, r2=0, x2=0, r0=0, x0=0,  lll=0, lg=0):
        self.EquipmentID = eqID # e.g. DEC_3002_1
        self.name = name    # e.g. ABBEYVIL RET
        self.R1 = r1
        self.X1 = x1
        self.R2 = r2
        self.X2 = x2
        self.R0 = r0
        self.X0 = x0
        self.LLL = lll
        self.LG = lg
        self.KV_nom = kv
        self.KV_op = kv


class DevicePlus:
    def __init__(self, dev):
        self.Device = dev
        self.DeviceNumber = dev.DeviceNumber
        self.DeviceType = dev.DeviceType
        self.DeviceType_english = english_devicetype(self.DeviceType)
        if self.DeviceType in [getattr(cympy.enums.DeviceType,x) for x in ['OverheadLine','OverheadByPhase','OverheadLineUnbalanced','Underground']]: # wires use downstream
            self.EID = devid_parser(dev, 'eid2')
        else:
            self.EID = devid_parser(dev, 'eid')
        self.Smallworld = devid_parser(dev, 'Smallworld')
        if self.DeviceType == cympy.enums.DeviceType.Source:  # Sources aren't on sections, they are nodes, among other differences
            self.EquipmentID = dev.EquipmentID
            self.NetworkID = dev.DeviceNumber
            self.Phase = 'ABC'
        else:
            self.Section = cympy.study.GetSection(dev.SectionID)
            self.EquipmentID = cympy.study.QueryInfoDevice('EqId',self.DeviceNumber,self.DeviceType)
            self.NetworkID= dev.NetworkID
            self.Phase = self.Section.GetValue('Phase')
        self.Station = cympy.study.GetValueTopo('Group1',self.NetworkID)
        self.kVLL = tryfloat(cympy.study.QueryInfoDevice('KVLLBase',self.DeviceNumber,self.DeviceType))
        self.Bank = cympy.study.GetValueTopo('Group2',self.NetworkID)
        self.Parent = cympy.study.QueryInfoDevice('UpstreamProtId',self.DeviceNumber,self.DeviceType)
        self.ParentType = 'N/A'
        self.ParentType_english = cympy.study.QueryInfoDevice('UpstreamProtType',self.DeviceNumber,self.DeviceType)
        self.ParentEqID = 'N/A'
        self.ProtectionLevel = cympy.study.QueryInfoDevice('ProtLevel',self.DeviceNumber,self.DeviceType)
        # self.Loading_Query()
        # self.ShortCircuit_Query()
        try:
            self.Status = self.Device.GetValue('ConnectionStatus')
        except:
            self.Status = cympy.study.QueryInfoDevice('EqStatus', self.DeviceNumber, self.DeviceType)
        self.OpenClose = cympy.study.QueryInfoDevice('EqState',self.DeviceNumber,self.DeviceType)
        if self.OpenClose != '':
            self.Status += self.OpenClose
        
        self.Placeholder = 'PLACEHOLDER'
        
    def __repr__(self):
        return self.__class__.__name__ + ' ' + self.EquipmentID + ' ' + self.DeviceNumber

    def Loading_Query(self):
        self.IA =  tryfloat(cympy.study.QueryInfoDevice('IA',self.DeviceNumber,self.DeviceType))
        self.IB =  tryfloat(cympy.study.QueryInfoDevice('IB',self.DeviceNumber,self.DeviceType))
        self.IC =  tryfloat(cympy.study.QueryInfoDevice('IC',self.DeviceNumber,self.DeviceType))
        self.I_max = max(self.IA,self.IB,self.IC)

        self.ckVAA =  round(tryfloat(cympy.study.QueryInfoDevice('DwSKVAA',self.DeviceNumber,self.DeviceType)),0)
        self.ckVAB =  round(tryfloat(cympy.study.QueryInfoDevice('DwSKVAB',self.DeviceNumber,self.DeviceType)),0)
        self.ckVAC =  round(tryfloat(cympy.study.QueryInfoDevice('DwSKVAC',self.DeviceNumber,self.DeviceType)),0)
        self.ckVA_total = sum(getattr(self,'ckVA'+ph) for ph in 'ABC')
        
               
    def ShortCircuit_Query(self):
        self.LLLa = round(tryfloat(cympy.study.QueryInfoDevice('LLLAamp',self.DeviceNumber,self.DeviceType)))
        self.LLLb = round(tryfloat(cympy.study.QueryInfoDevice('LLLBamp',self.DeviceNumber,self.DeviceType)))
        self.LLLc = round(tryfloat(cympy.study.QueryInfoDevice('LLLCamp',self.DeviceNumber,self.DeviceType)))
        
        self.LLLGa = round(tryfloat(cympy.study.QueryInfoDevice('LLLGAamp',self.DeviceNumber,self.DeviceType)))
        self.LLLGb = round(tryfloat(cympy.study.QueryInfoDevice('LLLGBamp',self.DeviceNumber,self.DeviceType)))
        self.LLLGc = round(tryfloat(cympy.study.QueryInfoDevice('LLLGCamp',self.DeviceNumber,self.DeviceType)))

        self.LLGaba = round(tryfloat(cympy.study.QueryInfoDevice('LLGAB_Aamp',self.DeviceNumber,self.DeviceType)))
        self.LLGabb = round(tryfloat(cympy.study.QueryInfoDevice('LLGAB_Bamp',self.DeviceNumber,self.DeviceType)))
        
        self.LLGbcb = round(tryfloat(cympy.study.QueryInfoDevice('LLGBC_Bamp',self.DeviceNumber,self.DeviceType)))
        self.LLGbcc = round(tryfloat(cympy.study.QueryInfoDevice('LLGBC_Camp',self.DeviceNumber,self.DeviceType)))
        
        self.LLGcac = round(tryfloat(cympy.study.QueryInfoDevice('LLGCA_Camp',self.DeviceNumber,self.DeviceType)))
        self.LLGcaa = round(tryfloat(cympy.study.QueryInfoDevice('LLGCA_Aamp',self.DeviceNumber,self.DeviceType)))
        
        self.LLa = round(tryfloat(cympy.study.QueryInfoDevice('LLABamp',self.DeviceNumber,self.DeviceType)))
        self.LLb = round(tryfloat(cympy.study.QueryInfoDevice('LLBCamp',self.DeviceNumber,self.DeviceType)))
        self.LLc = round(tryfloat(cympy.study.QueryInfoDevice('LLCAamp',self.DeviceNumber,self.DeviceType)))
        
        self.LGa = round(tryfloat(cympy.study.QueryInfoDevice('LGAamp',self.DeviceNumber,self.DeviceType)))
        self.LGb = round(tryfloat(cympy.study.QueryInfoDevice('LGBamp',self.DeviceNumber,self.DeviceType)))
        self.LGc = round(tryfloat(cympy.study.QueryInfoDevice('LGCamp',self.DeviceNumber,self.DeviceType)))
        
        self.LLL_max = max([self.LLLa, self.LLLb, self.LLLc, self.LLLGa, self.LLLGb, self.LLLGc])
        
        self.LL_max = max([self.LLa, self.LLb, self.LLc])
        try:
            self.LL_min = min([x for x in [self.LLa, self.LLb, self.LLc] if x > 0.001])
        except:
            self.LL_min = 0
        
        self.LG_max = max([self.LGa, self.LGb, self.LGc])
        try:
            self.LG_min = min([x for x in [self.LGa, self.LGb, self.LGc] if x > 0.001])
        except:
            self.LG_min = 0
            
        self.MaxPhaseFault = max([self.LLL_max, self.LL_max, self.LG_max])
        self.MaxGroundFault = max([self.LLLGa, self.LLLGb, self.LLLGc, self.LLGaba, self.LLGabb, self.LLGbcb, self.LLGbcc, self.LLGcac, self.LLGcac, self.LLGcaa, self.LG_max])
    
    def Upstream_Query(self):
        self.Upstream_Sections = sections_upstream(self.Device)
    
    def Update_Environment(self):
        upiter = cympy.study.NetworkIterator(self.Section.FromNode.ID,cympy.enums.IterationOption.Upstream)
        
        while upiter.Next():
            updevs = upiter.GetDevices()
            if any(x.DeviceType == cympy.enums.DeviceType.OverheadByPhase for x in updevs):
                self.Environment = 'Overhead'
                break
            elif any(x.DeviceType == cympy.enums.DeviceType.Underground for x in updevs):
                self.Environment = 'Underground'
                break
            else:
                self.Environment = 'Unknown'
    
    def SpotLoad_Query(self):
        tx.SpotKVA =  float(cympy.study.QueryInfoDevice('SpotCKVAT',tx.DeviceNumber,tx.DeviceType))

    
    def Downstream_Query(self):
        self.Customers_A = cympy.study.QueryInfoDevice('DwCUSTAT',self.DeviceNumber,self.DeviceType)
        self.Customers_B = cympy.study.QueryInfoDevice('DwCUSTBT',self.DeviceNumber,self.DeviceType)
        self.Customers_C = cympy.study.QueryInfoDevice('DwCUSTCT',self.DeviceNumber,self.DeviceType)
        self.Customers_Total = cympy.study.QueryInfoDevice('DwCustT',self.DeviceNumber,self.DeviceType)
        
        # self.Sections_Downstream = sections_downstream(self.Device)
        
        
        self.TXs = []
        self.TXs_3ph = []
        self.TXs_OH = []
        self.TXs_UG = []
        
        for tx in txs:
            if self.Section in tx.Upstream_Sections:
                self.TXs.append(tx)
                if tx.Phase == 'ABC':
                    self.TXs_3ph.append(tx)
                if tx.Environment == 'Overhead':
                    self.TXs_OH.append(tx)
                elif tx.Environment == 'Underground':
                    self.TXs_UG.append(tx)
        
        # for s in self.Sections_Downstream:
        #     self.TXs.extend(s.ListDevices(cympy.enums.DeviceType.SpotLoad))
        #     if s.GetValue('Phase') == 'ABC':
        #         self.TXs_3ph.extend(s.ListDevices(cympy.enums.DeviceType.SpotLoad))
        
        

        # for tx in self.TXs:
        #     tx_devplus = DevicePlus(tx)
        #     tx_devplus.ShortCircuit_Query()
        #     tx_devplus.SpotKVA =  float(cympy.study.QueryInfoDevice('SpotCKVAT',tx.DeviceNumber,tx.DeviceType))
        #     tx_sec = tx_devplus.Section
        #     txiter = cympy.study.NetworkIterator(tx_sec.FromNode.ID,cympy.enums.IterationOption.Upstream)
            
        #     while txiter.Next():
        #         tx_updevs = txiter.GetDevices()
        #         if any(x.DeviceType == cympy.enums.DeviceType.OverheadByPhase for x in tx_updevs):
        #             self.TXs_OH.append(tx_devplus)
        #             break
        #         elif any(x.DeviceType == cympy.enums.DeviceType.Underground for x in tx_updevs):
        #             self.TXs_UG.append(tx_devplus)
        #             break
                
        # self.TXs_OH.sort(key=lambda x: float(cympy.study.QueryInfoDevice('SpotCKVAT',x.DeviceNumber,x.DeviceType)), reverse=True)
        # self.TXs_UG.sort(key=lambda x: float(cympy.study.QueryInfoDevice('SpotCKVAT',x.DeviceNumber,x.DeviceType)), reverse=True)
        
        self.TXs_OH = sorted(self.TXs_OH, key = lambda x: (x.SpotKVA, x.LG_max), reverse=True)
            # (float(cympy.study.QueryInfoDevice('SpotCKVAT',x.DeviceNumber,x.DeviceType)), x.LGMax), reverse=True)

        self.TXs_UG = sorted(self.TXs_UG, key = lambda x: (x.SpotKVA, x.LG_max), reverse=True)
            # (float(cympy.study.QueryInfoDevice('SpotCKVAT',x.DeviceNumber,x.DeviceType)), x.LGMax), reverse=True)
        
            
        if len(self.TXs_OH)>0:
            largest_oh = self.TXs_OH[0]
            self.Largest_TX_OH_cKVA = largest_oh.SpotKVA
            self.Largest_TX_OH_PhaseFault = largest_oh.MaxPhaseFault
            self.Largest_TX_OH_GroundFault = largest_oh.MaxGroundFault
        else:
            self.Largest_TX_OH_cKVA = 'N/A'
            self.Largest_TX_OH_PhaseFault = 'N/A'
            self.Largest_TX_OH_GroundFault = 'N/A'
            
        if len(self.TXs_UG)>0:
            largest_ug = self.TXs_UG[0]
            self.Largest_TX_UG_cKVA = largest_ug.SpotKVA
            self.Largest_TX_UG_PhaseFault = largest_ug.MaxPhaseFault
            self.Largest_TX_UG_GroundFault = largest_ug.MaxGroundFault
        else:
            self.Largest_TX_UG_cKVA = 'N/A'
            self.Largest_TX_UG_PhaseFault = 'N/A'
            self.Largest_TX_UG_GroundFault = 'N/A'
        
        self.ThreePhase_ckVA = sum(float(cympy.study.QueryInfoDevice('SpotCKVAT', x.DeviceNumber, x.DeviceType)) for x in self.TXs_3ph)
        self.ThreePhase_Large_ckVA = sum(float(cympy.study.QueryInfoDevice('SpotCKVAT', x.DeviceNumber, x.DeviceType)) for x in self.TXs_3ph if float(cympy.study.QueryInfoDevice('SpotCKVAT', x.DeviceNumber, x.DeviceType))>=300)
        
        
    def Zone_Query(self):
        self.Sections = sections_in_zone(self.Device)
        self.Sections.append(self.Section) # in case device in question is conductor
        self.Conductors_OH = []
        self.Conductors_UG = []

        self.DERs = []
        
        for s in self.Sections:
            self.Conductors_OH.extend(s.ListDevices(cympy.enums.DeviceType.OverheadByPhase))
            self.Conductors_UG.extend(s.ListDevices(cympy.enums.DeviceType.Underground))
            self.DERs.extend(s.ListDevices(cympy.enums.DeviceType.ElectronicConverterGenerator))
        
        self.Conductors_OH_A = [(cympy.study.QueryInfoDevice('CondAId',oh.DeviceNumber,oh.DeviceType),float(cympy.study.QueryInfoDevice('CondAamp',oh.DeviceNumber,oh.DeviceType))) for oh in self.Conductors_OH if cympy.study.QueryInfoDevice('CondAId',oh.DeviceNumber,oh.DeviceType) != '']
        self.Conductors_OH_B = [(cympy.study.QueryInfoDevice('CondBId',oh.DeviceNumber,oh.DeviceType),float(cympy.study.QueryInfoDevice('CondBamp',oh.DeviceNumber,oh.DeviceType))) for oh in self.Conductors_OH if cympy.study.QueryInfoDevice('CondBId',oh.DeviceNumber,oh.DeviceType) != '']
        self.Conductors_OH_C = [(cympy.study.QueryInfoDevice('CondCId',oh.DeviceNumber,oh.DeviceType),float(cympy.study.QueryInfoDevice('CondCamp',oh.DeviceNumber,oh.DeviceType))) for oh in self.Conductors_OH if cympy.study.QueryInfoDevice('CondCId',oh.DeviceNumber,oh.DeviceType) != '']
        self.Conductors_OH_Neutral = [(cympy.study.QueryInfoDevice('Neutral1ID',oh.DeviceNumber,oh.DeviceType),float(cympy.study.QueryInfoDevice('Neutralamp',oh.DeviceNumber,oh.DeviceType))) for oh in self.Conductors_OH if cympy.study.QueryInfoDevice('Neutral1ID',oh.DeviceNumber,oh.DeviceType) != '']
        
        # sort by ampacity
        self.Conductors_OH_A.sort(key = lambda x: x[1])
        self.Conductors_OH_B.sort(key = lambda x: x[1])
        self.Conductors_OH_C.sort(key = lambda x: x[1])
        self.Conductors_OH_Neutral.sort(key = lambda x: x[1])
        
        for phase in ['A','B','C','Neutral']:
            if len(getattr(self, f'Conductors_OH_{phase}'))>0:
                smallest = getattr(self, f'Conductors_OH_{phase}')[0]
                # smallest_cond_id = smallest[0]
                setattr(self, f'Conductor_OH_Smallest_{phase}', smallest)
            else:
                setattr(self, f'Conductor_OH_Smallest_{phase}', ('N/A',999999))
                
        self.Conductors_UG_A = [(ug.EquipmentID, float(cympy.study.QueryInfoDevice('CableAmpacity',ug.DeviceNumber,ug.DeviceType))) for ug in self.Conductors_UG if 'A' in cympy.study.QueryInfoDevice('Phase',ug.DeviceNumber,ug.DeviceType)]
        self.Conductors_UG_B = [(ug.EquipmentID, float(cympy.study.QueryInfoDevice('CableAmpacity',ug.DeviceNumber,ug.DeviceType))) for ug in self.Conductors_UG if 'B' in cympy.study.QueryInfoDevice('Phase',ug.DeviceNumber,ug.DeviceType)]
        self.Conductors_UG_C = [(ug.EquipmentID, float(cympy.study.QueryInfoDevice('CableAmpacity',ug.DeviceNumber,ug.DeviceType))) for ug in self.Conductors_UG if 'C' in cympy.study.QueryInfoDevice('Phase',ug.DeviceNumber,ug.DeviceType)]
        self.Conductors_UG_Neutral = [(ug.EquipmentID, float(cympy.study.QueryInfoDevice('CableAmpacity',ug.DeviceNumber,ug.DeviceType))) for ug in self.Conductors_UG]
        
        # sort by ampacity
        self.Conductors_UG_A.sort(key = lambda x: x[1])
        self.Conductors_UG_B.sort(key = lambda x: x[1])
        self.Conductors_UG_C.sort(key = lambda x: x[1])
        self.Conductors_UG_Neutral.sort(key = lambda x: x[1])
        
        for phase in ['A','B','C','Neutral']:
            if len(getattr(self, f'Conductors_UG_{phase}'))>0:
                smallest = getattr(self, f'Conductors_UG_{phase}')[0]
                # smallest_cond_id = smallest[0]
                setattr(self, f'Conductor_UG_Smallest_{phase}', smallest)
            else:
                setattr(self, f'Conductor_UG_Smallest_{phase}', ('N/A',999999))
        
        for phase in ['A','B','C','Neutral']:
            oh = getattr(self, f'Conductor_OH_Smallest_{phase}')
            ug = getattr(self, f'Conductor_UG_Smallest_{phase}')
            if oh[1] <= ug[1]:
                setattr(self, f'Smallest_Conductor_{phase}', oh[0])
            else:
                setattr(self, f'Smallest_Conductor_{phase}', ug[0])
        
        
        # self.Conductors_OH_A = [oh for oh in self.Conductors_OH if cympy.study.QueryInfoDevice('CondAId',oh.DeviceNumber,oh.DeviceType) != '']
        # self.Conductors_OH_B = [oh for oh in self.Conductors_OH if cympy.study.QueryInfoDevice('CondBId',oh.DeviceNumber,oh.DeviceType) != '']
        # self.Conductors_OH_C = [oh for oh in self.Conductors_OH if cympy.study.QueryInfoDevice('CondCId',oh.DeviceNumber,oh.DeviceType) != '']
        # self.Conductors_OH_Neutral = [oh for oh in self.Conductors_OH if cympy.study.QueryInfoDevice('Neutral1ID',oh.DeviceNumber,oh.DeviceType) != '']
               
        # self.Conductors_OH_A.sort(key = lambda x:float(cympy.study.QueryInfoDevice('CondAamp', x.DeviceNumber, x.DeviceType)))
        # self.Conductors_OH_B.sort(key = lambda x:float(cympy.study.QueryInfoDevice('CondBamp', x.DeviceNumber, x.DeviceType)))
        # self.Conductors_OH_C.sort(key = lambda x:float(cympy.study.QueryInfoDevice('CondCamp', x.DeviceNumber, x.DeviceType)))
        # self.Conductors_OH_Neutral.sort(key = lambda x:float(cympy.study.QueryInfoDevice('Neutralamp', x.DeviceNumber, x.DeviceType)))
        
        # for phase in 'ABC':
        #     if len(getattr(self, f'Conductors_OH_{phase}'))>0:
        #         smallest = getattr(self, f'Conductors_OH_{phase}')[0]
        #         smallest_cond_id = cympy.study.QueryInfoDevice(f'Cond{phase}Id',smallest.DeviceNumber,smallest.DeviceType)
        #         setattr(self, f'Conductor_OH_Smallest_{phase}', smallest_cond_id)
        #     else:
        #         setattr(self, f'Conductor_OH_Smallest_{phase}', 'N/A')
                
        # if len(self.Conductors_OH_Neutral)>0:
        #     smallest = self.Conductors_OH_Neutral[0]
        #     self.Conductor_OH_Smallest_Neutral = cympy.study.QueryInfoDevice('Neutral1ID',smallest.DeviceNumber,smallest.DeviceType)
        # else:
        #     self.Conductor_UG_Smallest_Neutral = 'N/A'
        
        # self.Conductors_UG_A = [ug for ug in self.Conductors_UG if 'A' in cympy.study.QueryInfoDevice('Phase',ug.DeviceNumber,ug.DeviceType)]
        # self.Conductors_UG_B = [ug for ug in self.Conductors_UG if 'B' in cympy.study.QueryInfoDevice('Phase',ug.DeviceNumber,ug.DeviceType)]
        # self.Conductors_UG_C = [ug for ug in self.Conductors_UG if 'C' in cympy.study.QueryInfoDevice('Phase',ug.DeviceNumber,ug.DeviceType)]
        # self.Conductors_UG_Neutral = self.Conductors_UG.copy()
        
        # self.Conductors_UG_A.sort(key = lambda x:float(cympy.study.QueryInfoDevice('CableAmpacity', x.DeviceNumber, x.DeviceType)))
        # self.Conductors_UG_B.sort(key = lambda x:float(cympy.study.QueryInfoDevice('CableAmpacity', x.DeviceNumber, x.DeviceType)))
        # self.Conductors_UG_C.sort(key = lambda x:float(cympy.study.QueryInfoDevice('CableAmpacity', x.DeviceNumber, x.DeviceType)))
        # self.Conductors_UG_Neutral.sort(key = lambda x:float(cympy.study.QueryInfoDevice('CableAmpacity', x.DeviceNumber, x.DeviceType)))
        
        # for phase in 'ABC':
        #     if len(getattr(self, f'Conductors_UG_{phase}'))>0:
        #         smallest = getattr(self, f'Conductors_UG_{phase}')[0]
        #         setattr(self, f'Conductor_UG_Smallest_{phase}', smallest.EquipmentID)
        #     else:
        #         setattr(self, f'Conductor_UG_Smallest_{phase}', 'N/A')
        
        # if len(self.Conductors_UG_Neutral)>0:
        #     smallest = self.Conductors_UG_Neutral[0]
        #     self.Conductor_UG_Smallest_Neutral = smallest.EquipmentID
        # else:
        #     self.Conductor_UG_Smallest_Neutral = 'N/A'
            
        

        
        
        # self.Conductors_OH_A = [cympy.study.QueryInfoDevice('CondAID', oh.DeviceNumber, oh.DeviceType) for oh in self.Conductors_OH if 'A' in cympy.study.QueryInfoDevice('Phase',oh.DeviceNumber,oh.DeviceType)]
        # self.Conductors_OH_B = [cympy.study.QueryInfoDevice('CondBID', oh.DeviceNumber, oh.DeviceType) for oh in self.Conductors_OH if 'B' in cympy.study.QueryInfoDevice('Phase',oh.DeviceNumber,oh.DeviceType)]
        # self.Conductors_OH_C = [cympy.study.QueryInfoDevice('CondCID', oh.DeviceNumber, oh.DeviceType) for oh in self.Conductors_OH if 'C' in cympy.study.QueryInfoDevice('Phase',oh.DeviceNumber,oh.DeviceType)]
        # self.Conductors_OH_Neutral = [cympy.study.QueryInfoDevice('Neutral1ID', oh.DeviceNumber, oh.DeviceType) for oh in self.Conductors_OH]
        
        # self.Conductors_UG_A = [ug.EquipmentID for ug in self.Conductors_UG if 'A' in cympy.study.QueryInfoDevice('Phase',ug.DeviceNumber,ug.DeviceType)]
        # self.Conductors_UG_B = [ug.EquipmentID for ug in self.Conductors_UG if 'B' in cympy.study.QueryInfoDevice('Phase',ug.DeviceNumber,ug.DeviceType)]
        # self.Conductors_UG_C = [ug.EquipmentID for ug in self.Conductors_UG if 'C' in cympy.study.QueryInfoDevice('Phase',ug.DeviceNumber,ug.DeviceType)]
        # self.Conductors_UG_Neutral = [ug.EquipmentID for ug in self.Conductors_UG]
        

        

        self.DER_Gen_kW = sum(float(cympy.study.QueryInfoDevice('DERActiveGeneration', der.DeviceNumber, der.DeviceType)) for der in self.DERs)
        
class ProtectiveDevice(DevicePlus):
    def __init__(self, dev):
        super().__init__(dev)
    
    def Calculate_Bounds(self):
        self.Bounding_Devices = bounding_devices(self.Device)
        self.Bounding_Fuses = bounding_devices(self.Device, cympy.study.ListDevices(cympy.enums.DeviceType.Fuse))


class Breaker(ProtectiveDevice):
    def __init__(self, dev):
        super().__init__(dev)

class Recloser(ProtectiveDevice):
    def __init__(self, dev):
        super().__init__(dev)
                      

class Fuse(ProtectiveDevice):
    def __init__(self, dev):
        super().__init__(dev)
        self.Type = self.EquipmentID.rsplit('_',1)[0]
        self.Size = tryfloat(self.EquipmentID.rsplit('_',1)[1][:-1])




def english_devicetype(devtype):
    # for dt in ['Recloser','Fuse','Breaker','Sectionalizer','']:
    for devt in dir(cympy.enums.DeviceType):
        if not devt[0].isupper():
            continue
        if devtype == getattr(cympy.enums.DeviceType,devt):
            return  devt     
    return 'Unknown'

def tryfloat(s):
    try:
        return float(s)
    except:
        return 0

def net_add_leadingzeros(s):
    sparts = re.split(r'(\d+)',s)
    try:
        return sparts[0] + sparts[1].zfill(4)
    except:
        return s

def net_remove_leadingzeros(s):
    sparts = re.split(r'(\d+)',s)
    try:
        return sparts[0] + sparts[1].lstrip('0')
    except:
        return s

def q_distance(dev):
    return tryfloat(cympy.study.QueryInfoDevice('Distance',dev.DeviceNumber,dev.DeviceType))

def q_customers(dev):
    return tryfloat(cympy.study.QueryInfoDevice('DwCustT',dev.DeviceNumber,dev.DeviceType))

def upstream_node(dev):
    sec = cympy.study.GetSection(dev.SectionID)
    fn = sec.FromNode
    tn = sec.ToNode
    firstnode = fn # assume fromnode is upstream
    fn_d = tryfloat(cympy.study.QueryInfoNode('Distance',fn.ID))
    tn_d = tryfloat(cympy.study.QueryInfoNode('Distance',tn.ID))
    # first check for source nodes
    if fn.GetType() == cympy.enums.NodeType.SourceNode:
        return fn
    elif tn.GetType() == cympy.enums.NodeType.SourceNode:
        return tn    
    # next check if the device is at an interconnection
    if fn.GetType() == cympy.enums.NodeType.Interconnection:
        upiter = cympy.study.NetworkIterator(fn.ID,cympy.enums.IterationOption.Upstream)
        if sec in upiter.ListNextSections():
            return tn
        else:
            pass
    elif tn.GetType() == cympy.enums.NodeType.Interconnection:
        upiter = cympy.study.NetworkIterator(tn.ID,cympy.enums.IterationOption.Upstream)
        if sec in upiter.ListNextSections():
            return fn
        else:
            pass
    # check if the from node is further from the source than the to node
    if fn_d > tn_d > 0: 
        firstnode = tn
    elif fn_d > tn_d: # empty interconnection at to node
        firstnode = fn
    elif 0 < fn_d < tn_d:
        firstnode = fn
    elif fn_d < tn_d: # empty interconnection at from node
        firstnode = tn
    else:
        # if the chosen device is a protective device (on a zero-length section), use its FeedingNode
        try:
            nfnid = dev.GetValue('NormalFeedingNodeID')
            if nfnid == fn.ID:
                firstnode = fn
            elif nfnid == tn.ID:
                firstnode = tn
        except cympy.err.CymError:
            pass
    return firstnode

def downstream_node(dev):
    sec = cympy.study.GetSection(dev.SectionID)
    fn = sec.FromNode
    tn = sec.ToNode
    firstnode = tn # assume fromnode is upstream
    fn_d = tryfloat(cympy.study.QueryInfoNode('Distance',fn.ID))
    tn_d = tryfloat(cympy.study.QueryInfoNode('Distance',tn.ID))
    # first check for source nodes
    if fn.GetType() == cympy.enums.NodeType.SourceNode:
        return tn
    elif tn.GetType() == cympy.enums.NodeType.SourceNode:
        return fn    
    # next check if the device is at an interconnection
    if fn.GetType() == cympy.enums.NodeType.Interconnection:
        upiter = cympy.study.NetworkIterator(fn.ID,cympy.enums.IterationOption.Upstream)
        if sec in upiter.ListNextSections():
            return fn
        else:
            pass
    elif tn.GetType() == cympy.enums.NodeType.Interconnection:
        upiter = cympy.study.NetworkIterator(tn.ID,cympy.enums.IterationOption.Upstream)
        if sec in upiter.ListNextSections():
            return tn
        else:
            pass
    # check if the from node is further from the source than the to node
    if fn_d > tn_d > 0: 
        firstnode = fn
    elif fn_d > tn_d: # interconnection at to node
        firstnode = tn
    elif 0 < fn_d < tn_d:
        firstnode = tn
    elif fn_d < tn_d: # interconnection at from node
        firstnode = fn
    else:
        # if the chosen device is a protective device (on a zero-length section), use its FeedingNode
        try:
            nfnid = dev.GetValue('NormalFeedingNodeID')
            if nfnid == fn.ID:
                firstnode = tn
            elif nfnid == tn.ID:
                firstnode = fn
        except cympy.err.CymError: 
            pass
    return firstnode
    
def bounding_devices(dev, protdevs = pds): # list of devices bounding downstream zone
    startnode = downstream_node(dev)
    startnet = dev.NetworkID
    iterdirection = cympy.enums.IterationOption.Downstream
    iterrestriction = cympy.enums.IterationRestriction.SameTopo
    bdev_list = []
    expiter = cympy.study.NetworkIterator(startnode.ID,iterdirection,iterrestriction)
    
    while expiter.Next():
        if expiter.GetNetworkID() != startnet:
            expiter.Skip()
        dssec = expiter.GetSection()
        #check for protective devices on current section
        for d in dssec.ListDevices():
            if d == dev:
                continue
            if d in protdevs:
                bdev_list.append(d)
                expiter.Skip()
    return bdev_list

def parent_device(dev, protdevs = pds):
    net = dev.NetworkID
    if dev == null_device:
        return null_device
    fn = upstream_node(dev)
    parent_iter = cympy.study.NetworkIterator(fn.ID, cympy.enums.IterationOption.Upstream, cympy.enums.IterationRestriction.SameTopo)
    while parent_iter.Next():
        if parent_iter.GetNetworkID() != net:
            parent_iter.Skip()
        upsec = parent_iter.GetSection()
        for d in upsec.ListDevices():
            if d in protdevs:
                # if d.GetValue('ClosedPhase') == 'None':
                #     parent_iter.Skip()
                # else:
                #     return d
                return d
    return null_device # return null device if no upstream device found

def sections_in_zone(dev = cympy.study.Device(), protdevs = pds, phase_restriction = 0, startnode = '') :
    iterdirection = cympy.enums.IterationOption.Downstream
    iterrestriction = cympy.enums.IterationRestriction.SameTopo
    zone_sec_list = []
    sec_pdevlist = []
    if not(startnode):
        startnode = downstream_node(dev)
    expiter = cympy.study.NetworkIterator(startnode.ID,iterdirection,iterrestriction)
    
    while expiter.Next():
        del(sec_pdevlist[:])
        skip_branch = False
        dssec = expiter.GetSection()
        if len(dssec.GetValue('Phase')) < phase_restriction:
            skip_branch = True
        #list protective devices on current section
        # for pdt in pdts:
        #     sec_pdevlist.extend(dssec.ListDevices(pdt))
        #list devices on current section
        sec_pdevlist = dssec.ListDevices()
        for spd in sec_pdevlist:
            if spd in protdevs:
                skip_branch = True
        if skip_branch:
            expiter.Skip()
        else:
            zone_sec_list.append(dssec)
    return zone_sec_list

def exposure(dev,startlength, protdevs = pds):
    startnode = downstream_node(dev)
    net = dev.NetworkID
    iterdirection = cympy.enums.IterationOption.Downstream
    iterrestriction = cympy.enums.IterationRestriction.SameTopo
    meters = startlength
    sec_pdevlist = []
    expiter = cympy.study.NetworkIterator(startnode.ID,iterdirection,iterrestriction)
    
    while expiter.Next():
        if expiter.GetNetworkID() != net:
            expiter.Skip()
        del(sec_pdevlist[:])
        skip_branch = False
        dssec = expiter.GetSection()
        #check if protective device on current section
        for d in dssec.ListDevices():
            if d in protdevs:
                skip_branch = True
            elif d.DeviceType == cympy.enums.DeviceType.Switch:
                if d.GetValue('ClosedPhase') == 'None':
                    skip_branch = True
        try:
            if dssec.ListDevices()[0].GetValue('ConnectionStatus') == 'Disconnected':
                skip_branch = True
        except:
            pass
        if skip_branch:
            expiter.Skip()
        else:
            meters += float(dssec.Length)
    
    return meters/1609.344

def sections_upstream(dev, protdevs = []):
    startnode = upstream_node(dev)
    iterdirection = cympy.enums.IterationOption.Upstream
    iterrestriction = cympy.enums.IterationRestriction.SameTopo
    seclist = []
    upiter = cympy.study.NetworkIterator(startnode.ID,iterdirection,iterrestriction)
    
    while upiter.Next():
        upsec = upiter.GetSection()
        seclist.append(upsec)
    
    return seclist

def sections_downstream(dev, protdevs = []):
    startnode = downstream_node(dev)
    iterdirection = cympy.enums.IterationOption.Downstream
    iterrestriction = cympy.enums.IterationRestriction.SameTopo
    seclist = []
    upiter = cympy.study.NetworkIterator(startnode.ID,iterdirection,iterrestriction)
    
    while upiter.Next():
        upsec = upiter.GetSection()
        seclist.append(upsec)
    
    return seclist

def secid_parser(sec, part = ''):
    secsplit = sec.ID.split('_')
    parse_dict = {
        'eid'       :    0,
        'eid2'      :    1,
        'smallworld':   -1,
    }
    try:
        parsed = secsplit[parse_dict[part.lower()]]
    except:
        parsed = sec.ID
    return parsed
    
def devid_parser(dev, part = ''):
    if (dev.DeviceType == cympy.enums.DeviceType.RegulatorByPhase): # and (dev.NetworkID in dev.DeviceNumber):
        parsed = dev.DeviceNumber
    else:
        devsplit = dev.DeviceNumber.split('_')
        parse_dict = {
            'eid'       :    0,
            'eid2'      :    1,
            'smallworld':   -1,
        }
        try:
            parsed = devsplit[parse_dict[part.lower()]]
        except:
            parsed = dev.DeviceNumber
    return parsed


for studyfile in sorted(os.listdir(study_folder)):
    print('Opening ' + studyfile + '...')
    cympy.study.Open(os.path.join(study_folder,studyfile))
    
    # CYME setup
    null_device = cympy.study.Device()
    
    lm = cympy.study.GetActiveLoadModel()
    
    basevoltage = cympy.env.BaseVoltage
    cympy.app.ActivateRefresh(False)
    cympy.env.SetDataValidationOption(cympy.enums.DataValidationOption.AllowChangePrimaryKey, True)
    
    netlist = cympy.study.ListNetworks()
    netlist.sort()
    
    sc = cympy.sim.ShortCircuit()
    lf = cympy.sim.LoadFlow()
    
    devices = {(x.DeviceNumber, x.DeviceType):x for x in cympy.study.ListDevices()}
    
    for pdt in pdts:
        pds.extend(cympy.study.ListDevices(pdt))
    sds = pds.copy()
    for sdt in sdts:
        sds.extend(cympy.study.ListDevices(sdt))
    sdts += pdts
    
    # get relay info from DEC stations website
    
    banks_loaded = [source.EquipmentID for source in cympy.study.ListDevices(cympy.enums.DeviceType.Source)]
    stations_loaded = [bank.split('_')[-2] for bank in banks_loaded if 'DEFAULT' not in bank]
    
    ### populate stations dictionary with Station objects
    stations_url = 'http://sysprot.duke-energy.com/relay_info/decv2/aspen.asp?a=1&c=DEC'
    sta_response = urllib.request.urlopen(stations_url).read()
    station_tree = ET.fromstring(sta_response)
    for child in station_tree:
        sta_name = child.attrib['S01']
        sta_num = int(tryfloat(child.attrib['STANUM']))
        sta_aspenid = child.attrib['ID'].strip()
        sta_size = tryfloat(child.attrib['S04'])
    
        stations[sta_num] = Station(sta_num, sta_name, sta_aspenid, sta_size)
    
    ### find circuit breakers
    # for sta in list(stations.values())[:9]: # testing purposes
    for sta in stations.values():
        if str(sta.ID) not in stations_loaded:
            continue
        eq_url = 'http://sysprot.duke-energy.com/relay_info/decv2/aspen.asp?a=2&lid={0}&c=DEC'.format(sta.Aspen)
        eq_response = urllib.request.urlopen(eq_url).read()
        eq_tree = ET.fromstring(eq_response)
        
        if not eq_tree: # if no data found, try adding a space after aspenid (probably an error)
            eq_url = 'http://sysprot.duke-energy.com/relay_info/decv2/aspen.asp?a=2&lid={0}&c=DEC'.format(sta.Aspen+'%20')
            eq_response = urllib.request.urlopen(eq_url).read()
            eq_tree = ET.fromstring(eq_response)      
        
        cb_list = [eq.text for eq in list(eq_tree) if eq.text and eq.text.isnumeric() and len(eq.text)==4] # e.g. 1204
        cb_list.extend([eq.text for eq in list(eq_tree) if eq.text and re.search('^OCR\d+ \([0-9]{4}\)',eq.text)]) # e.g. OCR42848 (1204)
        cb_list.extend([eq.text for eq in list(eq_tree) if eq.text and re.search('^\d+ \([0-9]{4}\)',eq.text)]) # e.g. 13812 (1201)
        cb_list.extend([eq.text for eq in list(eq_tree) if eq.text and re.search('^[0-9]{4} \(\d+\)',eq.text)]) # e.g. 1201 (54812)
        cb_list.extend([eq.text for eq in list(eq_tree) if eq.text and re.search('^[0-9]{4} \(OCR\d+\)',eq.text)]) # e.g. 3404 (OCR51748) -- Marble Tie
        cb_list.extend([eq.text for eq in list(eq_tree) if eq.text and re.search('^[0-9]{4} \(NEW\)',eq.text)]) # e.g. 1203 (NEW) -- Belton Ret
        cb_list.extend([eq.text for eq in list(eq_tree) if eq.text and len(eq.text)==4 and re.search('^[0-9]{2}AX',eq.text)]) # e.g. 12AX -- Conway Ret
        cb_list.extend([eq.text for eq in list(eq_tree) if eq.text and re.search('^[0-9]{4} \[NEW\]',eq.text)]) # e.g. 2401 [NEW] -- Independence Hill Ret
        
    ### find relay IDs, create relay objects, then find relay settings
        for cb in cb_list:
            if len(cb) <= 4:
                cb_num = cb
            elif cb[:4].isnumeric() and cb[4] == ' ':
                cb_num = cb[:4]
            # elif cb[:3] == 'OCR':
            else:
                cb_num = cb[-5:-1]
            relay_name = str(sta.ID) + '_' + cb_num    
        
            breaker_url = 'http://sysprot.duke-energy.com/relay_info/decv2/aspen.asp?a=4&lid={0}&cid={1}&c=DEC'.format(sta.Aspen,cb.replace(' ','%20'))
            breaker_response = urllib.request.urlopen(breaker_url).read()
            breaker_tree = ET.fromstring(breaker_response)
            rels = breaker_tree.findall('RELAY')
            # rel = breaker_tree.find('RELAY')
            # relay = Relay(*[cb_num,cb]+[rel.attrib[x] for x in ['RELAYID','RELAYTYPE']])
            relay = Relay()
            
            for rel in rels:
                relay = Relay(*[cb_num,cb]+[rel.attrib[x] for x in ['RELAYID','RELAYTYPE']])
                
                if relay.Type in relay_settings: 
                    relay_url = 'http://sysprot.duke-energy.com/relay_info/decv2/aspen.asp?a=5&rid={0}&c=DEC'.format(relay.ID)
                    relay_response = urllib.request.urlopen(relay_url).read()
                    relay_tree = ET.fromstring(relay_response)
                    
                    try:
                        g1 = relay_tree.find('SETTINGS').find('REQUEST').find('GROUPS')[0] # group 1 settings
                    except:
                        try:
                            g1 = relay_tree.find('SETTINGS').find('REQUEST').find('GROUPS')[5] # group 6 settings
                        except:
                            sta.Relays[cb] = Relay(cb_num,cb)
                            continue
                    
                    for elem in list(g1):
                        a = elem.attrib
                        # if a['SETTINGNAME'] == relay_settings[relay.Type][0]:
                        #     relay.HLT_ph= tryfloat(a['SETTING'])
                        # elif a['SETTINGNAME'] == relay_settings[relay.Type][1]:
                        #     relay.HLT_g= tryfloat(a['SETTING'])
                        # elif a['SETTINGNAME'] == 'CTR':
                        #     relay.CTR = tryfloat(a['SETTING'])
                        for setname in relay_settings[relay.Type]:
                            if a['SETTINGNAME'] == setname:
                                setattr(relay, 'Setting_'+setname, a['SETTING'])
                        
                    break # if SEL relay, stop looking through relays (was probably the 1st one anyway)
                
                elif rel.attrib['S01'] == '79': # skips over CO-8, IAC, etc. 79 (?) relays
                    continue
                elif any(r_abbrev in relay.Type for r_abbrev in other_relays): # shortens IAC51, -53, OC-8 relay type strings
                    for r_abbrev in other_relays:
                        if r_abbrev in relay.Type:
                            relay.Type = r_abbrev
                    break # stop if relay type is found
                else:
                    continue
                    
            sta.Relays[cb] = relay
            relays[relay_name] = relay
    
    
    ### gather Short Circuit Data (impedances in p.u.) into Bank objects
    with open(source_impedance_file, mode = 'r') as infile:
        reader = csv.DictReader(infile, skipinitialspace=True)
        for row in reader:
            bus = row['Bus']
            if bus[-1] == 'r':
                continue # filters out entries ending in r - future replacements
            if bus[-1].isdigit():
                sub_bank = bus[-1]
            else:
                sub_bank = '1'
            stationtext = row['Location'][-4:]
            if not stationtext:
                continue # filters out non-distribution/industrial entries, e.g. Ameri Roller
            bank_eq = 'DEC_' + stationtext + '_' + sub_bank
            banks[bank_eq] = Bank(bank_eq,bus,tryfloat(row['KV']),tryfloat(row['R1_pu']),tryfloat(row['X1_pu']),
                tryfloat(row['R2_pu']),tryfloat(row['X2_pu']),tryfloat(row['R0_pu']),tryfloat(row['X0_pu']),
                tryfloat(row['LLL_amps']),tryfloat(row['LG_amps']))
    
    
    ## set impedances for Substation Equipment
    # print('Substation Equivalent Impedance Updates:')
    for bank in banks_loaded:
        if bank not in banks:
            continue
        b = banks[bank]
        sub = cympy.eq.GetEquipment(bank,cympy.enums.EquipmentType.Substation)
        kv = b.KV_nom
        kv_class = min(voltage_classes[op_region], key = lambda x:abs(x-kv))
        sub.SetValue(kv_class,'NominalKVLL')
        sub.SetValue(kv_class,'DesiredKVLL')
        sub.SetValue('PU','ImpedanceUnit')
        sub.SetValue(b.R1,'PositiveSequenceResistance')
        sub.SetValue(b.X1,'PositiveSequenceReactance')
        sub.SetValue(b.R0,'ZeroSequenceResistance')
        sub.SetValue(b.X0,'ZeroSequenceReactance')
        sub.SetValue(b.R2,'NegativeSequenceResistance')
        sub.SetValue(b.X2,'NegativeSequenceReactance')
    # print('    ' + ', '.join(banks[b].EquipmentID for b in banks) + ' equipment updated')
    
    
    
    
    cympy.app.ActivateRefresh(True)
    cympy.env.SetDataValidationOption(cympy.enums.DataValidationOption.AllowChangePrimaryKey, False)
    
    
    ### create Circuit objects
    for net in netlist:
        circuits[net] = Circuit(net, op_region)
        c = circuits[net]
        
    ### pair relay information to circuit object
        relay_name = str(c.StationID) + '_' + net[-4:]
        if relay_name in relays:
            c.Relay = relays[relay_name]
        else:
            c.Relay = Relay()
    
        
        ### output feeder relay settings
        with open(os.path.join(resultlocation, f'Feeder_Relay_Settings_{net}.csv'), 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Station Name',c.Substation])
            writer.writerow(['Circuit Number',c.name])
            
            if c.Relay.Type not in relay_settings:
                writer.writerow(['RID','UNKNOWN'])
                print('Feeder relay settings not found for ' + net)
            else:
                for setting in relay_settings[c.Relay.Type]:
                    writer.writerow([setting, str(getattr(c.Relay, f'Setting_{setting}'))])
    
    print(studyfile + ' - Feeder breaker relay settings saved to ' + resultlocation)
    
    
    if connect_dg:
        cympy.study.BeginMergeModifications()
        for dg in cympy.study.ListDevices(cympy.enums.DeviceType.ElectronicConverterGenerator):
            dg.SetValue('Connected','ConnectionStatus')
        cympy.study.EndMergeModifications('Connect all DG')
    elif disconnect_dg:
        cympy.study.BeginMergeModifications()
        for dg in cympy.study.ListDevices(cympy.enums.DeviceType.ElectronicConverterGenerator):
            dg.SetValue('Disconnected','ConnectionStatus')
        cympy.study.EndMergeModifications('Disconnect all DG')
    
    ### get device info, per network
    for net in netlist:
        ohs = [DevicePlus(x) for x in cympy.study.ListDevices(cympy.enums.DeviceType.OverheadByPhase,net)]
        ugs = [DevicePlus(x) for x in cympy.study.ListDevices(cympy.enums.DeviceType.Underground,net)]
        ecgs = [DevicePlus(x) for x in cympy.study.ListDevices(cympy.enums.DeviceType.ElectronicConverterGenerator,net)]
        breakers = [DevicePlus(x) for x in cympy.study.ListDevices(cympy.enums.DeviceType.Breaker,net)]
        fuses = [DevicePlus(x) for x in cympy.study.ListDevices(cympy.enums.DeviceType.Fuse,net)]
        reclosers = [DevicePlus(x) for x in cympy.study.ListDevices(cympy.enums.DeviceType.Recloser,net)]
        txs = [DevicePlus(x) for x in cympy.study.ListDevices(cympy.enums.DeviceType.SpotLoad,net)]
    
        devs = ohs + ugs + ecgs + breakers + fuses + reclosers
        
        # tx_upstream_sections = {tx.DeviceNumber : sections_upstream(tx) for tx in cympy.study.ListDevices(cympy.enums.DeviceType.SpotLoad, net)}
        
        # run short circuit simulation
        sc.SetValue(False,'ParametersConfigurations[0].AssumeLineTransposition')
        sc.SetValue('SC','ParametersConfigurations[0].Domain')
        sc.Execute('AnalysisNetworks.SelectedNetworks.Clear()')
        sc.Execute(f'AnalysisNetworks.SelectedNetworks.Add({net})')
        sc.Run()
        
        for dev in devs:
            dev.ShortCircuit_Query()
            dev.Zone_Query()
    
        for tx in txs:
            tx.ShortCircuit_Query()
            tx.Upstream_Query()
            tx.Update_Environment()
            tx.SpotLoad_Query()
        
        lf.Execute('AnalysisNetworks.SelectedNetworks.Clear()')
        lf.Execute(f'AnalysisNetworks.SelectedNetworks.Add({net})')
        lf.Run()
        for dev in devs:
            dev.Loading_Query()
        
        for dev in devs:
            dev.Downstream_Query()
        
        
        device_headers = [
            'Device Type',
            'Equipment Number',
            'Status',
            'Equipment ID',
            'Base Voltage (kVLL)',
            'Phase',
            'LLL / LLLG Max',
            'LL Max',
            'LL Min',
            'LG Max',
            'LG Min',
            'IA Load',
            'IB Load',
            'IC Load',
            'Protection Level',
            'Upstream Protective Device',
            'Conductor A',
            'Conductor B',
            'Conductor C',
            'Neutral Conductor',
            'Total downstream cKVA',
            'Downstream A PH cKVA',
            'Downstream B PH cKVA',
            'Downstream C PH cKVA',
            'Total downstream 3PH cKVA',
            'Summation of downstream 3PH cKVA >= 300KVA',
            'Largest downstream UG Spot Load (KVA)',
            'Max PH Fault Current at Largest downstream UG Spot Load',
            'Max GND Fault Current at Largest downstream UG Spot Load',
            'Largest downstream OH Spot Load (KVA)',
            'Max PH Fault Current at Largest downstream OH Spot Load',
            'Max GND Fault Current at Largest downstream OH Spot Load',
            'Downstream A PH Customers',
            'Downstream B PH Customers',
            'Downstream C PH Customers',
            'Total downstream Customers',
            'DER Generation kW',
            ]
        
        device_attribs = [
            'DeviceType_english',
            'DeviceNumber',
            'Status',
            'EquipmentID',
            'kVLL',
            'Phase',
            'LLL_max',
            'LL_max',
            'LL_min',
            'LG_max',
            'LG_min',
            'IA',
            'IB',
            'IC',
            'ProtectionLevel',
            'Parent',
            'Smallest_Conductor_A',
            'Smallest_Conductor_B',
            'Smallest_Conductor_C',
            'Smallest_Conductor_Neutral',
            'ckVA_total',
            'ckVAA',
            'ckVAB',
            'ckVAC',
            'ThreePhase_ckVA',
            'ThreePhase_Large_ckVA',
            'Largest_TX_UG_cKVA',
            'Largest_TX_UG_PhaseFault',
            'Largest_TX_UG_GroundFault',
            'Largest_TX_OH_cKVA',
            'Largest_TX_OH_PhaseFault',
            'Largest_TX_OH_GroundFault',
            'Customers_A',
            'Customers_B',
            'Customers_C',
            'Customers_Total',
            'DER_Gen_kW',
        ]
        
        dev_results_dict = {(d.DeviceNumber, d.DeviceType) : [getattr(d,a) for a in device_attribs] for d in devs}
    
        with open(os.path.join(resultlocation, f'CYME_Device_Report_{net}.csv'), 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(device_headers)
            writer.writerows(dev_results_dict.values())
        
    print(studyfile + ' - CYME device reports saved to ' + resultlocation)
    
    # overview file
    overview_fields = [
        'Station_Name',
        'Circuit_Number',
        'Base_Voltage',
        'R1',
        'X1',
        'R0',
        'X0',
        'R2',
        'X2',
        'Relay_Path',
        'CYME_Report_Path',
    ]
    
    for net in netlist:
        c = circuits[net]
        
        source = cympy.study.GetDevice(net, cympy.enums.DeviceType.Source)
        
        subname = c.Substation
        
        sub = cympy.eq.GetEquipment(source.EquipmentID,cympy.enums.EquipmentType.Substation)
        kv = sub.GetValue('NominalKVLL')
        r1 = sub.GetValue('PositiveSequenceResistance')
        x1 = sub.GetValue('PositiveSequenceReactance')
        r0 = sub.GetValue('ZeroSequenceResistance')
        x0 = sub.GetValue('ZeroSequenceReactance')
        r2 = sub.GetValue('NegativeSequenceResistance')
        x2 = sub.GetValue('NegativeSequenceReactance')
        
        overview_output = [subname, net, kv, r1, x1, r0, x0, r2, x2, os.path.join(resultlocation, f'Feeder_Relay_Settings_{net}.csv'), os.path.join(resultlocation, f'CYME_Device_Report_{net}.csv')]
        
        overview = zip(overview_fields, overview_output)
        
        
        with open(os.path.join(resultlocation, f'Feeder_Overview_{net}.csv'), 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(overview)
        print(studyfile + ' - Feeder overview reports saved to ' + resultlocation)
    
    cympy.study.Close(False)
    print(studyfile + ' complete')
print('All studies in ' + study_folder + ' complete')