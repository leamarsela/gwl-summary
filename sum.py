import os
import xlrd
import xlwt
import json
import sys
import time

loc = os.getcwd()

project = loc[(loc.find('\\',12)+1):(loc.find('\\',12)+1+6)]


def idRev():
    dataRev = []
    for file in os.listdir(loc):
        if(file.endswith('.xls')) and (file[0:7] == 'Summary'):
            dataRev.append(file)
    
    if(len(dataRev) == 0):
        rev = 0
    else:
        rev = len(dataRev)
        
    return rev

pathSummary = loc + '\\Summary ' + project + ' rev ' + str(idRev()) + '.xls' 

locInvoice = os.path.dirname(os.path.dirname(loc)) + '\\Admin'
pathInvoice = locInvoice + '\\' + project + ' harga.xls' 


print('Start....')

    
obj = {}
data = {}
items = {}
classification = {}
txuu = {}
ds = {}
dsc = {}
oc = {}
consolidation = {}
fallingHead = {}
txcu = {}
uct = {}


for file in os.listdir(loc):
    if file.endswith(".xlsm"):
        nameId = file[(file.find('-'))+1:(file.find(')'))+1]
        
        if nameId not in data.keys():
            data[nameId] = {}
            data[nameId]['items'] = {}
        
        
        if file.endswith('Classification.xlsm'):
            if 'classification' not in data[nameId]['items'].keys():
                data[nameId]['items'].update({'classification': classification})
        
            wb = xlrd.open_workbook(os.path.join(loc, file))
            sh = wb.sheet_by_name('Report')
            density = wb.sheet_by_name('Unit Weight')
            hydro = wb.sheet_by_name('Hydrometer')
            a = {                   
                    'MC': sh.cell_value(15, 7),
                    'Normal Density': sh.cell_value(12, 7),
                    'Dry Density': density.cell_value(10, 1),
                    'Void Ratio': sh.cell_value(17, 16),
                    'Porosity': sh.cell_value(18, 16),
                    'Saturation': sh.cell_value(16, 16),
                    'Gs': sh.cell_value(14, 7),
                    'LL': sh.cell_value(17, 7),
                    'PL': sh.cell_value(16, 7),
                    'PI': sh.cell_value(18, 7),
                    'Gravel': hydro.cell_value(32, 1),
                    'Sand': hydro.cell_value(33, 1),
                    'Silt': hydro.cell_value(34, 1),
                    'Clay': hydro.cell_value(35, 1),
                    'Fines': sh.cell_value(19, 7)
            }
            data[nameId]['items']['classification'] = a   
 
        if file.endswith("TX UU.xlsm"):
            if 'txuu' not in data[nameId]['items'].keys():
                data[nameId]['items'].update({'txuu': txuu})
            
            wb = xlrd.open_workbook(os.path.join(loc, file))
            sh = wb.sheet_by_name("Report")
            a = {                    
                    'Cohession': sh.cell_value(18, 16),
                    'Phi': sh.cell_value(19, 16)
            }
            data[nameId]['items']['txuu'] = a
        
        if file.endswith("UCT.xlsm"):
        
            if 'uct' not in data[nameId]['items'].keys():
                data[nameId]['items'].update({'uct': uct})
                
            wb = xlrd.open_workbook(os.path.join(loc, file))
            sh = wb.sheet_by_name("Report")
            a = {
                    'Cohession': sh.cell_value(15, 15),
                    'St': sh.cell_value(14, 15),
                    'Ei': sh.cell_value(16, 15),
                    
            }
            data[nameId]['items']['uct'] = a
        
        if file.endswith("Consolidation.xlsm"):
            if 'consolidation' not in data[nameId]['items'].keys():
                data[nameId]['items'].update({'consolidation': consolidation})
            
            wb = xlrd.open_workbook(os.path.join(loc, file))
            sh = wb.sheet_by_name("Report")
            calc = wb.sheet_by_name('Calc')
            a = {                    
                    'eo': sh.cell_value(52, 7),
                    'Cc': sh.cell_value(53, 7),
                    'Cr': sh.cell_value(54, 7),
                    'Cs': sh.cell_value(55, 7),
                    'Pc': sh.cell_value(56, 7),
                    'SPotential': calc.cell_value(5, 11),
                    'SPressure': calc.cell_value(5, 12)
            }
            data[nameId]['items']['consolidation'] = a
        
        if file.endswith("TX CU.xlsm"):
            if 'txcu' not in data[nameId]['items'].keys():
                data[nameId]['items'].update({'txcu': txcu})
                
            wb = xlrd.open_workbook(os.path.join(loc, file))
            sh = wb.sheet_by_name("Report")
            a = {                    
                    'Ctotal': sh.cell_value(18, 14),
                    'Phitotal': sh.cell_value(19, 14),
                    'Ceffective': sh.cell_value(18, 16),
                    'Phieffective': sh.cell_value(19, 16)
            }
            data[nameId]['items']['txcu'] = a
        
        if file.endswith("DS.xlsm"):
            if 'ds' not in data[nameId]['items'].keys():
                data[nameId]['items'].update({'ds': ds})
            
            wb = xlrd.open_workbook(os.path.join(loc, file))
            sh = wb.sheet_by_name("Report")
            a = {                    
                    'Cohession': sh.cell_value(18, 16),
                    'Phi': sh.cell_value(19, 16)
            }
            data[nameId]['items']['ds'] = a              
        
        if file.endswith("DSC.xlsm"):
            if 'dsc' not in data[nameId]['items'].keys():
                data[nameId]['items'].update({'dsc': dsc})
                
            wb = xlrd.open_workbook(os.path.join(loc, file))
            sh = wb.sheet_by_name("Report")
            a = {
                    'Cohession': sh.cell_value(18, 16),
                    'Phi': sh.cell_value(19, 16)
            }
            data[nameId]['items']['dsc'] = a
                
        if file.endswith("Falling Head.xlsm"):
            if 'fallingHead' not in data[nameId]['items'].keys():
                data[nameId]['items'].update({'fallingHead': fallingHead})
            
            wb = xlrd.open_workbook(os.path.join(loc, file))
            sh = wb.sheet_by_name("calc")
            header = wb.sheet_by_name("Header")
            valDirection = header.cell_value(0, 2)
            direction = "Vertical" if valDirection == "Falling Head" else "Horizontal"
            a = {                    
                    'valK': sh.cell_value(17, 8),
                    'direction': direction
            }
            data[nameId]['items']['fallingHead'] = a   
        
        if file.endswith("OC.xlsm"):
            if 'oc' not in data[nameId]['items'].keys():
                data[nameId]['items'].update({'oc': oc})
            
            wb = xlrd.open_workbook(os.path.join(loc, file))
            sh = wb.sheet_by_name("Report")
            a = {                    
                    'oc': sh.cell_value(13, 6)
            }
            data[nameId]['items']['oc'] = a
        
        
        
        obj.update(data)
           
# print(json.dumps(obj, True, indent=4))


def runStatus():
    statusClassification = False 
    for it in (obj.keys()):
        if 'classification' in obj[it]['items']:
            statusClassification = True
                
    statusTxuu = False
    for it in (obj.keys()):
        if 'txuu' in obj[it]['items']:
            statusTxuu = True

    statusUct = False
    for it in (obj.keys()):
        if 'uct' in obj[it]['items']:
            statusUct = True

    statusConsolidation = False
    for it in (obj.keys()):
        if 'consolidation' in obj[it]['items']:
            statusConsolidation = True

    statusTxcu = False
    for it in (obj.keys()):
        if 'txcu' in obj[it]['items']:
            statusTxcu = True

    statusDs = False
    for it in (obj.keys()):
        if 'ds' in obj[it]['items']:
            statusDs = True

    statusDsc = False
    for it in (obj.keys()):
        if 'dsc' in obj[it]['items']:
            statusDsc = True

    statusFallingHead = False
    for it in (obj.keys()):
        if 'fallingHead' in obj[it]['items']:
            statusFallingHead = True
            
    statusOc = False
    for it in (obj.keys()):
        if 'oc' in obj[it]['items']:
            statusOc = True

    statusSc = False
    for it in (obj.keys()):
        if 'sc' in obj[it]['items']:
            statusSc = True
    
    return [statusClassification, statusTxuu, statusUct, statusConsolidation, statusTxcu, statusDs, statusDsc, statusFallingHead, statusOc, statusSc]

[statusClassification, statusTxuu, statusUct, statusConsolidation, statusTxcu, statusDs, statusDsc, statusFallingHead, statusOc, statusSc] = runStatus()

def groupHeader():
    Classification = [
        'MC (%)',
        'Normal Density (kN/m3)',
        'Dry Density (kN/m3)',
        'Void Ratio',
        'Porosity',
        'Saturation',
        'Gs',
        'LL',
        'PL',
        'PI',
        'Gravel',
        'Sand',
        'Silt',
        'Clay',
        'Fines'
    ]

    Triaxial_UU = [
        'c_uu (kPa)',
        'phi_uu (degree)'    
    ]

    Uct = [
        'c (kPa)',
        'Sensivity',
        'Ei (kPa)'
    ]

    Consolidation = [
        'eo',
        'Cc',
        'Cr',
        'Cs',
        'Pc (kPa)',
        'Swelling Potential (%)',
        'Swelling Pressure (kPa)'
    ]

    Triaxial_CU = [
        'c_cu (kPa)',
        'phi_cu (degree)',
        "c' (kPa)",
        "phi' (degree)"
    ]

    Direct_Shear = [
        'c_ds (kPa)',
        'phi_ds (degree)'    
    ]

    Direct_Shear_Consolidated = [
        'c_dsc (kPa)',
        'phi_dsc (degree)'    
    ]

    Falling_Head = [
        'k (m/s)',
        'direction'
    ]
    
    Organic_Content = [
        'Organic Content (%)'
    ]

    Sulphate_Content = [
        'sulphate content (%)'
    ]
    
    return [Classification, Triaxial_UU, Uct, Consolidation, Triaxial_CU, Direct_Shear, Direct_Shear_Consolidated, Falling_Head, Organic_Content, Sulphate_Content]

[Classification, Triaxial_UU, Uct, Consolidation, Triaxial_CU, Direct_Shear, Direct_Shear_Consolidated, Falling_Head, Organic_Content, Sulphate_Content] = groupHeader()

def listHeader():
    listHeader = []
    if(statusClassification):
        listHeader.append(Classification)

    if(statusTxuu):
        listHeader.append(Triaxial_UU)

    if(statusUct):
        listHeader.append(Uct)

    if(statusConsolidation):
        listHeader.append(Consolidation)

    if(statusTxcu):
        listHeader.append(Triaxial_CU)

    if(statusDs):
        listHeader.append(Direct_Shear)

    if(statusDsc):
        listHeader.append(Direct_Shear_Consolidated)

    if(statusFallingHead):
        listHeader.append(Falling_Head)

    if(statusOc):
        listHeader.append(Organic_Content)

    if(statusSc):
        listHeader.append(Sulphate_Content)
    
    return listHeader

listHeader = listHeader()


header = []
for h in listHeader:
    iHeader = [k for k, v in locals().items() if v == h]
    tempHeader = sorted(iHeader)
    header.append(tempHeader[0])

# print(header)

wbook = xlwt.Workbook()
     
style_text_wrap_font_bold_black_color = xlwt.Style.easyxf('font: bold on, color-index black; align: wrap on, horiz center')
style_centre = xlwt.Style.easyxf('align: wrap on, horiz center')

col_no_width = 128 * 15
col_width = 128 * 32
wsbook = wbook.add_sheet('Summary')


wsbook.write_merge(0, 1, 0, 0, 'No.', style_text_wrap_font_bold_black_color)
wsbook.write_merge(0, 1, 1, 1, 'Borehole', style_text_wrap_font_bold_black_color)


# # # # 'header summary
a = 0
b = 0
for val, i in enumerate(listHeader):
    
    b = b + len(i)
    wsbook.write_merge(0, 0, b - len(i) + 2, b + 1, header[val], style_text_wrap_font_bold_black_color)
    
    for itm in (i):
        wsbook.write(1, a + 2, itm, style_centre)   
        a+=1                       


# # # # 'fill summary 
for it, itVal in enumerate(sorted(obj.keys()), start=2):
    wsbook.write(it, 0, it-1, style_centre)
    wsbook.write(it, 1, itVal, style_centre)
    
    for itm, itmVal in (obj[itVal]['items'].items()):
        g = 0      
        for x, y in enumerate(listHeader):   
       
            for ix, iy in enumerate(y):                
                
                # 'classification'
                if(itm == 'classification'):
                    if(itmVal['MC'] and iy == 'MC (%)'):
                        wsbook.write(it, g+2, itmVal['MC'], style_centre)                        
                    
                    if(itmVal['Normal Density'] and iy == 'Normal Density (kN/m3)'):
                        wsbook.write(it, g+2, itmVal['Normal Density'], style_centre)                        
                    
                    if(itmVal['Dry Density'] and iy == 'Dry Density (kN/m3)'):
                        wsbook.write(it, g+2, itmVal['Dry Density'], style_centre)                        
                    
                    if(itmVal['Void Ratio'] and iy == 'Void Ratio'):
                        wsbook.write(it, g+2, itmVal['Void Ratio'], style_centre)                        
                    
                    if(itmVal['Porosity'] and iy == 'Porosity'):
                        wsbook.write(it, g+2, itmVal['Porosity'], style_centre)                        
                    
                    if(itmVal['Saturation'] and iy == 'Saturation'):
                        wsbook.write(it, g+2, itmVal['Saturation'], style_centre)                        
                    
                    if(itmVal['Gs'] and iy == 'Gs'):
                        wsbook.write(it, g+2, itmVal['Gs'], style_centre)                        
                                        
                    if(itmVal['LL'] and iy == 'LL'):
                        wsbook.write(it, g+2, itmVal['LL'], style_centre)                        
                
                    if(itmVal['PL'] and iy == 'PL'):
                        wsbook.write(it, g+2, itmVal['PL'], style_centre)                        
                    
                    if(itmVal['PI'] and iy == 'PI'):
                        wsbook.write(it, g+2, itmVal['PI'], style_centre)                        
                    
                    if(itmVal['Gravel'] and iy == 'Gravel'):
                        wsbook.write(it, g+2, itmVal['Gravel'], style_centre)                        
                    
                    if(itmVal['Sand'] and iy == 'Sand'):
                        wsbook.write(it, g+2, itmVal['Sand'], style_centre)                        
                    
                    if(itmVal['Silt'] and iy == 'Silt'):
                        wsbook.write(it, g+2, itmVal['Silt'], style_centre)                        
                    
                    if(itmVal['Clay'] and iy == 'Clay'):
                        wsbook.write(it, g+2, itmVal['Clay'], style_centre)                        
                    
                    if(itmVal['Fines'] and iy == 'Fines'):
                        wsbook.write(it, g+2, itmVal['Fines'], style_centre)                        
                                        
                # 'tx uu'        
                if(itm == 'txuu'):
                    if(itmVal['Cohession'] and iy == 'c_uu (kPa)'):
                        wsbook.write(it, g+2, itmVal['Cohession'], style_centre)                        
                    
                    if(itmVal['Phi'] and iy == 'phi_uu (degree)'):
                        wsbook.write(it, g+2, itmVal['Phi'], style_centre)
                
                # 'uct'
                if(itm == 'uct'):
                    if(itmVal['Cohession'] and iy == 'c (kPa)'):
                        wsbook.write(it, g+2, itmVal['Cohession'], style_centre)
                    
                    if(itmVal['St'] and iy == 'Sensivity'):
                        wsbook.write(it, g+2, itmVal['St'], style_centre)
                    
                    if(itmVal['Ei'] and iy == 'Ei (kPa)'):
                        wsbook.write(it, g+2, itmVal['Ei'], style_centre)
                
                # 'consolidation'
                if(itm == 'consolidation'):
                    if(itmVal['eo'] and iy == 'eo'):
                        wsbook.write(it, g+2, itmVal['eo'], style_centre)
                    
                    if(itmVal['Cc'] and iy == 'Cc'):
                        wsbook.write(it, g+2, itmVal['Cc'], style_centre)
                    
                    if(itmVal['Cr'] and iy == 'Cr'):
                        wsbook.write(it, g+2, itmVal['Cr'], style_centre)
                    
                    if(itmVal['Cs'] and iy == 'Cs'):
                        wsbook.write(it, g+2, itmVal['Cs'], style_centre)
                    
                    if(itmVal['Pc'] and iy == 'Pc (kPa)'):
                        wsbook.write(it, g+2, itmVal['Pc'], style_centre)                    
                                        
                    if(itmVal['SPotential'] and iy == 'Swelling Potential (%)'):
                        wsbook.write(it, g+2, itmVal['SPotential'], style_centre)
                            
                    if(itmVal['SPressure'] != '' and iy == 'Swelling Pressure (kPa)'):
                        wsbook.write(it, g+2, itmVal['SPressure'], style_centre)
                    
                # 'txcu'
                if(itm == 'txcu'):
                    if(itmVal['Ctotal'] and iy == 'c_cu (kPa)'):
                        wsbook.write(it, g+2, itmVal['Ctotal'], style_centre)                    
                    
                    if(itmVal['Phitotal'] and iy == 'phi_cu (degree)'):
                        wsbook.write(it, g+2, itmVal['Phitotal'], style_centre)                        
                    
                    if(itmVal['Ceffective'] and iy == "c' (kPa)"):
                        wsbook.write(it, g+2, itmVal['Ceffective'], style_centre)                        
                    
                    if(itmVal['Phieffective'] and iy == "phi' (degree)"):
                        wsbook.write(it, g+2, itmVal['Phieffective'], style_centre)
                    
                # 'ds'
                if(itm == 'ds'):
                    if(itmVal['Cohession'] and iy == 'c_ds (kPa)'):
                        wsbook.write(it, g+2, itmVal['Cohession'], style_centre)
                    
                    if(itmVal['Phi'] and iy == 'phi_ds (degree)'):
                        wsbook.write(it, g+2, itmVal['Phi'], style_centre)
                        
                # 'dsc'
                if(itm == 'dsc'):
                    if(itmVal['Cohession'] and iy == 'c_dsc (kPa)'):
                        wsbook.write(it, g+2, itmVal['Cohession'], style_centre)
                    
                    if(itmVal['Phi'] and iy == 'phi_dsc (degree)'):
                        wsbook.write(it, g+2, itmVal['Phi'], style_centre)
                
                # 'oc'
                if(itm == 'oc'):
                    if(itmVal['oc'] and iy == 'Organic Content (%)'):
                        wsbook.write(it, g+2, itmVal['oc'], style_centre)
                
                # 'fallingHead'
                if(itm == 'fallingHead'):
                    if(itmVal['valK'] and iy == 'k (m/s)'):
                        wsbook.write(it, g+2, itmVal['valK'], style_centre)
                    
                    if(itmVal['direction'] and iy == 'direction'):
                        wsbook.write(it, g+2, itmVal['direction'], style_centre)
                
                
                g+=1
            

# # # 'setting column
for i in range(0, g+2):
    if(i == 0):
        wsbook.col(i).width = col_no_width
    elif(i == 1):
        wsbook.col(i).width = 256 * (len(loc) + 10)
    else:
        wsbook.col(i).width = col_width
        
           
def sumItems(itm, iitm, gitm):
    itm = 0
    for it in (obj.keys()):
        for ita in (obj[it].keys()):
            for itb in (obj[it][ita].keys()):
                if(itb == str(gitm)):
                    for itc in (obj[it][ita][itb].keys()):                
                        if(itc == str(iitm)):
                            if(obj[it][ita][itb][itc] != '-' and obj[it][ita][itb][itc] != ''):
                                itm += 1
                            elif(obj[it][ita][itb][itc] == ''):
                                itm += 0
                            else:
                                itm += 0
    return itm              


sumMC = 0
sumDensity = 0
sumGs = 0
sumAtterberg = 0
sumTotalPsd = 0
sumSieve = 0
sumTxuu = 0
sumUct = 0
sumTxcu = 0
sumConsolidation = 0
sumConsolSwelling = 0
sumDs = 0
sumDsc = 0
sumFallingHead = 0
sumOC = 0

for it in (sorted(obj.keys())):
    for itm, itmVal in sorted(obj[it]['items'].items()):
               
        if(itm == 'classification'):
            # 'MC'
            sumMC = sumItems('sumMC', 'MC', itm)
            
            # 'Density'
            sumDensity = sumItems('sumDensity', 'Normal Density', itm)
            
            # 'Gs'
            sumGs = sumItems('sumGs', 'Gs', itm)
            
            # 'Atterberg'
            sumAtterberg = sumItems('sumAtterberg', 'PL', itm)            
            
            # 'totalPsd'
            sumTotalPsd = sumItems('sumTotalPsd', 'Fines', itm)             
            
            # 'hydro'
            sumHydro = sumItems('sumHydro', 'Clay', itm) 
            
            sumSieve = sumTotalPsd - sumHydro   
                       
        if(itm == 'txuu'):
            sumTxuu = sumItems('sumTxuu', 'Cohession', itm)
        
        if(itm == 'uct'):            
            sumUct = sumItems('sumUct', 'Cohession', itm)
        
        if(itm == 'consolidation'):
            sumTotalConsolidation = sumItems('sumConsolidation', 'eo', itm)
            
            sumConsolSwelling = sumItems('sumConsolSwelling', 'SPotential', itm)
            
            sumConsolidation = sumTotalConsolidation - sumConsolSwelling        
                
        if(itm == 'txcu'):
            sumTxcu = sumItems('sumTxcu', 'Ctotal', itm)
        
        if(itm == 'ds'):
            sumDs = sumItems('sumDs', 'Cohession', itm)
        
        if(itm == 'dsc'):
            sumDsc = sumItems('sumDsc', 'Cohession', itm)
        
        if(itm == 'fallingHead'):
            sumFallingHead = sumItems('sumFallingHead', 'valK', itm)
            
        if(itm == 'oc'):
            sumOC = sumItems('sumOC', 'oc', itm)
            


dataSummary = {
    'Moisture Content': sumMC,
    'Unit Weight': sumDensity,
    'Spesific Gravity': sumGs,
    'Atterberg Limit': sumAtterberg,
    'Particle Size Distribution': sumHydro,
    'Sieve': sumSieve,
    'Triaxial UU': sumTxuu,
    'UCT': sumUct,
    'Consolidation': sumConsolidation,
    'Consolidation Swelling': sumConsolSwelling, 
    'Triaxial CU': sumTxcu,
    'Direct Shear Soil': sumDs,
    'Consolidated Direct Shear': sumDsc,
    'Falling Head': sumFallingHead,
    'Organic Content': sumOC
}


wsSummary = wbook.add_sheet('Total')

wsSummary.col(1).width = 256 * (len(max(dataSummary, key=len)))

wsSummary.write(0, 0, 'No.', style_text_wrap_font_bold_black_color)
wsSummary.write(0, 1, 'Items', style_text_wrap_font_bold_black_color)
wsSummary.write(0, 2, 'Quantity', style_text_wrap_font_bold_black_color)

index = 0
for items, sumItems in sorted(dataSummary.items()):
    if(sumItems == 0):
        pass
    else:
        index+=1
        wsSummary.write(index, 0, index, style_centre)
        wsSummary.write(index, 1, items, style_centre)
        wsSummary.write(index, 2, sumItems, style_centre)   

# 'save file summary
wbook.save(pathSummary)
  


# 'Invoice
wbInvoice = xlwt.Workbook()
wsInvoice = wbInvoice.add_sheet('Sheet1')

wsInvoice.col(1).width = 256 * (len(max(dataSummary, key=len)))

wsInvoice.write(0, 0, 'No.', style_text_wrap_font_bold_black_color)
wsInvoice.write(0, 1, 'Items', style_text_wrap_font_bold_black_color)
wsInvoice.write(0, 2, 'Quantity', style_text_wrap_font_bold_black_color)
wsInvoice.write(0, 3, 'Price', style_text_wrap_font_bold_black_color)


index = 0
for items, sumItems in sorted(dataSummary.items()):
    if(sumItems == 0):
        pass
    else:
        index+=1
        wsInvoice.write(index, 0, index, style_centre)
        wsInvoice.write(index, 1, items, style_centre)
        wsInvoice.write(index, 2, sumItems, style_centre)
        wsInvoice.write(index, 3, 0, style_centre)

# 'save file invoice    
wbInvoice.save(pathInvoice)

print('Finish')