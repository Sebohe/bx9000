import ezdxf
from itertools import groupby
import csv
from dbfread import DBF
import easygui
import sys
import traceback
from gui_auto import export_dxf, export_pdf
import os
import time

def determineBorderAreas (DWG):
    #This function determines the area of a border. returns a 2x2 matrix.
    """
    Dimensions of Border:
    """
    height = 21
    width = 33
    areasList = []
    for e in DWG.entities:
        if e.dxf.layer == "BORDER" and e.dxftype() == 'INSERT':
            #This exludes anything that is before the central meridian. (0=X)
            if e.dxf.insert[0] > -1:
                areasList.append(((e.dxf.insert[0],e.dxf.insert[0]+width),
                                    (e.dxf.insert[1],e.dxf.insert[1]+height)))

    return sorted(areasList)

def parseEntityMText(mText):

    mText = mText.replace("{","")
    mText = mText.replace("}","")
    #Example case of helvetica input:
    #{\Fhelvc1a.shx;FOR USE ON A SOLIDLY\PGROUNDED WYE SOURCE ONLY.}
    mText = mText.replace("\\Fhelvc1a.shx;","")
    mText = mText.replace(",\\P", ",")

    #change this to "in"
    if "\\P" in mText:
        return mText.split("\\P")
    else:
        return mText

def sheetToItemNumber(sheet, index):
    inTranslate ="ABCDEFGHI"
    outTranslate="123456789"
    index = index if index <= 9 else 9

    pageTypeIndex = sheet[1] if int(sheet[1]) <= 9 else '9'
    translateTable = str.maketrans(inTranslate, outTranslate)
    return sheet[0].translate(translateTable)+ '0' + pageTypeIndex+ str(index)

def breakApartMultipleParts(primitiveDictionary):

    partno = ''
    itemno = ''
    quantity = '1'
    refdesmemo = ''

    sections = primitiveDictionary['PARTNO'].split(",")
    for section in sections:
        #PARTNO
        if "5:" in section:
            partno = section.split(":")[1].strip()
        #ITEMNO
        elif "6:" in section:
            itemno = section.split(":")[1].strip()
        #QTY
        elif "7:" in section:
            quantity = section.split(":")[1].strip()
        #REFDESMEMO
        elif "19:" in section:
            refdesmemo = section.split(":")[1].strip()

    return {'Position':primitiveDictionary['Position'],
                    'BomType':primitiveDictionary['BomType'],
                    'PARTNO':partno,
                    'Sheet':primitiveDictionary['Sheet'],
                    'REFDESMEMO':refdesmemo,
                    'QTY':quantity,
                    'ITEMNO':itemno}

def insideOfArea(area, coordinate):
    #returns true or false if coordinate is inside of area
    x = False
    y = False

    if (area[0][0] < coordinate[0] and coordinate[0] < area[0][1]):
        x = True
    if (area[1][0] < coordinate[1] and coordinate[1] < area[1][1]):
        y = True

    return True if x and y else False

def release(f_name):

    DWG = ezdxf.readfile(f_name)

    for e in DWG.entities:

        rightBom = True if (e.dxf.layer == "BOM_CORDS"
                        or e.dxf.layer == "BOM_PANEL"
                        or e.dxf.layer == "BOM_WIRE") else False


        if rightBom:
            if e.dxftype() == "TEXT":
                if ":" in e.dxf.text:
                    sections = e.dxf.text.split(",")
                    text = ''
                    for section in sections:
                        if "5:" in section:
                            text += section.split(":")[1].strip()
                        #QTY
                        if "7:" in section:
                            text += section.replace("7:", " QTY:")

                    e.dxf.text = text

            if e.dxftype() == "MTEXT":
                if ":" in e.get_text():

                    mTextTemporary = ''
                    parts = parseEntityMText(e.get_text())

                    if type(parts) == type([]):
                        text = ''
                        for part in parts:
                            sections = part.split(",")

                            for section in sections:
                                if "5:" in section:
                                    text += section.split(":")[1].strip()
                                #QTY
                                if "7:" in section:
                                    text += section.replace("7:", " QTY:")

                                if ":" not in section:
                                    text += section


                            text += "\\P"
                        mTextTemporary += text

                    elif type(parts) == type(""):
                        text = ''
                        sections = parts.split(",")
                        for section in sections:
                            if "5:" in section:
                                text += section.split(":")[1].strip()
                            #QTY
                            if "7:" in section:
                                text += section.replace("7:", " QTY: ")

                            if ":" not in section:
                                text += section

                        mTextTemporary = text



                    e.set_text(mTextTemporary)

    DWG.saveas(f_name.replace(".dxf", ' RELEASED.dxf'))
    return(f_name.replace(".dxf", ' RELEASED.dxf'))

def scanner(fileX):

    DWG = ezdxf.readfile(fileX)

    cordsEntities = []
    panelEntities = []
    wireEntities = []
    sheetEntitites = []
    sheet_numberEntites = []
    panelBOMPrimitives = []
    cordBOMPrimitives = []
    wireBOMPrimitives = []

    sheetDictList = []
    structure = [[panelEntities,panelBOMPrimitives],
                [cordsEntities,cordBOMPrimitives],
                [wireEntities,wireBOMPrimitives]]

    #create the lists of all the entities we are interested in.
    for e in DWG.entities:

        if e.dxftype() == "MTEXT":
            if e.dxf.layer == "BOM_CORDS":
                cordsEntities.append(e)
            elif e.dxf.layer == "BOM_PANEL":
                panelEntities.append(e)
            elif e.dxf.layer == "BOM_WIRE":
                wireEntities.append(e)

        elif e.dxftype() == "TEXT":
            if "BOMNO" in e.dxf.text:
                BOMNO = e.dxf.text.split(":")[1].strip()
            elif "BOMDESCRI" in e.dxf.text:
                BOMDESCRI = e.dxf.text.split(":")[1].strip()
            #captures the sheet numbers entitites
            elif e.dxf.layer == "SHEETS":
                sheetEntitites.append(e)
            elif e.dxf.layer == "BOM_CORDS":
                cordsEntities.append(e)
            elif e.dxf.layer == "BOM_PANEL":
                panelEntities.append(e)
            elif e.dxf.layer == "BOM_WIRE":
                wireEntities.append(e)

            elif e.dxf.layer == "SHEET_NUMBERS":
                sheet_numberEntites.append(e)

    #this block creates the sheet Dict
    for area in determineBorderAreas(DWG):
        for e in sheetEntitites:
            #.get_pos() structure: ('LEFT', (29.75288587998724, 66.22258033165352, 0.0), None)
            if insideOfArea(area, e.get_pos()[1]):
                sheetDictList.append({'Area':area, 'PARTNO':e.dxf.text})
                break

    #the next line for loop obtains all of the primitves for the different BOMs
    for k, entityListIterator in enumerate(structure):

        for ent in entityListIterator[0]:
            if ent.dxftype() == "TEXT":
                position = ent.get_pos()[1]
                text = ent.dxf.text

            elif ent.dxftype() == "MTEXT":
                position = ent.dxf.insert
                text = ent.get_text()

            for i, sheet in enumerate(sheetDictList):
                if insideOfArea(sheet['Area'], position):
                    entityListIterator[1].append({'Position':position,
                                                    'BomType':ent.dxf.layer,
                                                    'PARTNO':parseEntityMText(text),
                                                    'Sheet':sheet['PARTNO']})
                    break
                #captures parts that are outside the borders. Like the parts TABLE
                elif i == len(sheetDictList)-1:
                    entityListIterator[1].append({'Position':position,
                                                    'BomType':ent.dxf.layer,
                                                    'PARTNO':parseEntityMText(text),
                                                    'Sheet':""})

        structure[k][1] = entityListIterator[1]
    #breaks apart lists to another primitive dictionary member.
    for k, primitiveDictIterator in enumerate(structure):
        tempList_1 = []
        for primitiveItemX in primitiveDictIterator[1]:
            if type(primitiveItemX['PARTNO']) == type([]):
                for partnoK in (primitiveItemX['PARTNO']):
                    tempList_1.append({'Position': primitiveItemX['Position'],
                    'BomType': primitiveItemX['BomType'],
                    'PARTNO': partnoK, 'Sheet': primitiveItemX['Sheet']})

            else:
                tempList_1.append(primitiveItemX)

        structure[k][1] = tempList_1

    bomGroup = []
    for member in structure:
        bomGroup.append(member[1])

    for k, bomX in enumerate(bomGroup):
        tempList_1 = []
        for item in sorted(bomX, key=lambda x: x['Sheet']):
            if ":" in item['PARTNO']:
                tempList_1.append(breakApartMultipleParts(item))
            else:
                item['REFDESMEMO'] = ''
                item['QTY'] = '1'
                item['ITEMNO'] = ''
                tempList_1.append(item)

        bomGroup[k] = (tempList_1)

    """
    ++++++++++++++++++++++++++++
    PANEL
    """
    #Adds the quantities
    PANELBOMLIST = []
    tempPanelList = sorted(bomGroup[0], key=lambda x: x['PARTNO'])

    for key, vals in groupby(tempPanelList, lambda x: x['PARTNO'].strip()):
        vals = sorted(list(vals), key=lambda x: x['Sheet'])
        quantity = 0.0
        for part in vals:
            quantity = float(part['QTY'])+quantity

        vals[0]['QTY'] = quantity
        PANELBOMLIST.append(vals[0])

    #Generates the ITEMNO and BOMNO AND BOMDESCRIPTION
    itemno = ''
    tempPanelList = sorted(PANELBOMLIST, key=lambda x: x['Sheet'])
    PANELBOMLIST = []
    for key, vals in groupby(tempPanelList, lambda x: x['Sheet'].strip()):
        vals = sorted(list(vals), key=lambda x: x['Position'])
        for index, part in enumerate(vals):
            if part['ITEMNO'] == '' and part['Sheet'] != '':
                itemno = sheetToItemNumber(part['Sheet'], index)
            else:
                itemno = part['ITEMNO']

            part['ITEMNO']=itemno
            part['BOMNO']=BOMNO
            part['BOMDESCRI']=BOMDESCRI
            PANELBOMLIST.append(part)

    bomGroup[0] = sorted(PANELBOMLIST, key=lambda x: x['ITEMNO'])
    """
    -------------------------------
    PANEL
    +++++++++++++++++++++++++++++++
    """


    """
    +++++++++++++++++++++++++++++++
    CORDS
    ------------------------------
    """
    #Generates the ITEMNO and BOMNO AND BOMDESCRIPTION
    itemno = ''
    tempCordsList = sorted(bomGroup[1], key=lambda x: x['Sheet'])
    tempList_1 = []
    for key, vals in groupby(tempCordsList, lambda x: x['Sheet'].strip()):
        vals = sorted(list(vals), key=lambda x: x['Position'])
        for index, part in enumerate(vals):
            if part['ITEMNO'] == '':
                itemno = sheetToItemNumber(part['Sheet'], index)
            else:
                itemno = part['ITEMNO']

            part['ITEMNO']=itemno
            part['BOMNO']=BOMNO
            part['BOMDESCRI']=BOMDESCRI
            tempList_1.append(part)

    bomGroup[1] = sorted(tempList_1, key=lambda x: x['PARTNO'])
    tempList_1 = []
    for key, vals in groupby(bomGroup[1], lambda x: x['PARTNO']):
        #This list is needed to return a list object. not a grouper object.
        vals = (list(vals))
        quantity = 0.0
        tempList_2 = []
        for part in vals:
            if part['REFDESMEMO'] == '':
                quantity = float(part['QTY'])+quantity
            else:
                tempList_2.append(part)

        if vals[0]['REFDESMEMO'] == '':
            vals[0]['QTY'] = quantity
            tempList_1.append(vals[0])
        else:
            tempList_1.extend(tempList_2)

    bomGroup[1] = sorted(tempList_1, key=lambda x: x['ITEMNO'])

    """
    -------------------------------
    CORDS
    +++++++++++++++++++++++++++++++
    """


    """
    ++++++++++++++++++++++++++++
    WIRE
    ----------------------------
    """
    #Generates the ITEMNO and BOMNO AND BOMDESCRIPTION
    itemno = ''
    tempWireList = sorted(bomGroup[2], key=lambda x: x['Sheet'])
    tempList_1 = []
    for key, vals in groupby(tempWireList, lambda x: x['Sheet'].strip()):
        vals = sorted(list(vals), key=lambda x: x['Position'])
        for index, part in enumerate(vals):
            if part['ITEMNO'] == '':
                itemno = sheetToItemNumber(part['Sheet'], index)
            else:
                itemno = part['ITEMNO']

            part['ITEMNO']=itemno
            part['BOMNO']=BOMNO
            part['BOMDESCRI']=BOMDESCRI
            tempList_1.append(part)

    bomGroup[2] = sorted(tempList_1, key=lambda x: x['ITEMNO'])

    """
    -------------------------------
    WIRE
    +++++++++++++++++++++++++++++++
    """

    """
    ++++++++++++++++++++++++++++++
    MRP
    -------------------------------
    """

    MRPPARTpath = [r'\\toptier-mrp\pcmrpw ver. 8.0\MRPPART.DBF',
                    r'\\toptier-mrp\pcmrpw ver. 8.0\MRPPART.FPT']

    PART_TABLE = DBF(MRPPARTpath[0], load=True)

    for j, bomX in enumerate(bomGroup):
        tempList_1 = []
        for dictX in bomX:
            for k, record in enumerate(PART_TABLE):
                if record['PARTNO'] in dictX['PARTNO']:

                    tempDict = dictX
                    tempDict['PART_ASSY'] = record['PART_ASSY']
                    tempDict['DESCRIPT'] = record['DESCRIPT']
                    tempList_1.append(tempDict)
                    break

                if k == len(PART_TABLE)-1:
                    if "LABOR" in dictX['PARTNO']:
                        tempDict = dictX
                        tempDict['PART_ASSY'] = 'L'
                        tempDict['DESCRIPT'] = record['DESCRIPT']
                        tempList_1.append(tempDict)
                        break
                    elif "HW" in dictX['PARTNO']:
                        tempDict = dictX
                        tempDict['PART_ASSY'] = 'P'
                        tempDict['DESCRIPT'] = record['DESCRIPT']
                        tempList_1.append(tempDict)
                        break

        bomGroup[j] = tempList_1
    """
    -------------------------------
    MRP
    ++++++++++++++++++++++++++++++
    """


    try:
        csv_name = fileX.split(".")[0]+"."+fileX.split(".")[1]+" "+BOMDESCRI+'.csv'
    except:
        csv_name = fileX.split(".")[0]+BOMDESCRI+'.csv'

    with open(csv_name, 'w+') as csvfile:
        writer = csv.writer(csvfile, lineterminator='\n')
        fieldnames = ['BOMNO', 'BOMDESCRI', 'PARTNO', 'ITEMNO', 'QTY',
        'PART_ASSY', 'REFDESMEMO', 'PARTDESC']
        writer.writerow(fieldnames)

        dictWriter = csv.DictWriter(csvfile, fieldnames=fieldnames, lineterminator='\n')

        for bomList in bomGroup:
            if bomList[0]['BomType'] == 'BOM_PANEL':
                bomType = ' '
            elif bomList[0]['BomType'] == 'BOM_CORDS':
                bomType = 'CORDS'
            elif bomList[0]['BomType'] == 'BOM_WIRE':
                bomType = 'WIRE'

            for item in bomList:

                item['PARTDESC']= ''
                item = {'BOMNO':item['BOMNO']+bomType[0].strip(),
                        'BOMDESCRI':item['BOMDESCRI']+' '+bomType,
                        'PARTNO':item['PARTNO'],
                        'ITEMNO':item['ITEMNO'],
                        'QTY':item['QTY'],
                        'PART_ASSY':item['PART_ASSY'],
                        'REFDESMEMO':item['REFDESMEMO'],
                        'PARTDESC':item['DESCRIPT']}
                dictWriter.writerow(item)




    #def determine_last_sheet
    last_sheet = sheetDictList[-1]
    for sheetX in sheetDictList:
        if sheetX['Area'][0][0] > last_sheet['Area'][0][0] and sheetX['Area'][1][0] <  last_sheet['Area'][1][0]:
            last_sheet = sheetX


    for number_ent in sheet_numberEntites:
        position = number_ent.get_pos()[1]
        if insideOfArea(last_sheet['Area'], position):
                sheets_quantities = int(number_ent.dxf.text)

    return sheets_quantities


if __name__=="__main__":

    try:
        path = os.path.join("T:", "JOBS", "2017")
        fileX = easygui.fileopenbox(msg="Select the DWG you want to relase.", default=path)

        if fileX:
            export_dxf("\"" + fileX + "\"")
            dxf = fileX.replace(".dwg", ".dxf")
            sheets_quantities = scanner(dxf)
            released_file = release(dxf)
            export_pdf("\"" + released_file + "\"", sheets_quantities)

            time.sleep(2)
            os.remove(released_file)

            os.remove(dxf)
            bak = fileX.replace(".dwg", ".bak")
            os.remove(bak)


            easygui.msgbox('Done releasing file and generating BOM')
        else:
            raise "File not found."

    except:
        easygui.msgbox(title="ERROR!", msg=str(traceback.format_exc()))
        sys.exit(-1)


    sys.exit()
