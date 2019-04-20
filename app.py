from extended_minidom import create_tag, create_var
from datetime import datetime


doc = create_tag()
print(doc)
project = create_tag('project', attributes={'xmlns': 'http://www.plcopen.org/xml/tc6_0200'})
fileHeader = create_tag('fileHeader', attributes={'companyName': '',
                                                  'productName': 'CODESYS',
                                                  'productVersion': 'CODESYS V3.5 SP7 Patch 4',
                                                  'creationDateTime': str(datetime.now())})
contentHeader = create_tag('contentHeader', attributes={'name': 'kp15_20190419.project',
                                                        'modificationDateTime': str(datetime.now())})

coordinateInfo = create_tag('coordinateInfo')
fbd = create_tag('fbd')
ld = create_tag('ld')
sfc = create_tag('sfc')
scaling_fbd = create_tag('scaling', attributes={'x': '1', 'y': '1'})
scaling_ld = create_tag('scaling', attributes={'x': '1', 'y': '1'})
scaling_sfc = create_tag('scaling', attributes={'x': '1', 'y': '1'})
addData_ProjectInformation = create_tag('addData')
data_ProjectInformation = create_tag('data',
                                     attributes={'name': 'http://www.3s-software.com/plcopenxml/projectinformation',
                                                 'handleUnknown': 'implementation'})
ProjectInformation = create_tag('ProjectInformation')
types = create_tag('types')
dataTypes = create_tag('dataTypes')
pous = create_tag('pous')
pou = create_tag('pou', attributes={'name': 'POU',
                                    'pouType': 'program'})
interface = create_tag('interface')
localVars = create_tag('localVars')

doc.appendChild(project)
project.appendChild(fileHeader)
project.appendChild(contentHeader)
contentHeader.appendChild(coordinateInfo)
coordinateInfo.appendChild(fbd)
fbd.appendChild(scaling_fbd)
coordinateInfo.appendChild(ld)
ld.appendChild(scaling_ld)
coordinateInfo.appendChild(sfc)
sfc.appendChild(scaling_sfc)
contentHeader.appendChild(addData_ProjectInformation)
addData_ProjectInformation.appendChild(data_ProjectInformation)
data_ProjectInformation.appendChild(ProjectInformation)
project.appendChild(types)
types.appendChild(dataTypes)
types.appendChild(pous)
pous.appendChild(pou)

pou.appendChild(interface)
interface.appendChild(localVars)
create_var(localVars, var_name='state', var_type='INT', var_simple_value='0')
create_var(localVars, var_name='i', var_type='WORD')


addData = create_tag('addData')
data = create_tag('data', attributes={'name': 'http://www.3s-software.com/plcopenxml/method',
                                      'handleUnknown': 'implementation'})
Method = create_tag('Method', attributes={'name': 'IEC104_Init', 'ObjectId': 'IEC104_Init'})
interface = create_tag('interface')
body = create_tag('body')
ST = create_tag('ST')
xhtml = create_tag('xhtml', '//Настройки Slave IEC104', attributes={'xmlns': 'http://www.w3.org/1999/xhtml'})

xml_str = doc.toprettyxml(indent="  ", encoding='utf-8')

with open("minidom_example2.xml", "wb") as f:
    f.write(xml_str)
