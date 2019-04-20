from extended_minidom import create_tag

doc = create_tag()

root = create_tag('root')
doc.appendChild(root)

addData = create_tag('addData')
data = create_tag('data', '', {'name': 'http://www.3s-software.com/plcopenxml/method', 'handleUnknown': 'implementation'})
Method = create_tag('Method', '', {'name': 'IEC104_Init', 'ObjectId': 'IEC104_Init'})
interface = create_tag('interface')
body = create_tag('body')
ST = create_tag('ST')
xhtml = create_tag('xhtml', '//Настройки Slave IEC104', {'xmlns': 'http://www.w3.org/1999/xhtml'})
root.appendChild(addData)
addData.appendChild(data)
data.appendChild(Method)
Method.appendChild(interface)
Method.appendChild(body)
body.appendChild(ST)
ST.appendChild(xhtml)



xml_str = doc.toprettyxml(indent="  ")
with open("minidom_example2.xml", "w") as f:
    f.write(xml_str)