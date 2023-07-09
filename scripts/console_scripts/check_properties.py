from data_tools.connection_init import acc, acu


walls = acc.GetElementsByType('CurtainWall')
thickness = acu.GetBuiltInPropertyId('General_Thickness')
element_id = acu.GetBuiltInPropertyId('General_ElementID')
values = acc.GetPropertyValuesOfElements(walls, [thickness])
ids = acc.GetPropertyValuesOfElements(walls, [element_id])
i = 0
for Propertyvalue in values:
    status = Propertyvalue.propertyValues[0].propertyValue.status
    id = ids[i].propertyValues[0].propertyValue.value
    if status == 'normal':
        value = Propertyvalue.propertyValues[0].propertyValue.value
        print("ID elementu: ", id," Wartość właściwości: ",value)
    else:
        value = "Nie określono"
    i += 1
