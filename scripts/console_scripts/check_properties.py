from data_tools.connection_init import acc, act, acu


walls = acc.GetElementsByType('CurtainWall')
a = acu.GetBuiltInPropertyId('General_Thickness')
b = acu.GetBuiltInPropertyId('General_ElementID')
values = acc.GetPropertyValuesOfElements(walls, [a])
ids = acc.GetPropertyValuesOfElements(walls, [b])
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
