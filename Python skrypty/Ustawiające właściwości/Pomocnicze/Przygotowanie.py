from archicad import ACConnection
conn = ACConnection.connect()
assert conn

acc = conn.commands
act = conn.types
acu = conn.utilities

zones = acc.GetElementsByType('Zone')
propertyId1 = acu.GetBuiltInPropertyId('Zone_ZoneNumber')
propertyId2 = acu.GetBuiltInPropertyId('General_ElementID')
zonesIds = acc.GetPropertyValuesOfElements(zones, [propertyId2])
elemPropertyValues = []
i = 0
for zone in zonesIds:
    id = zone.propertyValues[0].propertyValue.value
    elementId = zones[i].elementId
    propertyValue = act.NormalStringPropertyValue(id, 'string', 'normal')
    value = act.ElementPropertyValue(elementId, propertyId1, propertyValue)
    elemPropertyValues.append(value)
    i += 1
y = acc.SetPropertyValuesOfElements(elemPropertyValues)
if y[0].success == True:
    print('Operacja zakończyła się powodzeniem.')
else:
    print(f'Operacja zakończona niepowodzeniem.\n{y[0].error}')