from data_tools.connection_init import acc, act, acu


elementIdBuiltInPropertyUserId = act.BuiltInPropertyUserId('General_ElementID')
elementIdPropertyId = acc.GetPropertyIds([elementIdBuiltInPropertyUserId])[0].propertyId
elementIds = acc.GetAllElements()

propertyValuesForElements = acc.GetPropertyValuesOfElements(elementIds, [elementIdPropertyId])

elementIdPropertyValues = set()
conflictedIds= set()
for propertyValuesForElement in propertyValuesForElements:
    elementIdPropertyValue = propertyValuesForElement.propertyValues[0].propertyValue.value
    if elementIdPropertyValue in elementIdPropertyValues:
        conflictedIds.add(elementIdPropertyValue)
    elementIdPropertyValues.add(elementIdPropertyValue)

x=len(conflictedIds)
if x == 0:
    print('Żadne ID nie powtarza się.')
else:
    print('Niektóre ID powtarzają się. Lista powtarzających się ID:')
    for ID in conflictedIds:
        print(ID)