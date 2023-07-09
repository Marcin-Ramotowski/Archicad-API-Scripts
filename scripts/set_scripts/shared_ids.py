from data_tools.connection_init import acc, act, acu

elements = acc.GetAllElements()
typeId = acu.GetBuiltInPropertyId('General_Type')
types = acc.GetPropertyValuesOfElements(elements, [typeId])
elementID = acu.GetBuiltInPropertyId('General_ElementID')
elementIds = acc.GetPropertyValuesOfElements(elements, [elementID])

boundingBoxes = acc.Get2DBoundingBoxes(elements)
heightId = acu.GetBuiltInPropertyId('General_Height')
heights = acc.GetPropertyValuesOfElements(elements, [heightId])
dimensions = []
checked = []
setOfIds = []
indexes_occurrences = []
ElementPropertyValues = []
i = 0
for box in boundingBoxes:
    elem_type = types[i].propertyValues[0].propertyValue.value
    if elem_type == 'Kształt' or elem_type == 'Ściana':  # strefa wyjątków
        text = f"Off{i}"
        dimensions.append(text)
        i += 1
        continue
    xMin = box.boundingBox2D.xMin
    xMax = box.boundingBox2D.xMax
    yMin = box.boundingBox2D.yMin
    yMax = box.boundingBox2D.yMax
    width = round(xMax - xMin, 3)
    length = round(yMax - yMin, 3)
    height = round(heights[i].propertyValues[0].propertyValue.value, 3)
    text = f"{length}m x {width}m x {height}m"
    dimensions.append(text)
    i += 1
i = j = 0
k = 1
check = False
for dimension in dimensions:
    amount = dimensions.count(dimension)
    if amount > 1:
        element = elementIds[i].propertyValues[0].propertyValue.value
        setOfIds.append(element)
        if len(setOfIds) > 0:
            if check is False:
                print("Istnieją w projekcie obiekty o identycznych wymiarach.")
                check = True
        if dimension not in checked:
            for x in dimensions:
                if x == dimension:
                    indexes_occurrences.append(j)
                j += 1
            j = 0
            first = indexes_occurrences[0]
            newID = elementIds[first].propertyValues[0].propertyValue.value
            while k < len(indexes_occurrences):
                change = indexes_occurrences[k]
                toChange = elements[change].elementId
                propertyValue = act.NormalStringPropertyValue(newID, 'string', 'normal')
                x = act.ElementPropertyValue(toChange, elementID, propertyValue)
                ElementPropertyValues.append(x)
                k += 1
            k = 1
            indexes_occurrences = []
        checked.append(dimension)
    i += 1
if len(ElementPropertyValues) > 0:
    y = acc.SetPropertyValuesOfElements(ElementPropertyValues)
    if y[0].success is True:
        print('Operacja zakończyła się powodzeniem.')
    else:
        print(f'Operacja zakończona niepowodzeniem.\n{y[0].error}')
