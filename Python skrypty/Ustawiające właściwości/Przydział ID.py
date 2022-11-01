from archicad import ACConnection
from Pakiet import funkcje
conn = ACConnection.connect()
assert conn

acc = conn.commands
act = conn.types
acu = conn.utilities

allProperties = acc.GetAllPropertyNames()
neededProperties = (['Do skryptów', 'Położenie'], [])
run = funkcje.is_all_properties_available(allProperties, neededProperties)

if run:
    elements = acc.GetAllElements()
    propertyId = acu.GetBuiltInPropertyId('General_ElementID')
    propertyId2 = acu.GetUserDefinedPropertyId('Do skryptów', 'Położenie')
    propertyId3 = acu.GetBuiltInPropertyId('General_Type')
    positions = acc.GetPropertyValuesOfElements(elements, [propertyId2])
    type_raw = acc.GetPropertyValuesOfElements(elements, [propertyId3])
    ElementPropertyValues = []
    a = b = c = d = e = f = g = h = i = j = k = l = m = n = o = p = 0

    for type in type_raw:
        Type = type.propertyValues[0].propertyValue.value
        elementId = elements[a].elementId
        if Type =='Ściana':
            status = positions[a].propertyValues[0].propertyValue.status
            if status =='normal':
                position =positions[a].propertyValues[0].propertyValue.value
                if position =='Zewnętrze':
                    propertyValue = act.NormalStringPropertyValue(f'SZ{b}', 'string', 'normal')
                    b+=1
                elif position =='Wnętrze':
                    propertyValue = act.NormalStringPropertyValue(f'SW{c}', 'string', 'normal')
                    c+=1
                else:
                    print('Nie można określić położenia ściany.')
                    a+=1
                    continue
            else:
                print('Nie można określić położenia ściany.')
                a+=1
                continue
        elif Type == 'Słup':
            propertyValue = act.NormalStringPropertyValue(f'S{d}', 'string', 'normal')
            d+=1
        elif Type == 'Przegroda strukturalna':
            propertyValue = act.NormalStringPropertyValue(f'SK{e}', 'string', 'normal')
            e+=1
        elif Type == 'Strop':
            propertyValue = act.NormalStringPropertyValue(f'P{f}', 'string', 'normal')
            f+=1
        elif Type == 'Belka':
            propertyValue = act.NormalStringPropertyValue(f'B{g}', 'string', 'normal')
            g+=1
        elif Type == 'Dach':
            propertyValue = act.NormalStringPropertyValue(f'DA{h}', 'string', 'normal')
            h+=1
        elif Type == 'Drzwi':
            propertyValue = act.NormalStringPropertyValue(f'D{i}', 'string', 'normal')
            i+=1
        elif Type == 'Okno':
            propertyValue = act.NormalStringPropertyValue(f'O{j}', 'string', 'normal')
            j+=1
        elif Type =='Schody':
            propertyValue = act.NormalStringPropertyValue(f'SCH{k}', 'string', 'normal')
            k+=1
        elif Type == 'Obiekt':
            propertyValue = act.NormalStringPropertyValue(f'OB{l}', 'string', 'normal')
            l+=1
        elif Type =='Kształt':
            propertyValue = act.NormalStringPropertyValue(f'KS{m}', 'string', 'normal')
            m+=1
        elif Type == 'Lampa':
            propertyValue = act.NormalStringPropertyValue(f'LP{n}', 'string', 'normal')
            n+=1
        elif Type == 'Balustrada':
            propertyValue = act.NormalStringPropertyValue(f'BL{o}', 'string', 'normal')
            o+=1
        elif Type =='Strefa':
            propertyValue = act.NormalStringPropertyValue(f'POM{p}', 'string', 'normal')
            p+=1
        else:
            a+=1
            continue
        x = act.ElementPropertyValue(elementId, propertyId, propertyValue)
        ElementPropertyValues.append(x)
    y = acc.SetPropertyValuesOfElements(ElementPropertyValues)
    if y[0].success == True:
        print('Operacja zakończyła się powodzeniem.')
        i += 1
    else:
        print(f'Operacja zakończona niepowodzeniem.\n{y[0].error}')
    a+=1
