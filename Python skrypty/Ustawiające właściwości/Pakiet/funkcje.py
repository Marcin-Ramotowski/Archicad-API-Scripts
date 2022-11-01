
def is_inside(polygon, point):
    """ Sprawdza, czy dany punkt leży wewnątrz danej figury."""

    def cross(p1, p2):
        x1, y1 = p1
        x2, y2 = p2

        # horizontal line
        if y1 - y2 == 0:
            if y1 == point[1]:
                if min(x1, x2) <= point[0] <= max(x1, x2):
                    return 1, True
            return 0, False

        # vertical line
        if x1 - x2 == 0:
            if min(y1, y2) <= point[1] <= max(y1, y2):
                if point[0] <= max(x1, x2):
                    return 1, point[0] == max(x1, x2)
            return 0, False

        # diagonal line
        a = (y1 - y2) / (x1 - x2)
        b = y1 - x1 * a
        x = (point[1] - b) / a
        if point[0] <= x:
            if min(y1, y2) <= point[1] <= max(y1, y2):
                return 1, point[0] == x or point[1] in (y1, y2)
        return 0, False

    # MAIN
    cross_points = 0
    for x in range(len(polygon)):
        num, on_line = cross(polygon[x], polygon[x-1])
        if on_line:
            return True
        cross_points += num

    return False if cross_points % 2 == 0 else True


def second_index(text: str, symbol: str):
    """ Zwraca indeks drugiego wystąpienia danego elementu licząc od lewej strony"""
    num = text.find(symbol, text.find(symbol)+1)
    return num if num > -1 else None


def przerabiarka(rawCoordinates):
    """ Zwraca koordynaty przekonwertowane na krotkę """
    startIndex = second_index(rawCoordinates, "(")
    point = rawCoordinates[startIndex + 1:-3]
    coordinates = [float(point) for point in point.split(",")]
    coordinates = tuple([number / 10 for number in coordinates])
    return coordinates


def check_direction(rawInfo):
    """ Sprawdza kierunek, w który zmierzają kolejne koordynaty. """
    infos = przerabiarka(rawInfo)
    infos = [round(infos[0], 1), round(infos[1], 1)]
    infos = tuple(infos)
    isReversed = True if abs(infos[1]) > 0 else False
    if infos[0] >= 0:
        if infos[1] < 0:
            isPositiveX = True
            isPositiveY = False
        elif infos[1] == 0:
            isPositiveX = True
            isPositiveY = True
        else:
            isPositiveX = False
            isPositiveY = True
    elif infos[1] > 0:
        isPositiveX = False
        isPositiveY = True
        isReversed = False
    else:
        isPositiveX = False
        isPositiveY = False
        isReversed = False
    return isPositiveX, isPositiveY, isReversed


def generate_polygonZone(startPoint, points, isPositiveX, isPositiveY, isReversed):
    """ Zwraca zestaw wierzchołków danego wielokąta strefy."""
    polygonZone = []
    if isReversed is True:
        for point in points:
            if isPositiveX is True:
                firstCoord = startPoint[0] + point[1]
            else:
                firstCoord = startPoint[0] - point[1]
            if isPositiveY is True:
                secondCoord = startPoint[1] + point[0]
            else:
                secondCoord = startPoint[1] - point[0]
            coordinates = (firstCoord, secondCoord)
            polygonZone.append(coordinates)
    else:
        for point in points:
            if isPositiveX is True:
                firstCoord = startPoint[0] + point[0]
            else:
                firstCoord = startPoint[0] - point[0]
            if isPositiveY is True:
                secondCoord = startPoint[1] + point[1]
            else:
                secondCoord = startPoint[1] - point[1]
            coordinates = (firstCoord, secondCoord)
            polygonZone.append(coordinates)
    return polygonZone


def generate_zoneInfo(space):
    """ Zwraca informacje o danej strefie wyciągnięte z pliku IFC. """
    x = second_index(space, ",")
    a = space[x+1:]
    index1 = a.find("'")
    index2 = second_index(a, "'")
    zonenumber = a[index1 + 1: index2]
    b = a[index2 + 1:]
    index1 = b.find("'")
    index2 = second_index(b, "'")
    zonename = b[index1 + 1: index2]
    zonename = decode_polish(zonename)
    zoneInfo = zonenumber + " " + zonename
    return zoneInfo


def decode_polish(text):
    """ Wstawia znaki polskie do nazw wyciągniętych z pliku IFC."""
    translator = {'\\X2\\0105\\X0\\': 'ą', '\\X2\\0104\\X0\\': 'Ą', '\\X2\\0106\\X0\\': 'Ć', '\\X2\\0107\\X0\\': 'ć',
                  '\\X2\\0119\\X0\\': 'ę', '\\X2\\0118\\X0\\': 'Ę', '\\X2\\017B\\X0\\': 'Ż', '\\X2\\017C\\X0\\': 'ż',
                  '\\X2\\0143\\X0\\': 'Ń', '\\X2\\0144\\X0\\': 'ń', '\\X2\\0179\\X0\\': 'Ź', '\\X2\\017A\\X0\\': 'ź',
                  '\\X2\\015A\\X0\\': 'Ś', '\\X2\\015B\\X0\\': 'ś', '\\X2\\00D3\\X0\\': 'Ó', '\\X2\\00F3\\X0\\': 'ó',
                  '\\X2\\0141\\X0\\': 'Ł', '\\X2\\0142\\X0\\': 'ł'}
    for string in translator:
        if text.find(string) != -1:
            text = text.replace(string, translator[string])
    return text


def multi_is_inside(polygon1, polygon2):
    """ Sprawdza, czy wszystkie wierzchołki danego wielokąta mieszczą się w drugim wielokącie."""
    placement = None
    for point in polygon2:
        placement = is_inside(polygon1, point)
        if placement is True:
            break
    return placement


def generate_polygon_element(vertices):
    """ Zwraca przybliżony zestaw wierzchołków danego elementu."""
    xMin, xMax, yMin, yMax = vertices
    const = 0.1
    xMin -= const
    xMax += const
    yMin -= const
    yMax += const
    lengthx = xMax - xMin
    lengthy = yMax - yMin
    polygonElement = []
    if lengthx > 100:
        i = xMin
        while i < xMax:
            i += 100
            if i >= xMax:
                break
            point = (i, yMin)
            point2 = (i, yMax)
            polygonElement.append(point)
            polygonElement.append(point2)
    if lengthy > 100:
        i = yMin
        while i < yMax:
            i += 100
            if i >= yMax:
                break
            point = (xMin, i)
            point2 = (xMax, i)
            polygonElement.append(point)
            polygonElement.append(point2)
    if lengthx <= 100 and lengthy <= 100:
        medX = (xMin + xMax) / 2
        medY = (yMin + yMax) / 2
        point1 = (medX, yMin)
        point2 = (xMin, medY)
        point3 = (medX, yMax)
        point4 = (xMax, medY)
        polygonElement = [point1, point2, point3, point4]
    return polygonElement


def auto_fit_worksheet_columns(ws):
    """ Automatycznie ustawia szerokość kolumn arkusza na podstawie maksymalnej długości tekstu w każdej kolumnie. """
    for columnCells in ws.columns:
        length = max(len(str(cell.value)) for cell in columnCells) + 8
        ws.column_dimensions[columnCells[0].column_letter].width = length


def set_format_of_cell(sheet, row, column, value, font=None, border=None, alignment=None):
    """" Formatuje daną komórkę arkusza."""
    sheet.cell(row, column).value = value
    if font is not None:
        sheet.cell(row, column).font = font
    if border is not None:
        sheet.cell(row, column).border = border
    if alignment is not None:
        sheet.cell(row, column).alignment = alignment


def is_all_properties_available(allPropertiesNames, neededProperties):
    """ Sprawdza, czy wszystkie właściwości wymagane dla danego skryptu są dostępne. """
    userProperties = []
    for propertyname in allPropertiesNames:
        typ = propertyname.type
        if typ == 'BuiltIn':
            name = propertyname.nonLocalizedName
        else:
            name = propertyname.localizedName
            userProperties.append(name)

    missingProperties = [
        property for property in neededProperties if property not in userProperties and property != []]
    if missingProperties:
        print("W twoim projekcie brakuje następujących wymaganych właściwości:")
        for property in missingProperties:
            print(property)
        print("Zaimportuj je, aby móc uruchomić ten skrypt.")
        return False
    else:
        print("Wszystkie wymagane właściwości są w projekcie. Uruchamiamy skrypt.")
        return True


def translate_types(polishTypes):
    """ Zwraca nazwy typów obiektów przetłumaczone na język angielski. """
    dictionary = {'Ściana': 'Wall', 'Słup': 'Column', 'Belka': 'Beam', 'Okno': 'Window', 'Drzwi': 'Door',
                  'Obiekt': 'Object', 'Lampa': 'Lamp', 'Strop': 'Slab', 'Dach': 'Roof', 'Siatka': 'Mesh',
                  'Strefa': 'Zone', 'Przegroda': 'CurtainWall', 'Powłoka': 'Shell', "Świetlik": 'Skylight',
                  'Kształt': 'Morph', 'Schody': 'Stair', 'Balustrada': 'Railing', 'Otwór': 'Opening'}
    englishTypes = []
    for type in polishTypes:
        type = type.lower().capitalize()
        englishType = dictionary[type]
        englishTypes.append(englishType)
    return englishTypes
