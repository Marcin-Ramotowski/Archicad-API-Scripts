from archicad import ACConnection
from typing import List, Tuple, Iterable
import string

conn = ACConnection.connect()
assert conn

acc = conn.commands
act = conn.types
acu = conn.utilities

################################ CONFIGURATION #################################
propertyId = acu.GetBuiltInPropertyId('General_ElementID')
classificationItem = acu.FindClassificationItemInSystem(
    'Klasyfikacja ARCHICAD', 'Krzesło')
elements = acc.GetElementsByClassification(
    classificationItem.classificationItemId)

ROW_GROUPING_LIMIT = 0.25

def GeneratePropertyValueString(rowIndex: int, indexInRow: int, isRight: bool) -> str:
    return f'{string.ascii_uppercase[rowIndex]}.{indexInRow}/{"Prawo" if isRight else "Lewo"}'
################################################################################


def generatePropertyValue(rowIndex: int, indexInRow: int, isRight: bool) -> act.NormalStringPropertyValue:
    return act.NormalStringPropertyValue(GeneratePropertyValueString(rowIndex, indexInRow, isRight))


def getXMin(elemWithBoundingBox: Tuple[act.ElementIdArrayItem, act.BoundingBox3DWrapper]) -> float:
    return elemWithBoundingBox[1].boundingBox3D.xMin


def generateNewPropertyValuesForElements(elemsWithBoundingBox: Iterable[Tuple[act.ElementIdArrayItem, act.BoundingBox3DWrapper]], isRight: bool, rowIndex: int):
    propertyValues = []
    indexInRow = 1
    for (elem, _) in sorted(elemsWithBoundingBox, key=getXMin, reverse=isRight):
        propertyValues.append(act.ElementPropertyValue(
            elem.elementId, propertyId, generatePropertyValue(rowIndex, indexInRow, isRight)))
        indexInRow += 1
    return propertyValues


def createClusters(positions: Iterable[float], limit: float) -> List[Tuple[float, float]]:
    positions = sorted(positions)
    if len(positions) == 0:
        return []

    clusters = []
    posIter = iter(positions)
    firstPos = lastPos = next(posIter)

    for pos in posIter:
        if pos - lastPos <= limit:
            lastPos = pos
        else:
            clusters.append((firstPos, lastPos))
            firstPos = lastPos = pos

    clusters.append((firstPos, lastPos))
    return clusters

boundingBoxes = acc.Get3DBoundingBoxes(elements)

averageXPosition = sum([bb.boundingBox3D.xMin for bb in boundingBoxes]) / len(boundingBoxes)
zClusters = createClusters((bb.boundingBox3D.zMin for bb in boundingBoxes), ROW_GROUPING_LIMIT)
elementsWithBoundingBoxes = list(zip(elements, boundingBoxes))

rowIndex = 0
elemPropertyValues = []
for (zMin, zMax) in zClusters:
    elemsInRow = [e for e in elementsWithBoundingBoxes if zMin <= e[1].boundingBox3D.zMin <= zMax]
    rightSide, leftSide = [], []
    for elemWithBoundingBox in elemsInRow:
        if getXMin(elemWithBoundingBox) > averageXPosition:
            leftSide.append(elemWithBoundingBox)
        else:
            rightSide.append(elemWithBoundingBox)

    elemPropertyValues.extend(generateNewPropertyValuesForElements(rightSide, True, rowIndex))
    elemPropertyValues.extend(generateNewPropertyValuesForElements(leftSide, False, rowIndex))
    rowIndex += 1

acc.SetPropertyValuesOfElements(elemPropertyValues)

# Print the result
newValues = acc.GetPropertyValuesOfElements(elements, [propertyId])
elemAndValuePairs = [(elements[i].elementId.guid, v.propertyValue.value) for i in range(len(newValues)) for v in newValues[i].propertyValues]
for elemAndValuePair in sorted(elemAndValuePairs, key=lambda p: p[1]):
    print(elemAndValuePair)
