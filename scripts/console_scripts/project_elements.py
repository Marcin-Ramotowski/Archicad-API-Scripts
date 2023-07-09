from data_tools.connection_init import acc


alls = acc.GetAllElements()
walls = acc.GetElementsByType("Wall")
columns = acc.GetElementsByType("Column")
ceilings = acc.GetElementsByType("Slab")
beams = acc.GetElementsByType("Beam")
roofs = acc.GetElementsByType("Roof")
doors = acc.GetElementsByType("Door")
windows = acc.GetElementsByType("Window")
stairs = acc.GetElementsByType("Stair")

print("Liczba ścian w projekcie:", len(walls))
print("Liczba słupów w projekcie:", len(columns))
print("Liczba stropów w projekcie:", len(ceilings))
print("Liczba belek w projekcie:", len(beams))
print("Liczba dachów w projekcie:", len(roofs))
print("Liczba drzwi w projekcie:", len(doors))
print("Liczba okien w projekcie:", len(windows))
print("Liczba schodów w projekcie:", len(stairs))
print("Liczba wszystkich elementów projektu:", len(alls))
