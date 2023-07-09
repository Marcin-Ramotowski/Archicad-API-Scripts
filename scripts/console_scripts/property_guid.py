from archicad import ACConnection

conn = ACConnection.connect()
assert conn

acc = conn.commands
act = conn.types
acu = conn.utilities

propertyID = acu.GetBuiltInPropertyId('General_3DLength')
print(propertyID.guid)