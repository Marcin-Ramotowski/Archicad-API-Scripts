from archicad import ACConnection

conn = ACConnection.connect()
assert conn
acc = conn.commands
act = conn.types

all_names=acc.GetAllPropertyNames()
for propertyname in all_names:
    type = propertyname.type
    if type == 'BuiltIn':
        name = propertyname.nonLocalizedName
        print(name)
    else:
        name = propertyname.localizedName
        print(name)