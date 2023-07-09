from data_tools.connection_init import acc, act, acu


all_names=acc.GetAllPropertyNames()
for propertyname in all_names:
    type = propertyname.type
    if type == 'BuiltIn':
        name = propertyname.nonLocalizedName
        print(name)
    else:
        name = propertyname.localizedName
        print(name)