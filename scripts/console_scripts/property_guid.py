from data_tools.connection_init import acu


propertyID = acu.GetBuiltInPropertyId('General_3DLength')
print(propertyID.guid)