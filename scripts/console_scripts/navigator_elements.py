from data_tools.connection_init import acc, act, acu


viewMapTree = acc.GetNavigatorItemTree(act.NavigatorTreeId('ViewMap'))
layoutBookTree = acc.GetNavigatorItemTree(act.NavigatorTreeId('LayoutBook'))

print('Mapa widoków: ')
level1 = viewMapTree.rootItem.children
print(level1[0].navigatorItem.name)
level2 = level1[0].navigatorItem.children
for zestaw in level2:
    typ = zestaw.navigatorItem.type
    if typ == 'FolderItem':
        name = zestaw.navigatorItem.name
        print(' ', name)
        level3 = zestaw.navigatorItem.children
        for arkusz in level3:
            name = arkusz.navigatorItem.name
            print('     ',name)
            typ = arkusz.navigatorItem.type
            if typ == 'FolderItem':
                level4 = arkusz.navigatorItem.children
                for rysunek in level4:
                    name = rysunek.navigatorItem.name
                    print('         ', name)
    else:
        name = zestaw.navigatorItem.name
        print(' ', name)
print()

level1 = layoutBookTree.rootItem.children
print("Układ arkuszy: ")
print(level1[0].navigatorItem.name)
level2 = level1[0].navigatorItem.children
toRename = []
for zestaw in level2:
    typ = zestaw.navigatorItem.type
    if typ == 'SubsetItem':
        name = zestaw.navigatorItem.name
        print(' ', name)
        level3 = zestaw.navigatorItem.children
        for arkusz in level3:
            name = arkusz.navigatorItem.name
            print('     ',name)
            level4 = arkusz.navigatorItem.children
            for rysunek in level4:
                name = rysunek.navigatorItem.name
                print('         ', name)
                '''if len(name) < 10:
                    guid = rysunek.navigatorItem.navigatorItemId.guid
                    newName = name + "_test"
                    Id = act.NavigatorItemId(guid)
                    acc.RenameNavigatorItem(Id, newName)''' # przykładowy pokaz procesu zmiany nazwy
    else:
        name = zestaw.navigatorItem.name
        print(' ', name)
        level3 = zestaw.navigatorItem.children
        for szablon in level3:
            name = szablon.navigatorItem.name
            print('     ', name)
