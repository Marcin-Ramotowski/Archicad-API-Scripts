from data_tools.connection_init import acc, act


viewMapTree = acc.GetNavigatorItemTree(act.NavigatorTreeId('ViewMap'))
layoutBookTree = acc.GetNavigatorItemTree(act.NavigatorTreeId('LayoutBook'))

def getElementsFromLevel(childrens, level=0):
    for elem in childrens:
        name = elem.navigatorItem.name
        print(' ' * level * 2, name)
        childrens = elem.navigatorItem.children
        if childrens:
            getElementsFromLevel(childrens, level+1)


print('Mapa widoków: ')
getElementsFromLevel(viewMapTree.rootItem.children)

print()

print("Układ arkuszy: ")
getElementsFromLevel(layoutBookTree.rootItem.children)
