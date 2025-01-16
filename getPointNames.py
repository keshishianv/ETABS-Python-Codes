import comtypes.client

def get_Point_Names(unit_index):
    '''
    Connect to open ETABS page and pull the relevant info
    It cleans the data and turns all tuples into arrays and all single info into arrays.
    :return:
    '''
    try:
        # Attach to an existing ETABS instance
        helper = comtypes.client.CreateObject('ETABSv1.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        etabs_obj = helper.GetObject("CSI.ETABS.API.ETABSObject")

        # Get the SapModel interface
        sap_model = etabs_obj.SapModel

        ret = sap_model.SetPresentUnits(unit_index)


        # Call the API function

        allNames = sap_model.PointObj.GetNameList()

        xLoc = []
        yLoc = []
        zLoc = []

        for i in allNames[1]:
            temp = sap_model.PointObj.GetCoordCartesian(i)
            xLoc.append(temp[0])
            yLoc.append(temp[1])
            zLoc.append(temp[2])


        return [[allNames[0]],list(allNames[1]),xLoc,yLoc,zLoc]

    except Exception as e:
        print(f"An error occurred: {e}")
        return []
