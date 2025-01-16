import comtypes.client

def get_frames():
    '''
    Connect to open ETABS page and pull the relevant info
    It cleans the data and turns all tuples into arrays and all single info into arrays.
    :return:
    An Array full of Arrays of the following Data
        NumberNames,
        MyName,
        PropName,
        StoryName,
        PointName1,
        PointName2,
        Point1X,
        Point1Y,
        Point1Z,
        Point2X,
        Point2Y,
        Point2Z,
        Angle,
        Offset1X,
        Offset2X,
        Offset1Y,
        Offset2Y,
        Offset1Z,
        Offset2Z,
        CardinalPoint,
        Global

    '''
    try:
        # Attach to an existing ETABS instance
        helper = comtypes.client.CreateObject('ETABSv1.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        etabs_obj = helper.GetObject("CSI.ETABS.API.ETABSObject")

        # Get the SapModel interface
        sap_model = etabs_obj.SapModel

        # Call the API function
        allFrames = sap_model.FrameObj.GetALLFrames()

        # Clean Data by turning all data into lists/arrays
        data = []
        for i in allFrames:
            if isinstance(i, tuple):
                data.append(list(i))
            else:
                data.append([i])

        # Return cleaned list of data
        return data

    except Exception as e:
        print(f"An error occurred: {e}")
        return []

