import comtypes.client

def get_Frame_ELM(unit_index):
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
        allFrames = sap_model.FrameObj.GetALLFrames()

        # Clean Data by turning all data into lists/arrays
        data = []
        for kk in allFrames[1]:
            temp =  sap_model.FrameObj.GetElm(kk)
            dataTemp = []
            for i in temp:
                if isinstance(i, tuple):
                    dataTemp.append(list(i))
                else:
                    dataTemp.append([i])
            data.append([[kk]] + dataTemp)

        def pad_lists(lists):
            # Find the maximum length of the inner lists
            max_length = max(len(inner_list) for inner_list in lists)

            # Pad each inner list with empty strings
            padded_lists = [inner_list + [''] * (max_length - len(inner_list)) for inner_list in lists]

            return padded_lists

        paddedData =[]
        for i in data:
            paddedData.append(pad_lists(i))


        return paddedData

    except Exception as e:
        print(f"An error occurred: {e}")
        return []
