import comtypes.client
import ast
import re

def get_frames_forces(unit_index):
    '''
    Connect to open ETABS page and pull the relevant info
    It cleans the data and turns all tuples into arrays and all single info into arrays.
    :return:
    An Array full of Arrays of the following Data
        NumberResults,
        Obj,
        ObjSta,
        Elm,
        ElmSta,
        LoadCase,
        StepType,
        StepNum,
        P,
        V2,
        V3,
        T,
        M2,
        M3


         lb_in_F    1
         lb_ft_F    2
         kip_in_F   3
         kip_ft_F   4
         kN_mm_C    5
         kN_m_C     6
         kgf_mm_C   7
         kgf_m_C    8
         N_mm_C     9
         N_m_C      10
         Ton_mm_C   11
         Ton_m_C    12
         kN_cm_C    13
         kgf_cm_C   14
         N_cm_C     15
         Ton_cm_C   16

    '''
    try:
        # Attach to an existing ETABS instance
        helper = comtypes.client.CreateObject('ETABSv1.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        etabs_obj = helper.GetObject("CSI.ETABS.API.ETABSObject")

        # Get the SapModel interface
        sap_model = etabs_obj.SapModel

        # Call the API function
        ret = sap_model.SetPresentUnits(unit_index)
        #4
        #cases = ['Dead', 'Live', 'LiveRoof', 'DeadSuper']
        #combos = ['1.2DL+1.6LL','1.4DL','DL','DL+LL']
        cases = list(sap_model.loadCases.GetNameList()[1])[:-1]
        combos = list(sap_model.respCombo.GetNameList()[1])

        # Raw Data List
        dataOrg = []

        # Loop through case names to get data
        for cas in cases:
            sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
            sap_model.Results.Setup.SetCaseSelectedForOutput(cas)
            frameForces = sap_model.Results.FrameForce("ALL", 2)
            # Clean up data by turning tuples to lists.
            temp = []
            for i in frameForces:
                if isinstance(i, tuple):
                    temp.append(list(i))
                else:
                    temp.append([i])
            # Create list of empty strings.
            emptyCombos = ["" for _ in range(temp[0][0])]
            # Insert list of empty strings into data.
            temp.insert(6, emptyCombos)
            # Insert list into Raw data List
            dataOrg.append(temp.copy())

            # Clear temp list for safety
            temp.clear()

        # Loop through Combo names to get data
        for com in combos:
            sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
            sap_model.Results.Setup.SetComboSelectedForOutput(com)
            frameForces = sap_model.Results.FrameForce("ALL", 2)
            # Clean up data by turning tuples to lists.
            temp = []
            for i in frameForces:
                if isinstance(i, tuple):
                    temp.append(list(i))
                else:
                    temp.append([i])
            # Create list of empty strings.
            emptyCombos = ["" for _ in range(temp[0][0])]
            temp.insert(5, emptyCombos)
            # Insert list into Raw data List
            dataOrg.append(temp.copy())

            # Clear temp list for safety
            temp.clear()


        # Restructure Lists to be organized and ready for excel.
        data = [[] for _ in range(len(dataOrg[0]))]
        for vv in range(len(dataOrg[0])):
            for ll in dataOrg:
                data[vv] = data[vv] + ll[vv]

        # Return cleaned list of case names, combo names and data as one big list.
        return [cases] + [combos] + data




    except Exception as e:
        print(f"An error occurred: {e}")
        return []
