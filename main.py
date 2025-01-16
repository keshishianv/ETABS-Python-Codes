from getFrameMembers import get_frames
from writeFrameExcel import update_excel_with_data

from getFrameForces import get_frames_forces
from writeFrameForcesExcel import update_excel_with_frame_forces_data

from getFrameElm import get_Frame_ELM
from writeFrameElm import update_excel_with_FrameElm

from getPointNames import get_Point_Names
from writePointNames import update_excel_with_PointNames

# Get All Frame Data from Open ETABS Building
def run(file_string, unit_index):
    data = get_frames()
    update_excel_with_data(file_string, data)

    data_frame_forces = get_frames_forces(unit_index)
    update_excel_with_frame_forces_data(file_string, data_frame_forces)

    data_frame_ELM = get_Frame_ELM(unit_index)

    update_excel_with_FrameElm(file_string, data_frame_ELM)

    data_point_names = get_Point_Names(unit_index)
    update_excel_with_PointNames(file_string, data_point_names)

