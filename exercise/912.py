
import traceback
try:
    # print('Please drag in the folder for all the BS statements')
    path_folder_BS = input('Please drag in the folder for all the BS statements')
    path_folder_BS = path_folder_BS.replace('"', '')
    # path_folder_BS = sys.argv[1]
    # print('Please drag in the folder for all the GL files')
    path_folder_GL = input('Please drag in the folder for all the GL files')
    path_folder_GL = path_folder_GL.replace('"', '')
    # path_folder_GL = sys.argv[2]
    # print('Please drag in the folder for all reimbursement files')
    path_folder_reimRegister = input('Please drag in the folder for all reimbursement files')
    path_folder_reimRegister = path_folder_reimRegister.replace('"', '')
    # path_folder_reimRegister = sys.argv[3]
    # print('Please drag in the file of AP_Vendor Mapping')
    directory_AP_Vendor = input('Please drag in the file of AP_Vendor Mapping')
    directory_AP_Vendor = directory_AP_Vendor.replace('"', '')
    # directory_AP_Vendor = sys.argv[4]
    # print('Please drag in the file of AP_Employee Mapping')
    directory_AP_Employee = input('Please drag in the file of AP_Employee Mapping')
    directory_AP_Employee = directory_AP_Employee.replace('"', '')
    # directory_AP_Employee = sys.argv[5]
    # print('Please drag in the file of Commercial Mapping')
    directory_Commercial = input('Please drag in the file of Commercial Mapping')
    directory_Commercial = directory_Commercial.replace('"', '')
    # directory_Commercial = sys.argv[6]
    # print('Please drag in the folder for storing results')
    path_folder_target = input('Please drag in the folder for storing results')
    path_folder_target = path_folder_target.replace('"', '')
    # directory_Commercial = sys.argv[7]

    month_period = input("Please enter the covered month periods:")

    print('path_folder_BS', path_folder_BS)
    print('path_folder_GL', path_folder_GL)
    print('path_folder_reimRegister', path_folder_reimRegister)
    print('directory_AP_Vendor', directory_AP_Vendor)
    print('directory_AP_Employee', directory_AP_Employee)
    print('directory_Commercial', directory_Commercial)
    print('path_folder_target', path_folder_target)







except Exception as ex:
        print('Error Occurred')
        traceback.print_exc()

input('Finished:')

# directory_AP_Vendor = input("Please enter the file link of the AP_Vendor Mapping:")
# directory_AP_Employee = input("Please enter the file link of the AP_Employee Mapping:")
# directory_Commercial = input("Please enter the file link of the Commercial Mapping:")
# month_period = input("Please enter the covered month periods:")
# path_folder_target = input("Please enter the folder directory where you want to store the results:")