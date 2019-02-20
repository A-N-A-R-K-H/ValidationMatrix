try:
    import csvToExcel
    import sys
    import os

    # Get the full path to the working directory #
    dir_path = os.path.dirname(os.path.realpath(__file__))
    # Path to the csv you want to put in Kvint39 Verification matrix
    pathCSV = r'\csv'
    pathCSV = dir_path + pathCSV
    # Path to the Kvint39 Verification matrix that will be used as a template.
    pathExcelCopy = r'\template\Kvint39_Verification_matrix.xlsx'
    pathExcelCopy = dir_path + pathExcelCopy
    # Path to ouput for the Kvint39 Verification matrix
    pathExcelOuput = r'\output\test.xlsx'
    pathExcelOuput = dir_path + pathExcelOuput
    # Frequency list for singleton in GHz
    freq_list_single = [37, 38.5, 40]
    # Frequency modulation bandwith offset in GHz
    mod_bw_offset = 0.5
    # Temperature list
    #temp_list = [-20, 55, 85]
    temp_list = [55]

    csvToExcel.run(pathCSV, pathExcelCopy, pathExcelOuput, temp_list, freq_list_single, mod_bw_offset)
    print("The data from DRS CSV has been transferred to the Valadition matrix file name: " + pathExcelOuput.split('\\')[-1])
except Exception as e:
    print('Error')
    print(str(e))

except:
    print("Unexpected error:", sys.exc_info()[0])

input("Press Enter to continue...")
