import os
import re
import xlsxwriter


def main():
    # Initialize variables
    lst_content_list = list[str]()
    result_list = list[list[str]]()

    # Read root directory from config file
    with open("Maschinenzeiten_TF.cfg", 'r') as f:
        directory = f.readline()

    # Loop through directories and files to get data
    lst_content_list = loop_through_directories(directory)

    i = -1
    for entry in lst_content_list:
        # Check if entry matches pattern for internal data with code 420
        if re.match("^DA,.+'INTERNAL_DATA',420,.+", entry):
            result = entry.split(',')[8].replace("'", '')
            if len(result) > 0:
                result_list.append([result, '', ''])
                i += 1

        # Check if entry matches pattern for internal data with code 570
        if re.match("^DA,.+'INTERNAL_DATA',570,.+", entry):
            result = entry.split(',')[8].replace("'", '').rstrip(" min")
            result_list[i][1] = result

            time = float(result)
            m = int(time)
            h = m / 60
            s = int(time * 60)
            result_list[i][2] = ":".join([format(h, "2"), format(m, "2"), format(s, "2")])

    # Write data to Excel file
    write_to_excel(result_list)

    # Open Excel file
    os.startfile('Maschinenzeiten_TF.xlsx')


def loop_through_directories(directory):
    # Initialize variable
    data = []

    # Loop through directories and files to get data
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.lst'):
                with open(os.path.join(root, file), 'r') as f:
                    [data.append(line) for line in f]
        for d in dirs:
            data.extend(loop_through_directories(os.path.join(root, d)))
    return data


def write_to_excel(data: list[list[str]]):
    # Open workbook
    workbook = xlsxwriter.Workbook('Maschinenzeiten_TF.xlsx')

    # Either get the first worksheet or create one if none exists
    if len(workbook.worksheets()) == 0:
        worksheet = workbook.add_worksheet()
    else:
        worksheet = workbook.worksheets()[0]

    # Write headers to worksheet
    worksheet.write('A1', 'Materialnummer')
    worksheet.write('B1', 'Zeit [min]')
    worksheet.write('C1', 'SD:MI:SE')

    # Write data to worksheet
    for i, value in enumerate([value for value in data if all(len(ele) > 0 for ele in value)]):
        i += 2
        worksheet.write('A' + str(i), value[0])
        worksheet.write('B' + str(i), format(float(value[1]), ".2f"))
        worksheet.write('C' + str(i), value[2])
        worksheet.autofit()

    # Close workbook
    workbook.close()


if __name__ == '__main__':
    main()
