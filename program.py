from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import xlrd
import xlsxwriter
import tkinter.filedialog

# Building the Gui including all the frames
root = Tk()
root.title('VertiCal')
root.wm_iconbitmap('sfsf logo.ico')
root.geometry('440x440+500+200')

# Setting heights and widths for every frame.
frame0 = Frame(height=65, width=400)
frame00 = Frame(height=65, width=400)
frame1 = Frame(height=75, width=400)
frame2 = Frame(height=40, width=400)
frame3 = Frame(height=120, width=400)
frame4 = Frame(height=120, width=400)
frame5 = Frame(height=8000, width=500)
frame6 = Frame(height=120, width=400)
frame7 = Frame(height=120, width=400)
frame8 = Frame(height=120, width=400)
frame9 = Frame(height=120, width=400)
frame10 = Frame(height=90, width=400)
frame11 = Frame(height=70, width=400)
frame110 = Frame(height=140, width=400)
frame111 = Frame(height=40, width=400)
frame12 = Frame(height=120, width=400)
frame13 = Frame(height=120, width=400)
frame14 = Frame(height=290, width=400)
frame15 = Frame(height=120, width=400)
frame16 = Frame(height=90, width=400)
frame160 = Frame(height=120, width=400)
frame17 = Frame(height=50, width=400)
frame170 = Frame(height=120, width=400)
frame18 = Frame(height=140, width=400)
frame180 = Frame(height=120, width=400)
frame19 = Frame(height=70, width=400)
frame190 = Frame(height=120, width=400)
frame20 = Frame(height=90, width=400)
frame200 = Frame(height=120, width=400)
frame21 = Frame(height=70, width=400)
frame210 = Frame(height=120, width=400)
frame_finish = Frame(height=75, width=400)
all_frames = [frame1, frame2, frame3, frame4, frame5, frame6, frame7, frame8, frame9, frame10, frame11,
              frame12, frame13, frame14, frame15, frame16, frame160, frame17, frame170, frame18,
              frame180, frame19, frame190, frame20, frame200, frame21, frame210, frame_finish]

# Set all the frames to a certain size
for frame in all_frames:
    frame.grid_propagate(0)

# ----------------------------------------
# Below, all functions of the program are created

# Function worksheetoutput does all the calculations to perform the LCA and writes the result to an Excel sheet
def worksheetoutput(dictionary_name):
    # Opening excel file in order to get parameters
    workbook = xlrd.open_workbook('Database_full.xlsx')
    # sheet = workbook.sheet_by_name(ans1.get())

    # Package.
    # 1 = yes, it is packaged on the farm
    # 2 = No, it is not packaged.
    # Moet nog gedaan worden
    # if ans171.get() == 1:
    #     Pac1 = float(sheet.cell_value(1, 39))  # Packaging Co2 equivalent/kg
    #     Pac2 = float(sheet.cell_value(1, 40))  # Packaging energy equivalent/kg
    # else:
    #     Pac1 = 0
    #     Pac2 = 0
    Pac1 = 0
    Pac2 = 0
    # C1 - C10 are emission values for electricity, retrieved from an Excel sheet
    # eventually create a loop for this.
    dicp = {}

    # sheet = workbook.sheet_by_name('CO2eq (kg)')

    for tabs in workbook.sheet_names():
        if tabs == 'Crop parameters':
            sheet = workbook.sheet_by_name(tabs)
            count = 1
            for keys, values in dictionary_name.items():
                if keys == sheet.cell_value(count, 0):
                    dictionary_name[keys] += [sheet.cell_value(count, 4)]
                count += 1

        else:
            sheet = workbook.sheet_by_name(tabs)
            for i in range(1, sheet.nrows):
                if sheet.cell_value(i, 2) != '' and sheet.cell_value(i, 3) != '':
                    country = 4  # 4 is Netherlands
                    if sheet.cell_value(i, country) == '':
                        country = 3  # 3 is average
                    dicp[sheet.cell_value(i, 2)] = sheet.cell_value(i, country)

    non_count = str()
    # If choose 'I don't know’ option, set the value back to zero
    if ans890.get() == 1:
        ans87.set(0);
        ans88.set(0);
        non_count = ('Specification of electricity,')
    if ans133.get() == 1:
        ans121.set(0);
        ans122.set(0);
        ans123.set(0);
        ans124.set(0);
        ans125.set(0);
        ans126.set(0);
        ans127.set(0);
        ans128.set(0);
        ans131.set(0);
        ans132.set(0)
        non_count = (non_count + 'NPK chemicals,')
    if ans147.get() == 1:
        ans141.set(0);
        ans142.set(0);
        ans143.set(0);
        ans144.set(0);
        ans145.set(0);
        ans146.set(0)
        non_count = (non_count + 'Substrate,')
    if ans152.get() == 1:
        ans151.set(0)
        non_count = (non_count + 'Water,')
    if ans166.get() == 1:
        ans161.set(0);
        ans162.set(0);
        ans163.set(0);
        ans164.set(0);
        ans165.set(0)
        non_count = (non_count + 'Pesticides,')
    if ans204.get() == 1:
        ans201.set(0);
        ans202.set(0);
        non_count = (non_count + 'NPK chemicals,')


    # Create the output: an Excel file
    wb = xlsxwriter.Workbook(farm_name.get() + '.xlsx')

    sheet = workbook.sheet_by_name('Crop parameters')
    Total_Eoc = 0
    for keys, values in dictionary_name.items():
        for i in range(1, len(lis)+1):
            if keys == sheet.cell_value(i, 0):
                dictionary_name[keys] += [sheet.cell_value(i, 1)]
                Total_Eoc += sheet.cell_value(i, 1)
    Average_Eoc = Total_Eoc / (len(dictionary_name) - 1)
    dictionary_name[list(dictionary_name.keys())[0]] += [Average_Eoc]

    for keys, values in dictionary_name.items():
        cropname = keys
        kg_prod = values[2]
        frac_surf = values[0]
        frac_kg = values[1]
        Eoc = values[3]

        # Calculation for total C02 of electricity usage
        Eco2 = frac_surf * ((dicp['C1'] * ans61.get()) + (dicp['C3'] * ans62.get()) + (dicp['C5'] * ans71.get()) + (dicp['C7'] * ans73.get()) + (
                dicp['C9'] * ans72.get()) - (ans87.get() * dicp['C1']) - (ans88.get() * dicp['C3']))

        # Calculation for total energy of electricity usage
        Eenergy = frac_surf * ((dicp['C2'] * ans61.get()) + (dicp['C4'] * ans62.get()) + (dicp['C6'] * ans71.get()) + (dicp['C8'] * ans73.get()) + (
                dicp['C10'] * ans72.get()) - (ans87.get() * dicp['C2']) - (ans88.get() * dicp['C4']))

        # Calculation for total Co2 of fossil fuels use
        Fco2 = frac_surf * ((dicp['Fo1'] * ans91.get()) + (dicp['Fo3'] * ans92.get()) + (dicp['Fo7'] * ans94.get()) + (dicp['Fo9'] * ans95.get()))

        # Calculation for total energy of fossil fuel use
        Fenergy = frac_surf * (
                (dicp['Fo2'] * ans91.get()) + (dicp['Fo4'] * ans92.get()) + (dicp['Fo8'] * ans94.get()) + (dicp['Fo10'] * ans95.get()))

        # Calculation for total Co2 of fertilizers
        FERco2 = frac_surf * ((
                (dicp['Fe1'] * ans121.get()) + (dicp['Fe3'] * ans122.get()) + (dicp['Fe5'] * ans123.get()) + (dicp['Fe7'] * ans124.get()) + (
                dicp['Fe9'] * ans125.get()) + (dicp['Fe11'] * ans126.get()) + (dicp['Fe13'] * ans127.get()) + (
                        dicp['Fe15'] * ans128.get()) + (dicp['Fe21'] * ans131.get()) + (dicp['Fe22'] * ans132.get())))

        # Calculation for total energy of fertilizers
        FERenergy = frac_surf * ((
                (dicp['Fe2'] * ans121.get()) + (dicp['Fe4'] * ans122.get()) + (dicp['Fe6'] * ans123.get()) + (dicp['Fe8'] * ans124.get()) + (
                dicp['Fe10'] * ans125.get()) + (dicp['Fe12'] * ans126.get()) + (dicp['Fe14'] * ans127.get()) + (
                        dicp['Fe16'] * ans128.get()) + (dicp['Fe22'] * ans131.get()) + (dicp['Fe24'] * ans132.get())))

        # Calculation for total Co2 of substrates
        Sco2 = frac_surf * (
                (dicp['S1'] * ans141.get()) + (dicp['S3'] * ans142.get()) + (dicp['S5'] * ans143.get()) + (dicp['S7'] * ans144.get()) + (
                dicp['S9'] * ans145.get()) + (dicp['S11'] * ans146.get()))

        # Calculation for total energy of substrates
        Senergy = frac_surf * (
                (dicp['S2'] * ans141.get()) + (dicp['S4'] * ans142.get()) + (dicp['S6'] * ans143.get()) + (dicp['S8'] * ans144.get()) + (
                dicp['S10'] * ans145.get()) + (dicp['S12'] * ans146.get()))

        # Calculation for total Co2 of water
        Wco2 = frac_surf * (dicp['Wa1'] * ans151.get())

        # Calculation for total energy of water
        Wenergy = frac_surf * (dicp['Wa2'] * ans151.get())

        # Calculation for total Co2 of pesticides
        Pco2 = frac_surf * (
                (dicp['P1'] * ans161.get()) + (dicp['P3'] * ans162.get()) + (dicp['P5'] * ans163.get()) + (dicp['P7'] * ans164.get()) + (
                dicp['P9'] * ans165.get()))

        # Calculation for total energy of pesticides
        Penergy = frac_surf * (
                (dicp['P2'] * ans161.get()) + (dicp['P4'] * ans162.get()) + +(dicp['P6'] * ans163.get()) + +(dicp['P8'] * ans164.get()) + (
                dicp['P10'] * ans165.get()))

        # Calculation for total Co2 of transport
        Tco2 = frac_kg * ((dicp['T3'] * ans201.get()) + (dicp['T1'] * ans202.get()) )

        # Calculation for total energy of transport
        Tenergy = frac_kg * ((dicp['T4'] * ans201.get()) + (dicp['T2'] * ans202.get()) )

        # Calculation for the total Co2 of packaging
        Pacco2 = kg_prod * Pac1

        # Calculation for the total energy of packaging
        Pacenergy = kg_prod * Pac2

        # calculations for the total Co2 and energy
        Totalco2 = Eco2 + Fco2 + FERco2 + Sco2 + Wco2 + Pco2 + Tco2 + Pacco2
        Totalenergy = Eenergy + Fenergy + FERenergy + Senergy + Wenergy + Penergy + Tenergy + Pacenergy

        # calculations for the total Co2 and energy per kg product #ans5 moet kg worden
        Totalco2_per_kg_product = Totalco2 / kg_prod
        Totalenergy_per_kg_product = Totalenergy / kg_prod

        # calculations for the total Co2 and energy per KJ product
        Totalco2_per_KJ_product = Totalco2_per_kg_product / Eoc
        Totalenergy_per_KJ_product = Totalenergy_per_kg_product / Eoc

        # Writing the outputs to the previously created Excel sheet
        ws = wb.add_worksheet(cropname)
        ws.write(0, 2, 'Total CO2 emitted(Kg)')
        ws.write(0, 4, 'Total energy used(MJ)')

        # also labels in Dutch/other languages?
        labels_output = ['Electricity', 'Fossil fuels', 'Fertilizer', 'Substrates', 'Water', 'Pesticides',
                         'Transport', 'Package']
        Co2_emitted = [Eco2, Fco2, FERco2, Sco2, Wco2, Pco2, Tco2, Pacco2]
        Co2_emitted_round = [round(elem, 2) for elem in Co2_emitted]
        energy_used = [Eenergy, Fenergy, FERenergy, Senergy, Wenergy, Penergy, Tenergy, Pacenergy]
        energy_used_round = [round(elem, 2) for elem in energy_used]

        for x in range(len(labels_output)):
            ws.write(1 + x, 0, labels_output[x])
            ws.write(1 + x, 2, Co2_emitted_round[x])
            ws.write(1 + x, 4, energy_used_round[x])

        labels_total = ['Total CO2 emitted (Kg)', 'Total energy used(MJ)',
                        'Total CO2 emitted per kg product (Kg/Kg)', 'Total Energy used per kg product (KJ/Kg)',
                        'Total CO2 emitted per KJ product (Kg/KJ)', 'Total energy used per KJ product (KJ/KJ)']
        totals_output = [Totalco2, Totalenergy, Totalco2_per_kg_product, Totalenergy_per_kg_product,
                         Totalco2_per_KJ_product, Totalenergy_per_KJ_product]
        totals_output_round = [round(elem, 2) for elem in totals_output]
        for x in range(len(labels_total)):
            ws.write(1 + x, 6, labels_total[x])
            ws.write(1 + x, 9, totals_output_round[x])

        # labels_diff_aspects = ['Heating', 'Cooling', 'Electricity', 'Tillage', 'Sowing', 'Weeding', 'Harvest',
        #                        'Fertilizer', 'Irrigation', 'Pesticide', 'Other']
        # for x in range(len(labels_diff_aspects)):
        #     ws.write(43 + x, 1, labels_diff_aspects[x])
        #     ws.write(43 + x, 2, list_ans[22 + x].get()) * frac_surf

        # # These values are currently written to a super random place in the script, it needs to be reconsidered
        # # ws.write(43, 9, 'Heating')
        # # ws.write(44, 9, 'Cooling')
        # # ws.write(45, 9, 'Ventilation')
        # # ws.write(46, 9, 'Lighting')
        # # ws.write(47, 9, 'Machinery')
        # # ws.write(48, 9, 'Storage')
        # ws.write(49, 9, 'Selling renewables')
        # ws.write(50, 9, 'Selling non-renewables')

        # # ws.write(43, 10, ans81.get()) * frac_surf
        # # ws.write(44, 10, ans82.get()) * frac_surf
        # # ws.write(45, 10, ans83.get()) * frac_surf
        # # ws.write(46, 10, ans84.get()) * frac_surf
        # # ws.write(47, 10, ans85.get()) * frac_surf
        # # ws.write(48, 10, ans86.get()) * frac_surf
        # ws.write(49, 10, ans87.get()) * frac_surf
        # ws.write(50, 10, ans88.get()) * frac_surf

        if ans890.get() == 1 or ans131.get() == 1 or ans147.get() == 1 or ans152.get() == 1 or ans166.get() == 1 or ans204.get() == 1:
            ws.write(10, 1, non_count + 'is not taken into account because of lacking data')

        # Creating bar and pie charts
        chart_col = wb.add_chart({'type': 'column'})
        chart_col.add_series({
            'name': [cropname, 0, 2],
            'categories': [cropname, 1, 0, 10, 0],
            'values': [cropname, 1, 2, 10, 2],
            'line': {cropname: 'yellow'}
        })
        chart_col.set_title({'name': 'Total CO2 emitted from different sources'})
        chart_col.set_y_axis({'name': 'Total CO2 emitted'})
        chart_col.set_x_axis({'name': 'Different sources'})

        chart_col.set_style(2)
        ws.insert_chart('A13', chart_col, {'x_offset': 20, 'y_offset': 8})

        chart_col = wb.add_chart({'type': 'column'})
        chart_col.add_series({
            'name': [cropname, 0, 4],
            'categories': [cropname, 1, 0, 10, 0],
            'values': [cropname, 1, 4, 10, 4],
            'line': {'color': 'yellow'}
        })
        chart_col.set_title({'name': 'Total energy used from different sources'})
        chart_col.set_y_axis({'name': 'Total energy used'})
        chart_col.set_x_axis({'name': 'Different sources'})

        chart_col.set_style(2)
        ws.insert_chart('I13', chart_col, {'x_offset': 20, 'y_offset': 8})

        chart_col = wb.add_chart({'type': 'pie'})
        chart_col.add_series({
            'name': [cropname, 0, 1],
            'categories': [cropname, 1, 0, 10, 0],
            'values': [cropname, 1, 2, 10, 2],
            'points': [{'fill': {'color': 'blue'}},
                       {'fill': {'color': 'yellow'}},
                       {'fill': {'color': 'red'}},
                       {'fill': {'color': 'gray'}},
                       {'fill': {'color': 'black'}},
                       {'fill': {'color': 'purple'}},
                       {'fill': {'color': 'pink'}},
                       ],
        })
        chart_col.set_title({'name': 'Total CO2 emitted from different \nsources'})
        chart_col.set_style(2)
        ws.insert_chart('A28', chart_col, {'x_offset': 20, 'y_offset': 8})

        chart_col = wb.add_chart({'type': 'pie'})
        chart_col.add_series({
            'name': [cropname, 0, 2],
            'categories': [cropname, 1, 0, 10, 0],
            'values': [cropname, 1, 4, 10, 4],
            'points': [{'fill': {'color': 'blue'}},
                       {'fill': {'color': 'yellow'}},
                       {'fill': {'color': 'red'}},
                       {'fill': {'color': 'gray'}},
                       {'fill': {'color': 'black'}},
                       {'fill': {'color': 'purple'}},
                       {'fill': {'color': 'pink'}},
                       ],
        })
        chart_col.set_title({'name': 'Total energy used from different sources'})
        chart_col.set_style(2)
        ws.insert_chart('I28', chart_col, {'x_offset': 20, 'y_offset': 8})

        # chart_col = wb.add_chart({'type': 'pie'})
        # chart_col.add_series({
        #     'name': 'Fossil fuels used',
        #     'categories': [cropname, 43, 1, 53, 1],
        #     'values': [cropname, 43, 2, 53, 2],
        #     'points': [{'fill': {'color': 'blue'}},
        #                {'fill': {'color': 'yellow'}},
        #                {'fill': {'color': 'red'}},
        #                {'fill': {'color': 'gray'}},
        #                {'fill': {'color': 'black'}},
        #                {'fill': {'color': 'purple'}},
        #                {'fill': {'color': 'pink'}},
        #                {'fill': {'color': 'cyan'}},
        #                {'fill': {'color': 'magenta'}},
        #                {'fill': {'color': 'brown'}},
        #                ],
        # })
        # chart_col.set_title({'name': 'Fossil fuels used for different aspects'})
        # chart_col.set_style(5)
        # ws.insert_chart('A43', chart_col, {'x_offset': 25, 'y_offset': 10})

        # chart_col = wb.add_chart({'type': 'pie'})
        # chart_col.add_series({
        #     'name': 'Electricity used',
        #     'categories': [cropname, 43, 9, 50, 9],
        #     'values': [cropname, 43, 10, 50, 10],
        #     'points': [{'fill': {'color': 'blue'}},
        #                {'fill': {'color': 'yellow'}},
        #                {'fill': {'color': 'red'}},
        #                {'fill': {'color': 'gray'}},
        #                {'fill': {'color': 'black'}},
        #                {'fill': {'color': 'purple'}},
        #                {'fill': {'color': 'pink'}},
        #                {'fill': {'color': 'cyan'}},
        #
        #                ],
        # })
        # chart_col.set_title({'name': 'Electricity used for different aspects'})
        # chart_col.set_style(4)
        # ws.insert_chart('I43', chart_col, {'x_offset': 25, 'y_offset': 10})

    wb.close()
    # root.destroy()
    return
    # ^^ End of function worksheet output


# def pre() enables to go back to the previous question
# i.e. forgetting the current frames and introducing new frames ??
count = 0


def pre():
    global count
    for i in range(len(list_ans)):  # if there is no value in Entry, make it back to 0
        try:
            if i != 0 or 2 or 1:
                list_ans[i].get() != ''

        except TclError:
            list_ans[i].set(00)

    count -= 1
    if count < 1:
        count = 1
    if count == 1:
        var.set('1. In which country is your farm located?')
        frame5.grid_forget()
        frame3.grid(sticky=W)
    if count == 2:
        var.set('2. Which crops do you produce? \nWhat area is each crop grown on? \nHow many kg of each crop do you sell every year?')
        frame8.grid_forget()
        frame5.grid(sticky=W)
    if count == 3:
        var.set(
            '3. How much renewable and non-renewable electricity (kWh) \ndo you buy per year?')
        frame9.grid_forget()
        frame8.grid(sticky=W)
    if count == 4:
        var.set('4. Do you produce your own renewable energy and how much (kWh) \ndo you produce?')
        frame10.grid_forget()
        frame9.grid(sticky=W)
    if count == 5:
        var.set("5. How much electricity (kWh) do you sell?")
        frame11.grid_forget()
        frame110.grid_forget()
        frame111.grid_forget()
        frame10.grid(sticky=W)
    if count == 6:
        var.set("6. Do you use any fossil fuels (excluding transportation), \nand how much do you use?")
        frame14.grid_forget()
        frame11.grid(sticky=W)
        frame111.grid(sticky=W)
        frame110.grid(sticky=W)
    if count == 7:
        var.set('7. How many NPK chemicals (kg) do you use per year? ')
        frame16.grid_forget()
        frame160.grid_forget()
        frame14.grid(sticky=W)
    if count == 8:
        var.set('8. Do you use substrate (kg) and how much per year?')
        frame17.grid_forget()
        frame170.grid_forget()
        frame16.grid(sticky=W)
        frame160.grid(sticky=W)
    if count == 9:
        var.set('9. How much water (L) do you buy?') #Check whether this is in litres or in m3
        frame18.grid_forget()
        frame180.grid_forget()
        frame17.grid(sticky=W)
        frame170.grid(sticky=W)
    if count == 10:
        var.set('10. How much pesticides (kg) do you use? ')
        frame19.grid_forget()
        frame190.grid_forget()
        frame18.grid(sticky=W)
        frame180.grid(sticky=W)
    if count == 11:
        var.set('11. Is the product sold to the customer packaged? ')
        frame20.grid_forget()
        frame200.grid_forget()
        frame19.grid(sticky=W)
        frame190.grid(sticky=W)
    if count == 12:
        var.set('12. How far (km) does your product travel to the distribution center \non average?')
        frame_finish.grid_forget()
        frame20.grid(sticky=W)
        frame200.grid(sticky=W)
    if count == 13:
        var.set('13. This is was the questionnaire, are you finished?')
        frame_finish.grid(sticky=W)
    return


# def next1() enables to go to the next question.
# i.e. forgetting the current frames and introducing new frames
# v = IntVar()


def next1():
    global count
    global v
    for i in range(len(list_ans)):  # if there is no value in Entry, set it back to 0
        try:
            if i != 0 or 2 or 1:
                list_ans[i].get() != ''
        except TclError:
            list_ans[i].set(00)
    count += 1
    if count == 2:
        var.set('2. Which crops do you produce? \nWhat area is each crop grown on? \nHow many kg of each crop do you sell every year? ')
        frame3.grid_forget()
        frame5.grid(sticky=W)
    if count == 3:
        var.set(
            '3. How much renewable and non-renewable electricity (kWh)\ndo you buy per year?')
        frame5.grid_forget()
        frame8.grid(sticky=W)
    if count == 4:
        var.set('4. Do you produce your own renewable energy and how much (kWh) \ndo you produce?')
        frame8.grid_forget()
        frame9.grid(sticky=W)
    if count == 5:
        var.set("5. How much electricity (kWh) do you sell?")
        frame9.grid_forget()
        frame10.grid(sticky=W)
    if count == 6:
        var.set(
            "6. Do you use any fossil fuels (excluding transportation), \nand how much do you use?")
        frame10.grid_forget()
        frame11.grid(sticky=W)
        frame111.grid(sticky=W)
        frame110.grid(sticky=W)
    if count == 7:
        var.set('7. How many NPK chemicals (kg) do you use per year? ')
        frame11.grid_forget()
        frame111.grid_forget()
        frame110.grid_forget()
        frame14.grid(sticky=W)
    if count == 8:
        var.set('8. Do you use substrate (kg) and how much per year?')
        frame14.grid_forget()
        frame16.grid(sticky=W)
        frame160.grid(sticky=W)
    if count == 9:
        var.set('9. How much water (L) do you buy?')
        frame16.grid_forget()
        frame160.grid_forget()
        frame17.grid(sticky=W)
        frame170.grid(sticky=W)
    if count == 10:
        var.set('10. How much pesticides (kg) do you use?')
        frame17.grid_forget()
        frame170.grid_forget()
        frame18.grid(sticky=W)
        frame180.grid(sticky=W)
    if count == 11:
        var.set('11. Is the product sold to the customer packaged? ')
        frame18.grid_forget()
        frame180.grid_forget()
        frame19.grid(sticky=W)
        frame190.grid(sticky=W)
    if count == 12:
        var.set('12. How far (km) does your produce travel to the distribution center \non average?')
        frame19.grid_forget()
        frame190.grid_forget()
        frame20.grid(sticky=W)
        frame200.grid(sticky=W)
    if count == 13:
        var.set('13. This is was the questionnaire, are you finished? ')
        frame20.grid_forget()
        frame200.grid_forget()
        frame_finish.grid(sticky=W)
    return


# Closes the program
def quit1():
    root.destroy()
    return

def close_program():
    worksheetoutput(dic_crops)
    quit1()
    return


# Enables the key 'enter' to go to the next question
def enter(event):
    if count > 0:
        next1()
    return


# The command of 'Start' button at the beginning
def start():
    frame0.pack_forget()
    frame00.pack(anchor=CENTER)

    return


# The command of 'Next' button after you input the farm name
def next2():
    global count
    frame00.pack_forget()
    frame1.grid()
    frame2.grid()
    frame3.grid()
    count += 1
    # The code below is necessary in the last frame to instruct the user where he can find the results of the analysis.
    print_finish = 'If you click on the finish button, the questionnaire will close.'
    print_finish_2 = 'Results of the analysis can then be found in: '
    print_farm_name = farm_name.get() + '.xlsx.'
    Label(frame_finish, text=print_finish).grid(row=1, column=0, sticky=W)
    Label(frame_finish, text=print_finish_2).grid(row=2, column=0, sticky=W)
    Label(frame_finish, text=print_farm_name).grid(row=3, column=0, sticky=W)
    return



# The function file_open can load previously filled in data, stored in a .txt file
def file_open():
    try:
        path1 = StringVar()

        path1 = tkinter.filedialog.askopenfilename()
        f = open(path1)
        lines = f.readlines()
        num = -1
        for line in lines:
            num += 1
            if num < 3:
                list_ans[num].set(str(line.strip('\n')))
            if num >= 3:
                list_ans[num].set(int(line.strip('\n')))
    except:
        pass
    return


# The function file_save saves data filled in in a questionnaire in a .txt file
def file_save():
    try:
        path2 = StringVar()
        path2 = tkinter.filedialog.asksaveasfilename(**root.file_opt)
        f1 = open(path2, 'w')
        for i in range(len(list_ans)):
            f1.write(str(list_ans[i].get()) + '\n')
        f1.close()
    except:
        pass
    return


# cal2 is a function that processes answers on Q2 into a dictionary for use in function 'worksheetoutput'
def cal2(event):
    global dic_crops
    total_area = 0
    total_kg = 0
    for i in range(0, len(ansVeg)):
        if ansVeg[i].get() == 0:
            kgVeg[i].set(0)
            surVeg[i].set(0)

        total_area += surVeg[i].get()
        total_kg += kgVeg[i].get()

    # Calculating the fraction crop over the full area and fraction of kg
    fracLetsur = fracEndsur = fracSpisur = fracBeasur = fracParsur = fracKalsur = fracBassur = fracRucsur = fracMicsur = fracMinsur = 0
    frac_sur = [fracLetsur, fracEndsur, fracSpisur, fracBeasur, fracParsur, fracKalsur, fracBassur, fracRucsur,
                fracMicsur,fracMinsur]
    fracLetkg = fracEndkg = fracSpikg = fracBeakg = fracParkg = fracKalkg = fracBaskg = fracRuckg = fracMickg = fracMinkg = 0
    frac_kg = [fracLetkg, fracEndkg, fracSpikg, fracBeakg, fracParkg, fracKalkg, fracBaskg, fracRuckg, fracMickg, fracMinkg]
    for i in range(0, len(frac_sur)):
        frac_sur[i] = surVeg[i].get() / total_area
        frac_kg[i] = kgVeg[i].get() / total_kg

    # Creating a dictionary of all parameters: [fraction surface, fraction kg,kg vegetation]
    dic_crops = {}
    dic_crops['Total'] = [1, 1, total_kg]
    for i in range(0, len(frac_sur)):
        dic_crops[lis[i]] = [frac_sur[i], frac_kg[i], kgVeg[i].get()]
    dic_crops = {x: y for x, y in dic_crops.items() if y != [0, 0, 0]}
    return dic_crops

def rid_of_zeros_sur(event):
    for i in range(0, len(ansVeg)):
        if ansVeg[i].get() == 1 and surVeg[i].get() <= 0:
            surVeg[i].set('')
        if ansVeg[i].get() == 0 and surVeg[i].get() >= 0:
            surVeg[i].set(0)
    return

def rid_of_zeros_kg(event):
    for i in range(0, len(ansVeg)):
        if ansVeg[i].get() == 1 and kgVeg[i].get() <= 0:
            kgVeg[i].set('')
        if ansVeg[i].get() == 0 and kgVeg[i].get() > 0:
            kgVeg[i].set(0)
    return


# ^^ End of functions for the program. Below, the GUI of the program is further developed.
# ------------------------------------------
# Here the start button at the beginning is created
startbutton = Button(frame0, text='Start', command=start, font=12)
startbutton.pack(fill=X, side=BOTTOM, anchor=CENTER)

# The first page you see when starting the questionnaire
startlabel = Label(frame0, text='\n\n\n\nQuestionnaire for Life Cycle Analysis of vertical farms\n\n\n', font=12)
startlabel.pack(fill=BOTH, side=BOTTOM)
my_image = PhotoImage(file = "avf logo.png") # your image
Label(frame0, image = my_image).pack(side=BOTTOM)

# Enter farm's name
frame0.pack(anchor=CENTER)
farm_name = StringVar()
Button(frame00, text='Next', command=next2).pack(fill=BOTH, side=BOTTOM, anchor=CENTER)
Entry(frame00, textvariable=farm_name).pack(fill=BOTH, side=BOTTOM, anchor=CENTER)
Label(frame00, text='\n\n\n\nEnter the name of your farm').pack(fill=BOTH, side=BOTTOM)

# Basic frame containing previous and next labels
button2 = Button(frame2, text=('Previous'), command=pre,padx=10)
button2.grid(row=0, column = 0, padx=10, sticky=W)
shitlabel = Label(frame2, text='                                   ').grid(row=0, column=1)
button1 = Button(frame2, text=('  Next  '), command=next1, padx=10)
button1.grid(row=0, column=2, sticky=E, padx=10)
root.bind('<Return>', enter)


# Define the 'file' Menu
root.file_opt = options = {}
options['defaultextension'] = '.txt'
options['filetypes'] = [('all files', '.*'), ('text files', '.txt')]
options['initialfile'] = 'myfile.txt'
options['parent'] = root
options['title'] = 'This is a title'

menu = Menu(root)
filemenu = Menu(menu, tearoff=0)
filemenu.add_command(label='Load', command=file_open)
filemenu.add_command(label='Save', command=file_save)
filemenu.add_command(label='Quit', command=root.quit)
menu.add_cascade(label='File', menu=filemenu)
root.config(menu=menu)

# Question 1: Where is your farm located?
# (If you are in this frame, you can't go back and change your name)
# Q1 needs to be specified here because pre and next are not initialized yet
v = IntVar()
var = StringVar()
var.set('1. In which country is your farm located?')
helloLabel = Label(frame1, textvariable=var).grid(row=0, column=0, padx=10, pady=10, sticky=W)
ans1 = StringVar()
country = ttk.Combobox(frame3, textvariable=ans1, state='readonly')
country['values'] = ('Netherlands', "Germany", "Switzerland", "United States", 'Japan')
country.current(0)
country.grid(padx=10)

# Here a list of all the possible crops a farmer can choose is read in. This is needed for Q2.
wb = xlrd.open_workbook('Database_full.xlsx')
lis = []
database = wb.sheet_by_name('Crop parameters')
for i in range(1, len(database.col_values(0))):
    if database.col_values(0)[i] == "":
        break
    lis.append(database.col_values(0)[i])

# Initialize variables to choose different crops in Q2
ansLet = IntVar()
ansEnd = IntVar()
ansSpi = IntVar()
ansBea = IntVar()
ansPar = IntVar()
ansKal = IntVar()
ansBas = IntVar()
ansRuc = IntVar()
ansMic = IntVar()
ansMin = IntVar()
ansVeg = [ansLet, ansEnd, ansSpi, ansBea, ansPar, ansKal, ansBas, ansRuc, ansMic, ansMin]

# Initialize variables for surface of a specific crop in Q2
surLet = IntVar()
surEnd = IntVar()
surSpi = IntVar()
surBea = IntVar()
surPar = IntVar()
surKal = IntVar()
surBas = IntVar()
surRuc = IntVar()
surMic = IntVar()
surMin = IntVar()
surVeg = [surLet, surEnd, surSpi, surBea, surPar, surKal, surBas, surRuc, surMic, surMin]

# Initialize variables for sold produce of a specific crop in Q2
kgLet = IntVar()
kgEnd = IntVar()
kgSpi = IntVar()
kgBea = IntVar()
kgPar = IntVar()
kgKal = IntVar()
kgBas = IntVar()
kgRuc = IntVar()
kgMic = IntVar()
kgMin = IntVar()
kgVeg = [kgLet, kgEnd, kgSpi, kgBea, kgPar, kgKal, kgBas, kgRuc, kgMic, kgMin]

Label(frame5, text='Crop [-]').grid(row=0, column=0, padx=5, sticky=W)
Label(frame5, text='Surface [m2]').grid(row=0, column=1, padx=5, sticky=W)
Label(frame5, text='Sold products [kg/year]').grid(row=0, column=2, padx=5, sticky=W)

# In this for loop, the fields for Q2 are created
for i in range(0, len(lis)):
    Checkbutton(frame5, text=lis[i], variable=ansVeg[i]).grid(row=i + 1, column=0, sticky=W, padx=5)
    EntSur = Entry(frame5, textvariable=surVeg[i])
    EntSur.grid(row=i + 1, column=1, sticky=W, padx=5)
    Entkg = Entry(frame5, textvariable=kgVeg[i])
    Entkg.grid(row=i + 1, column=2, sticky=W, padx=5)
    EntSur.bind('<FocusOut>', cal2)
    Entkg.bind('<FocusOut>', cal2)
    EntSur.bind("<Button-1>", rid_of_zeros_sur)
    Entkg.bind("<Button-1>", rid_of_zeros_kg)

# Here the fields for question 3 (buying electricity) are created
ans61 = IntVar()
ans62 = IntVar()
greenlabel = Label(frame8, text='Renewable').grid(row=1, column=0, padx=20, sticky=W)
greenentry = Entry(frame8, width=10, textvariable=ans61).grid(row=1, column=1)
greylabel = Label(frame8, text='Non-renewable').grid(row=2, column=0, padx=20, sticky=W)
greyentry = Entry(frame8, width=10, textvariable=ans62).grid(row=2, column=1)

# Here the fields for question 4 (creation of renewable energy) are created
ans71 = IntVar()
ans72 = IntVar()
ans73 = IntVar()
solarlabel = Label(frame9, text='Solar energy').grid(row=1, column=0, padx=20, sticky=W)
solarentry = Entry(frame9, width=10, textvariable=ans71).grid(row=1, column=1)
biomasslabel = Label(frame9, text='Biomass').grid(row=2, column=0, padx=20, sticky=W)
biomassentry = Entry(frame9, width=10, textvariable=ans72).grid(row=2, column=1)
windlabel = Label(frame9, text='Windpower').grid(row=3, column=0, padx=20, sticky=W)
windentry = Entry(frame9, width=10, textvariable=ans73).grid(row=3, column=1)

# Here the fields for Q5 (how electricity is used) are created
ans87 = IntVar()
ans88 = IntVar()
ans890 = IntVar()
Label(frame10, text='Selling renewable').grid(row=0, column=0, sticky=W, padx=5)
Entry(frame10, width=10, textvariable=ans87).grid(row=0, column=1)
Label(frame10, text='Selling non-renewable').grid(row=1, column=0, sticky=W, padx=5)
Entry(frame10, width=10, textvariable=ans88).grid(row=1, column=1)
Checkbutton(frame10, text='I don\'t know', variable=ans890).grid(row=3, column=0, sticky=W, padx=5)

# Here the fields for Q6 (fossil fuel use) are created
ans91 = IntVar()
ans92 = IntVar()
ans93 = IntVar()
ans94 = IntVar()
ans95 = IntVar()
ans96 = IntVar()
petroll = Label(frame11, text='Petrol (L)').grid(row=0, column=0, padx=5, sticky=W)
petroly = Entry(frame11, width=5, textvariable=ans91).grid(row=0, column=1)
diesell = Label(frame11, text='Diesel (L)').grid(row=1, column=0, padx=5, sticky=W)
diesely = Entry(frame11, width=5, textvariable=ans92).grid(row=1, column=1)
Ngasl = Label(frame11, text='Natural gas (M3)').grid(row=0, column=2, padx=10, sticky=W)
Ngasy = Entry(frame11, width=5, textvariable=ans94).grid(row=0, column=3)
oill = Label(frame11, text='Oil (L)').grid(row=1, column=2, padx=10, sticky=W)
oily = Entry(frame11, width=5, textvariable=ans95).grid(row=1, column=3)


# Here the field for fertilizer use are created (Q7)
ans121 = IntVar()
ans122 = IntVar()
ans123 = IntVar()
ans124 = IntVar()
ans125 = IntVar()
ans126 = IntVar()
ans127 = IntVar()
ans128 = IntVar()
ans131 = IntVar()
ans132 = IntVar()
ans133 = IntVar()
Label(frame14, text='Ammoniumnitrate').grid(row=1, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans121).grid(row=1, column=1)
Label(frame14, text='Calciumammoniumnitrate').grid(row=2, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans122).grid(row=2, column=1)
Label(frame14, text='Ammoniumsulphate').grid(row=3, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans123).grid(row=3, column=1)
Label(frame14, text='Triplesuperphosphate').grid(row=4, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans124).grid(row=4, column=1)
Label(frame14, text='Single super phosphate').grid(row=5, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans125).grid(row=5, column=1)
Label(frame14, text='Ammonia').grid(row=6, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans126).grid(row=6, column=1)
Label(frame14, text='Limestone').grid(row=7, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans127).grid(row=7, column=1)
Label(frame14, text='NPK 15-15-15').grid(row=8, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans128).grid(row=8, column=1)
Label(frame14, text='Phosphoric acid').grid(row=9, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans131).grid(row=9, column=1)
Label(frame14, text='Mono-ammonium phosphate').grid(row=10, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans132).grid(row=10, column=1)
Checkbutton(frame14, text='I don\'t know', variable=ans133).grid(row = 11, column = 0, padx=5, sticky=W)

# Here the fields for substrate use (Q8) are created
ans141 = IntVar()
ans142 = IntVar()
ans143 = IntVar()
ans144 = IntVar()
ans145 = IntVar()
ans146 = IntVar()
ans147 = IntVar()
Label(frame16, text='Rockwool').grid(row=1, column=0, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans141).grid(row=1, column=1)
Label(frame16, text='Perlite').grid(row=2, column=0, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans142).grid(row=2, column=1)
Label(frame16, text='Cocofiber').grid(row=1, column=2, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans143).grid(row=1, column=3)
Label(frame16, text='Hemp fiber').grid(row=2, column=2, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans144).grid(row=2, column=3)
Label(frame16, text='Peat').grid(row=3, column=0, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans145).grid(row=3, column=1)
Label(frame16, text='Peat Moss').grid(row=3, column=2, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans146).grid(row=3, column=3)
Checkbutton(frame160, text='No substrate is used', variable=ans147).grid(padx=5)

# Here the fields for water use (Q9) are created
ans151 = IntVar()
ans152 = IntVar()
Label(frame17, text='Tap water').grid(row=1, column=0, padx=5, sticky=W)
Entry(frame17, width=10, textvariable=ans151).grid(row=1, column=1)
Checkbutton(frame170, text='I don\'t know', variable=ans152).grid(sticky=W, padx=5)

# Here the fields for pesticide use (Q10) are created
ans161 = IntVar()
ans162 = IntVar()
ans163 = IntVar()
ans164 = IntVar()
ans165 = IntVar()
ans166 = IntVar()
Label(frame18, text='Atrazine').grid(row=1, column=0, padx=5, sticky=W)
Entry(frame18, width=10, textvariable=ans161).grid(row=1, column=1)
Label(frame18, text='Glyphosphate').grid(row=2, column=0, padx=5, sticky=W)
Entry(frame18, width=10, textvariable=ans162).grid(row=2, column=1)
Label(frame18, text='Metolachlor').grid(row=3, column=0, padx=5, sticky=W)
Entry(frame18, width=10, textvariable=ans163).grid(row=3, column=1)
Label(frame18, text='Herbicide').grid(row=4, column=0, padx=5, sticky=W)
Entry(frame18, width=10, textvariable=ans164).grid(row=4, column=1)
Label(frame18, text='Insectiside').grid(row=5, column=0, padx=5, sticky=W)
Entry(frame18, width=10, textvariable=ans165).grid(row=5, column=1)
Checkbutton(frame180, text='I don\'t know', variable=ans166).grid(sticky=W, padx=5)

# Here the fields for packaging (Q11) are created
ans171 = IntVar()
ans172 = IntVar()
ans173 = IntVar()
Radiobutton(frame19, text='Yes, it is', variable=ans171, value=1).grid(sticky=W, padx=5)
Radiobutton(frame19, text='No, it isn\'t', variable=ans171, value=2).grid(sticky=W, padx=5)

# Here the fields for transportation (Q13)are created
ans201 = IntVar()
ans202 = IntVar()
ans203 = IntVar()
ans204 = IntVar()
Label(frame20, text='Van').grid(row=1, column=0, padx=40, sticky=W)
Entry(frame20, width=10, textvariable=ans201).grid(row=1, column=1)
Label(frame20, text='Truck').grid(row=2, column=0, padx=40, sticky=W)
Entry(frame20, width=10, textvariable=ans202).grid(row=2, column=1)
Checkbutton(frame200, text='I don\'t know', variable=ans204).grid(sticky=W, padx=40)

# Here fields for finishing the questionnaire are created
Button_finish = Button(frame_finish, text=('Finish!'), command=close_program, padx = 10)
Button_finish.grid(row=1, column=1, padx=10, sticky = E)


# At the end, a list containing all the variables is created. It is needed to be able to load previously filled in results
list_ans = [farm_name, ans1, v, ans61, ans62, ans71, ans72, ans73, ans87,
            ans88, ans890, ans91, ans92, ans94, ans95, ans121, ans122, ans123, ans124, ans125, ans126, ans127, ans128, ans131, ans132, ans133, ans141, ans142, ans143, ans144, ans145, ans146, ans147, ans151, ans152,
            ans161, ans162, ans163, ans164, ans165, ans166, ans171, ans172, ans201,
            ans202, ans204]

# Important statement. If not placed here, program crashes. Assures that all information from above is in the program
root.mainloop()
