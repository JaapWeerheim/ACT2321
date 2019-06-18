from tkinter import *
from tkinter import ttk
import xlrd
import xlsxwriter
import tkinter.filedialog

# Building the Gui including all the frames
root = Tk()
root.title('VertiCal')
root.wm_iconbitmap('sfsf logo.ico')
root.geometry('440x440+500+200')

# Setting heights and widths for every frame.
frame_start = Frame(height=65, width=400)
frame_farm_name = Frame(height=65, width=400)
frame_location = Frame(height=75, width=400)
frame_location_extension = Frame(height=120, width=400)
frame_previous_next = Frame(height=40, width=400)
frame_crop_species = Frame(height=8000, width=400)
frame_buy_energy = Frame(height=120, width=400)
frame_create_renewable = Frame(height=120, width=400)
frame_sell_renewable = Frame(height=90, width=400)
frame_fuel_use = Frame(height=70, width=400)
frame_fertilizer_use = Frame(height=290, width=400)
frame_substrate_use = Frame(height=90, width=400)
frame_water_use = Frame(height=50, width=400)
frame_pesticide_use = Frame(height=140, width=400)
frame_packaging_use = Frame(height=70, width=400)
frame_transport = Frame(height=120, width=400)
frame_finish = Frame(height=120, width=400)

all_frames = [frame_start, frame_farm_name, frame_location, frame_previous_next, frame_location_extension,
              frame_crop_species, frame_buy_energy, frame_create_renewable, frame_sell_renewable, frame_fuel_use,
              frame_fertilizer_use, frame_substrate_use, frame_water_use, frame_pesticide_use, frame_packaging_use,
              frame_transport, frame_finish]

# Set all the frames to a certain size
for frame in all_frames:
    frame.grid_propagate(0)

# ----------------------------------------
# Below, all functions of the program are created

# Function worksheetoutput does all the calculations to perform the LCA and writes the result to an Excel sheet
def worksheetoutput(dictionary_name):
    # Opening excel file in order to get parameters
    workbook = xlrd.open_workbook('Database_full.xlsx')
    dicp = {}
    for tabs in workbook.sheet_names():
        if tabs != 'Crop parameters':
            sheet = workbook.sheet_by_name(tabs)
            for i in range(1, sheet.nrows):
                if sheet.cell_value(i, 2) != '':
                    # print (sheet.cell_value(i,2))
                    for j in range(0, sheet.ncols):
                        if sheet.cell_value(0,j) == ans_country.get():
                            country = j
                            if sheet.cell_value(i, j) == '':
                                country = 9  # 9 is world
                            dicp[sheet.cell_value(i, 2)] = sheet.cell_value(i, country)
                            if ans_packaging.get() == 0:
                                dicp['Pac1'] = 0
                                dicp['Pac2'] = 0
    non_count = str()
    nr_dont_know = 0
    # If choose 'I don't know’ option, set the value back to zero
    if ans_check_buy_energy.get() == 1:
        ans_buy_renew.set(0); ans_buy_nonrenew.set(0)
        non_count = ('bought electricity, ')
        nr_dont_know += 1
    if ans_check_create_renewable.get() == 1:
        ans_prod_biomass.set(0); ans_prod_solar.set(0); ans_prod_wind.set(0);
        non_count = (non_count + 'renewable energy production, ')
        nr_dont_know += 1
    if ans_check_sell_energy.get() == 1:
        ans_sel_renew.set(0); ans_sel_non_renew.set(0)
        non_count = (non_count + 'sold electricity, ')
        nr_dont_know += 1
    if ans_check_fossil_fuel_use.get() == 1:
        ans_oil_use.set(0); ans_natural_gas_use.set(0); ans_diesel_use.set(0); ans_petrol_use.set(0)
        non_count = (non_count + 'fossil fuel use, ')
        nr_dont_know += 1
    if ans_check_fertilizer_use.get() == 1:
        ans_ammonium_nitrate_use.set(0); ans_calcium_ammonium_nitrate_use.set(0); ans_ammonium_sulphate_use.set(0);
        ans_triple_super_phosphate_use.set(0); ans_single_super_phosphate_use.set(0); ans_ammonia_use.set(0);
        ans_limestone_use.set(0); ans_NPK_151515_use.set(0); ans_phosphoric_acid_use.set(0); ans_mono_ammonium_phosphate_use.set(0)
        non_count = (non_count + 'NPK chemicals, ')
        nr_dont_know += 1
    if ans_check_substrate_use.get() == 1:
        ans_rockwool_use.set(0); ans_perlite_use.set(0); ans_cocofiber_use.set(0); ans_hempfiber_use.set(0);
        ans_peat_use.set(0); ans_peatmoss_use.set(0)
        non_count = (non_count + 'substrate, ')
        nr_dont_know += 1
    if ans_check_tap_water_use.get() == 1:
        ans_tap_water_use.set(0)
        non_count = (non_count + 'water, ')
        nr_dont_know += 1
    if ans_check_pesticide_use.get() == 1:
        ans_atrazine_use.set(0); ans_glyphosphate_use.set(0); ans_metolachlor_use.set(0); ans_herbicide_use.set(0);
        ans_insecticide_use.set(0)
        non_count = (non_count + 'pesticides, ')
        nr_dont_know += 1
    if ans_check_transport.get() == 1:
        ans_van_use.set(0); ans_truck_use.set(0);
        #ans_percentage_truck_use.set(0); ans_percentage_van_use.set(0);
        non_count = (non_count + 'transport')
        nr_dont_know += 1

    # Create the output: an Excel file
    wb = xlsxwriter.Workbook(farm_name.get() + '.xlsx')

    sheet = workbook.sheet_by_name('Crop parameters')
    Total_Eoc = 0
    total_growth_cycles = 0
    for keys, values in dictionary_name.items():
        for i in range(1, len(list_crop_species) + 1):
            if keys == sheet.cell_value(i, 0):
                # Energy content
                dictionary_name[keys] += [sheet.cell_value(i, 4)]
                Total_Eoc += sheet.cell_value(i, 4)
                # Growth period
                dictionary_name[keys] += [365/sheet.cell_value(i, 5)]
                total_growth_cycles += (365/sheet.cell_value(i, 5))
    Average_Eoc = Total_Eoc / (len(dictionary_name) - 1)
    average_growth_period = total_growth_cycles / (len(dictionary_name)-1)
    dictionary_name[list(dictionary_name.keys())[0]] += [Average_Eoc]
    dictionary_name[list(dictionary_name.keys())[0]] += [average_growth_period]


    # Calculation to find out the total combination of fraction surface and fraction growth period, needed for the substrate
    # Take care: this for loop cannot be combined with the for loop behind it since it needs to run through this whole loop,
    # to find the final sum_frac in order to properly do the calculations on substrate in the next for loop.
    sum_frac = 0
    for keys, values in dictionary_name.items():
        if keys != 'Total':
            frac_surf = values[0]
            growth_cycles = values[4]
            # Substrate calculations
            frac_growth = growth_cycles / total_growth_cycles
            frac_surf_growth = frac_growth * frac_surf
            sum_frac += frac_surf_growth

    # doing the calculations
    for keys, values in dictionary_name.items():
        cropname = keys
        frac_surf = values[0]
        frac_kg = values[1]
        kg_prod = values[2]
        Eoc = values[3]
        growth_cycles = values[4]

        # Calculation for total C02 of electricity usage
        Eco2 = frac_surf * ((dicp['C1'] * ans_buy_renew.get()) + (dicp['C3'] * ans_buy_nonrenew.get()) + (dicp['C5'] * ans_prod_solar.get()) + (dicp['C7'] * ans_prod_wind.get()) + (
                dicp['C9'] * ans_prod_biomass.get()) - (ans_sel_renew.get() * dicp['C1']) - (ans_sel_non_renew.get() * dicp['C3']))

        # Calculation for total energy of electricity usage
        Eenergy = frac_surf * ((dicp['C2'] * ans_buy_renew.get()) + (dicp['C4'] * ans_buy_nonrenew.get()) + (dicp['C6'] * ans_prod_solar.get()) + (dicp['C8'] * ans_prod_wind.get()) + (
                dicp['C10'] * ans_prod_biomass.get()) - (ans_sel_renew.get() * dicp['C2']) - (ans_sel_non_renew.get() * dicp['C4']))

        # Calculation for total Co2 of fossil fuels use
        Fco2 = frac_surf * ((dicp['Fo1'] * ans_petrol_use.get()) + (dicp['Fo3'] * ans_diesel_use.get()) + (dicp['Fo7'] * ans_natural_gas_use.get()) + (dicp['Fo9'] * ans_oil_use.get()))

        # Calculation for total energy of fossil fuel use
        Fenergy = frac_surf * (
                (dicp['Fo2'] * ans_petrol_use.get()) + (dicp['Fo4'] * ans_diesel_use.get()) + (dicp['Fo8'] * ans_natural_gas_use.get()) + (dicp['Fo10'] * ans_oil_use.get()))

        # Calculation for total Co2 of fertilizers
        FERco2 = frac_surf * ((
                (dicp['Fe1'] * ans_ammonium_nitrate_use.get()) + (dicp['Fe3'] * ans_calcium_ammonium_nitrate_use.get()) + (dicp['Fe5'] * ans_ammonium_sulphate_use.get()) + (dicp['Fe7'] * ans_triple_super_phosphate_use.get()) + (
                dicp['Fe9'] * ans_single_super_phosphate_use.get()) + (dicp['Fe11'] * ans_ammonia_use.get()) + (dicp['Fe13'] * ans_limestone_use.get()) + (
                        dicp['Fe15'] * ans_NPK_151515_use.get()) + (dicp['Fe21'] * ans_phosphoric_acid_use.get()) + (dicp['Fe22'] * ans_mono_ammonium_phosphate_use.get())))

        # Calculation for total energy of fertilizers
        FERenergy = frac_surf * ((
                (dicp['Fe2'] * ans_ammonium_nitrate_use.get()) + (dicp['Fe4'] * ans_calcium_ammonium_nitrate_use.get()) + (dicp['Fe6'] * ans_ammonium_sulphate_use.get()) + (dicp['Fe8'] * ans_triple_super_phosphate_use.get()) + (
                dicp['Fe10'] * ans_single_super_phosphate_use.get()) + (dicp['Fe12'] * ans_ammonia_use.get()) + (dicp['Fe14'] * ans_limestone_use.get()) + (
                        dicp['Fe16'] * ans_NPK_151515_use.get()) + (dicp['Fe22'] * ans_phosphoric_acid_use.get()) + (dicp['Fe24'] * ans_mono_ammonium_phosphate_use.get())))

        # Calculation in which growth period and fraction of surface are combined into one fraction for substrate calculation
        if keys != 'Total':
            frac_growth = growth_cycles / total_growth_cycles
            frac_surf_growth = frac_growth * frac_surf
            frac_substrate = frac_surf_growth/sum_frac
        else:
            # frac_substrate for total just need to be one
            frac_substrate = 1


        # Calculation for total Co2 of substrates
        Sco2 = frac_substrate * (
                (dicp['S1'] * ans_rockwool_use.get()) + (dicp['S3'] * ans_perlite_use.get()) + (dicp['S5'] * ans_cocofiber_use.get()) + (dicp['S7'] * ans_hempfiber_use.get()) + (
                dicp['S9'] * ans_peat_use.get()) + (dicp['S11'] * ans_peatmoss_use.get()))

        # Calculation for total energy of substrates
        Senergy = frac_substrate * (
                (dicp['S2'] * ans_rockwool_use.get()) + (dicp['S4'] * ans_perlite_use.get()) + (dicp['S6'] * ans_cocofiber_use.get()) + (dicp['S8'] * ans_hempfiber_use.get()) + (
                dicp['S10'] * ans_peat_use.get()) + (dicp['S12'] * ans_peatmoss_use.get()))

        # Calculation for total Co2 of water
        Wco2 = frac_surf * (dicp['Wa1'] * ans_tap_water_use.get())

        # Calculation for total energy of water
        Wenergy = frac_surf * (dicp['Wa2'] * ans_tap_water_use.get())

        # Calculation for total Co2 of pesticides
        Pco2 = frac_surf * (
                (dicp['P1'] * ans_atrazine_use.get()) + (dicp['P3'] * ans_glyphosphate_use.get()) + (dicp['P5'] * ans_metolachlor_use.get()) + (dicp['P7'] * ans_herbicide_use.get()) + (
                dicp['P9'] * ans_insecticide_use.get()))

        # Calculation for total energy of pesticides
        Penergy = frac_surf * (
                (dicp['P2'] * ans_atrazine_use.get()) + (dicp['P4'] * ans_glyphosphate_use.get()) + +(dicp['P6'] * ans_metolachlor_use.get()) + +(dicp['P8'] * ans_herbicide_use.get()) + (
                dicp['P10'] * ans_insecticide_use.get()))

        # Scaling the percentages of transportation means
        if ans_percentage_van_use.get() or ans_percentage_truck_use.get() > 0:
            truck_use_percent = ans_percentage_truck_use.get()/(ans_percentage_truck_use.get() + ans_percentage_van_use.get())
            van_use_percent = ans_percentage_van_use.get()/(ans_van_use.get() + ans_percentage_truck_use.get())
        else:
            truck_use_percent = 50
            van_use_percent = 50

        # Calculation for total Co2 of transport
        Tco2 = kg_prod * ((dicp['T3'] * ans_van_use.get() * van_use_percent * van_owner()) + (dicp['T1'] * ans_truck_use.get() * truck_use_percent * truck_owner()))

        # Calculation for total energy of transport
        Tenergy = kg_prod * ((dicp['T4'] * ans_van_use.get() * van_use_percent * van_owner()) + (dicp['T2'] * ans_truck_use.get() * truck_use_percent * truck_owner()))

        # Calculation for the total Co2 of packaging
        Pacco2 = kg_prod * dicp['Pac1']

        # Calculation for the total energy of packaging
        Pacenergy = kg_prod * dicp['Pac2']

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
        cell_format_bold = wb.add_format({'bold': True,
                                          'align': 'right'})
        cell_format_header = wb.add_format({'bold': True,
                                            'font_size': 16,
                                            'align': 'center',
                                            'fg_color': '#cdcdcd'})
        ws.merge_range('B1:C1', 'CO\u2082 emitted', cell_format_header)
        ws.write(1, 1, 'Total [kg]', cell_format_bold)
        ws.write(1, 2, 'Per kg crop [kg/kg]', cell_format_bold)
        ws.merge_range('D1:E1', 'Energy used', cell_format_header)
        ws.write(1, 3, 'Total [MJ]', cell_format_bold)
        ws.write(1, 4, 'Per kg crop [kg/kg]', cell_format_bold)
        ws.set_column(2, 1, len('Per kg crop [kg/kg]'))
        ws.set_column(3, 2, len('Per kg crop [kg/kg]'))
        ws.set_column(4, 3, len('Per kg crop [kg/kg]'))
        ws.set_column(5, 4, len('Per kg crop [kg/kg]'))
        # also labels in Dutch/other languages?
        labels_output = ['Electricity', 'Fossil fuels', 'Fertilizer', 'Substrates', 'Water', 'Pesticides',
                         'Transport', 'Package']
        Co2_emitted = [Eco2, Fco2, FERco2, Sco2, Wco2, Pco2, Tco2, Pacco2]
        Co2_emitted_round = [round(elem, 0) for elem in Co2_emitted]
        energy_used = [Eenergy, Fenergy, FERenergy, Senergy, Wenergy, Penergy, Tenergy, Pacenergy]
        energy_used_round = [round(elem, 0) for elem in energy_used]
        Co2_crop = []
        energy_crop = []
        sum_co2_per_crop = 0
        sum_energy_per_crop = 0
        for i in range(len(Co2_emitted)):
            Co2_crop += [Co2_emitted[i] / dic_crops[cropname][2]]
            energy_crop += [energy_used[i] / dic_crops[cropname][2]]
            Co2_crop_round = [round(elem, 3) for elem in Co2_crop]
            energy_crop_round = [round(elem, 3) for elem in energy_crop]
            sum_co2_per_crop += Co2_emitted[i] / dic_crops[cropname][2]
            sum_energy_per_crop += energy_used_round[i] / dic_crops[cropname][2]
        cell_format_background1 = wb.add_format({'bg_color': '#cdcdcd'})
        cell_format_background2 = wb.add_format({'bg_color': '#ffffff'})
        cell_formats = [cell_format_background1, cell_format_background2]
        cell_format_border = wb.add_format({'right': 1})
        for x in range(len(labels_output)):
            if x % 2 == 0:
                i = 1
            else:
                i = 0
            ws.write(2 + x, 0, labels_output[x], cell_format_bold)
            ws.write(2 + x, 1, Co2_emitted_round[x], cell_formats[i])
            ws.write(2 + x, 2, Co2_crop_round[x], cell_formats[i])
            ws.write(2 + x, 3, energy_used_round[x], cell_formats[i])
            ws.write(2 + x, 4, energy_crop_round[x], cell_formats[i])

        cell_format_bottom = wb.add_format({'bottom': 1})
        ws.write(3 + len(labels_output), 0, 'Total', cell_format_bold)
        ws.write(3 + len(labels_output), 1, round(sum(Co2_emitted), 0), cell_format_bottom)
        ws.write(3 + len(labels_output), 2, round(sum(Co2_crop), 3), cell_format_bottom)
        ws.write(3 + len(labels_output), 3, round(sum(energy_used), 0), cell_format_bottom)
        ws.write(3 + len(labels_output), 4, round(sum(energy_crop), 3), cell_format_bottom)
        ws.set_column(0, 0, len('Fossil fuels'))

        labels_total = ['Total CO\u2082 emitted [kg per year]', 'Total energy used [MJ per year]',
                        'Total CO\u2082 emitted per kg product [kg/kg per year]',
                        'Total Energy used per kg product [KJ/Kg per year]',
                        'Total CO\u2082 emitted per KJ product [kg/KJ per year]',
                        'Total energy used per KJ product [KJ/KJ per year]']
        totals_output = [Totalco2, Totalenergy, Totalco2_per_kg_product, Totalenergy_per_kg_product,
                         Totalco2_per_KJ_product, Totalenergy_per_KJ_product]
        totals_output_round = [round(elem, 2) for elem in totals_output]
        for x in range(len(labels_total)):
            ws.write(1 + x, 6, labels_total[x])
            ws.write(1 + x, 7, totals_output_round[x])
        ws.set_column(6, 6, len('Total energy used per KJ product [KJ/KJ per year]'))

        if ans_check_buy_energy.get() == 1 or ans_check_create_renewable.get() == 1 or ans_check_sell_energy.get() == 1 \
                or ans_check_fossil_fuel_use.get() == 1 or ans_check_fertilizer_use.get() == 1 or \
                ans_check_substrate_use.get() == 1 or ans_check_tap_water_use.get() == 1 or \
                ans_check_pesticide_use.get() == 1 or ans_check_transport.get() == 1:
            if nr_dont_know <= 4:
                ws.write(10, 1,
                         "Specifications of " + non_count + ' are not taken into account because of lacking data.')
            else:
                warning_format = wb.add_format({'bold': True, 'font_size': 16})
                ws.write(10, 1,
                         "You have used the 'I don't know' button too often. The analysis is missing too much data to show significant results. Please try again.",
                         warning_format)

        # Creating bar charts
        chart_col = wb.add_chart({'type': 'column'})
        chart_col.add_series({
            'name': [cropname, 0, 1],
            'categories': [cropname, 2, 0, 9, 0],
            'values': [cropname, 2, 2, 9, 2],
            'fill': {'color': 'black'}
        })
        chart_col.set_title({'name': 'Total CO\u2082 emitted from different sources',
                             'name_font': {'size': 12}})
        chart_col.set_y_axis({'name': 'Total CO\u2082 emitted',
                              'major_gridlines': {
                                  'visible': False
                              }})
        chart_col.set_x_axis({'name': 'Sources'})
        ws.insert_chart('A13', chart_col, {'x_offset': 20, 'y_offset': 8})

        chart_col = wb.add_chart({'type': 'column'})
        chart_col.add_series({
            'name': [cropname, 0, 3],
            'categories': [cropname, 2, 0, 9, 0],
            'values': [cropname, 2, 3, 9, 3],
            'fill': {'color': 'black'}
        })
        chart_col.set_title({'name': 'Total energy used from different sources',
                             'name_font': {'size': 12}})
        chart_col.set_y_axis({'name': 'Total energy used',
                              'major_gridlines': {
                                  'visible': False
                              }})
        chart_col.set_x_axis({'name': 'Sources'})

        ws.insert_chart('E13', chart_col, {'x_offset': 20, 'y_offset': 8})

        if cropname == 'Total':
            chart_col = wb.add_chart({'type': 'column'})
            chart_col.add_series({
                'name': [cropname, 0, 1],
                'categories': [cropname, 2, 0, 9, 0],
                'values': [cropname, 2, 1, 9, 1],
                'fill': {'color': 'black'}
            })
            chart_col.set_title({'name': 'Total CO\u2082 emitted from different sources',
                                 'name_font': {'size': 12}})
            chart_col.set_y_axis({'name': 'Total CO\u2082 emitted',
                                  'major_gridlines': {
                                      'visible': False
                                  }})
            chart_col.set_x_axis({'name': 'Sources', })

            ws.insert_chart('A13', chart_col, {'x_offset': 20, 'y_offset': 8})

            chart_col = wb.add_chart({'type': 'column'})
            chart_col.add_series({
                'name': [cropname, 0, 4],
                'categories': [cropname, 1, 0, 10, 0],
                'values': [cropname, 1, 3, 10, 3],
                'fill': {'color': 'black'}
            })
            chart_col.set_title({'name': 'Total energy used from different sources',
                                 'name_font': {'size': 12}})
            chart_col.set_y_axis({'name': 'Total energy used',
                                  'major_gridlines': {
                                      'visible': False
                                  }})
            chart_col.set_x_axis({'name': 'Sources'})

            chart_co2 = wb.add_chart({'type': 'column'})
            chart_co2.set_title({'name': 'CO\u2082 emitted per kg crop',
                                 'name_font': {'size': 12}})
            chart_energy = wb.add_chart({'type': 'column'})
            chart_energy.set_title({'name': 'Energy used per kg crop',
                                    'name_font': {'size': 12}})
            headings = ['#ffffff', '#000000', '#cdcdcd', '#373737', '#828282', '#505050', '#e6e6e6', '#1e1e1e',
                        '#b4b4b4', '#696969', '#9b9b9b']
            count = 0
            for keys, values in dictionary_name.items():
                if keys != 'Total':
                    chart_co2.add_series({'name': keys,
                                          'values': [keys, 2, 2, 9, 2],
                                          'categories': [keys, 2, 0, 9, 0],
                                          'fill': {'color': headings[count]},
                                          'border': {'color': 'black'}
                                          })
                    chart_energy.add_series({'name': keys,
                                             'values': [keys, 2, 4, 9, 4],
                                             'categories': [keys, 2, 0, 9, 0],
                                             'fill': {'color': headings[count]},
                                             'border': {'color': 'black'}
                                             })
                    count += 1
            chart_co2.set_y_axis({'name': 'Co\u2082 used',
                                  'major_gridlines': {
                                      'visible': False
                                  }})
            chart_co2.set_x_axis({'name': 'Sources'})
            chart_energy.set_y_axis({'name': 'Energy used',
                                     'major_gridlines': {
                                         'visible': False
                                     }})
            chart_energy.set_x_axis({'name': 'Sources'})
            # chart_co2.set_size({'width': 960, 'height': 285})
            ws.insert_chart('A28', chart_co2, {'x_offset': 20, 'y_offset': 8})
            ws.insert_chart('E28', chart_energy, {'x_offset': 20, 'y_offset': 8})

    # Write the raw data from the questionnaire to the Excel sheet
    ws = wb.add_worksheet("Raw data")

    # Write question 1
    ws.write(0,0,  "Question 1")
    ws.write(1, 0, 'Country:')
    ws.write(1, 1, ans_country.get())

    # Write question 2
    ws.write(3,0, "Question 2")
    ws.write(4,0, "Crop type")
    ws.write(4,1, 'Surface [m2]')
    ws.write(4,2, "Sold products [kg/year]")
    for i in range(0,len(ansVeg)):
        ws.write(5+i, 0, list_crop_species[i])
        ws.write(5+i, 1, surVeg[i].get())
        ws.write(5+i, 2, kgVeg[i].get())

    # Write question 3, 4 and 5
    ws.write(16,0, "Question 3-5")
    ws.write(17,0, "Electricity type")
    ws.write(17,1, "Amount [kWh/year]")
    list_electricity = [ans_buy_renew.get(), ans_buy_nonrenew.get(), ans_prod_solar.get(), ans_prod_biomass.get(),
                        ans_prod_wind.get(), ans_sel_renew.get(), ans_sel_non_renew.get()]
    list_electricity_names = ["Bought renewable", "Bought non-renewable", "Produced solar",
                              "Produced biomass", "Produced wind", "Sold renewable",
                              "Sold non-renewable"]
    for i in range(0,len(list_electricity)):
        ws.write(18+i, 0, list_electricity_names[i])
        ws.write(18+i, 1, list_electricity[i])

    # Question 6
    ws.write(26, 0, "Question 6")
    ws.write(27, 0, "Fossil fuel type")
    ws.write(27,1, "Consumption [per year]")
    list_fuel = [ans_petrol_use.get(), ans_diesel_use.get(), ans_natural_gas_use.get(), ans_oil_use.get()]
    list_fuel_names = ["Petrol (L)", "Diesel (L)", "Oil (L)", "Natural gas (m3)"]
    for i in range(0, len(list_fuel)):
        ws.write(28+i, 0, list_fuel_names[i])
        ws.write(28+i, 1, list_fuel[i])

    # Question 7
    ws.write(33,0, "Question 7")
    ws.write(34,0, "Fertilizer type")
    ws.write(34,1, "Consumption [kg/year]")
    list_fertilizers = [ans_ammonium_nitrate_use.get(), ans_calcium_ammonium_nitrate_use.get(),
                        ans_ammonium_sulphate_use.get(),
                        ans_triple_super_phosphate_use.get(), ans_single_super_phosphate_use.get(),
                        ans_ammonia_use.get(),
                        ans_limestone_use.get(), ans_NPK_151515_use.get(), ans_phosphoric_acid_use.get(),
                        ans_mono_ammonium_phosphate_use.get()]

    list_fertilizer_names = ["Ammonium nitrate", "Calcium ammonium nitrate", "Ammonium sulphate",
                             "Triple super phosphate",
                             "Single super phosphate", "Ammonia", "Limestone", "NPK 15-15-15", "Phosphoric acid",
                             "Mono-ammonium phosphate"]
    for i in range(0, len(list_fertilizers)):
        ws.write(35+i, 0, list_fertilizer_names[i])
        ws.write(35+i, 1, list_fertilizers[i])

    # Question 8
    ws.write(46,0, "Question 8")
    ws.write(47,0, "Substrate type")
    ws.write(47,1, "Consumption [kg/year]")
    list_substrates = [ans_rockwool_use.get(), ans_perlite_use.get(), ans_cocofiber_use.get(), ans_hempfiber_use.get(),
                       ans_peat_use.get(), ans_peatmoss_use.get()]
    list_substrates_names = ["Rockwool", "Perlite", "Cocofiber", "Hempfiber", 'Peat', "Peatmoss"]
    for i in range(0, len(list_substrates)):
        ws.write(48+i, 0, list_substrates_names[i])
        ws.write(48+i, 1, list_substrates[i])

    # Question 9
    ws.write(55,0, "Question 9")
    ws.write(56,0, "Water consumption:")
    ws.write(56,1, ans_tap_water_use.get())
    ws.write(56,2 ,"[l/year]")

    # Question 10
    ws.write(58,0, "Question 10")
    ws.write(59,0, "Pesticide type")
    ws.write(59,1, "Consumption [kg/year]")
    list_pesticides = [ans_atrazine_use.get(), ans_glyphosphate_use.get(),
                       ans_metolachlor_use.get(), ans_herbicide_use.get(), ans_insecticide_use.get()]
    list_pesticides_names = ["Atrazine", "Glyphosphate", "Metolachlore", "Herbicide", "Insecticide"]
    for i in range(0, len(list_pesticides)):
        ws.write(60+i, 0, list_pesticides_names[i])
        ws.write(60+i, 1, list_pesticides[i])

    # Question 11
    ws.write(66,0, "Question 11")
    ws.write(67,0, "Packaged [Yes/No]")
    if ans_packaging.get == 0:
        ws.write(67,1, "No")
    else:
        ws.write(67,1, "Yes")

    # Question 12
    ws.write(69,0, "Question 12")
    ws.write(70,0, "Transportation means")
    ws.write(70,1, "Average distance [km]")
    ws.write(71, 0, "Van")
    ws.write(71,1, ans_van_use.get())
    ws.write(72,0, "Truck")
    ws.write(72,1, ans_truck_use.get())
    ws.write(70,2, "Percentage of products [%]")
    ws.write(71,2, ans_percentage_van_use.get())
    ws.write(72,2, ans_percentage_truck_use.get())
    ws.write(70,3, "Owner")
    ws.write(71,3, ans_van_own.get())
    ws.write(72,3, ans_truck_own.get())

    # Close the workbook again
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
        var.set(question_location)
        frame_crop_species.grid_forget()
        frame_location_extension.grid(sticky=W)
    if count == 2:
        var.set(question_crop_types)
        frame_buy_energy.grid_forget()
        frame_crop_species.grid(sticky=W)
    if count == 3:
        var.set(question_buy_renewable)
        frame_create_renewable.grid_forget()
        frame_buy_energy.grid(sticky=W)
    if count == 4:
        var.set(question_produce_renewable)
        frame_sell_renewable.grid_forget()
        frame_create_renewable.grid(sticky=W)
    if count == 5:
        var.set(question_sell_electricity)
        frame_fuel_use.grid_forget()
        frame_sell_renewable.grid(sticky=W)
    if count == 6:
        var.set(question_fossil_fuel_use)
        frame_fertilizer_use.grid_forget()
        frame_fuel_use.grid(sticky=W)
    if count == 7:
        var.set(question_npk_use)
        frame_substrate_use.grid_forget()
        frame_fertilizer_use.grid(sticky=W)
    if count == 8:
        var.set(question_substrate_use)
        frame_water_use.grid_forget()
        frame_substrate_use.grid(sticky=W)
    if count == 9:
        var.set(question_water_use)
        frame_pesticide_use.grid_forget()
        frame_water_use.grid(sticky=W)
    if count == 10:
        var.set(question_pesticide_use)
        frame_packaging_use.grid_forget()
        frame_pesticide_use.grid(sticky=W)
    if count == 11:
        var.set(question_packaging_use)
        frame_transport.grid_forget()
        frame_packaging_use.grid(sticky=W)
    if count == 12:
        var.set(question_transport)
        frame_finish.grid_forget()
        frame_transport.grid(sticky=W)
    if count == 13:
        var.set(question_finish)
        frame_finish.grid(sticky=W)
    if count > 12:
        count = 13
    return


# def next1() enables to go to the next question.
# i.e. forgetting the current frames and introducing new frames

def next1():
    global count
    for i in range(len(list_ans)):  # if there is no value in Entry, set it back to 0
        try:
            if i != 0 or 2 or 1:
                list_ans[i].get() != ''
        except TclError:
            list_ans[i].set(00)
    count += 1
    if count == 2:
        var.set(question_crop_types)
        frame_location_extension.grid_forget()
        frame_crop_species.grid(sticky=W)
    if count == 3:
        var.set(question_buy_renewable)
        frame_crop_species.grid_forget()
        frame_buy_energy.grid(sticky=W)
    if count == 4:
        var.set(question_produce_renewable)
        frame_buy_energy.grid_forget()
        frame_create_renewable.grid(sticky=W)
    if count == 5:
        var.set(question_sell_electricity)
        frame_create_renewable.grid_forget()
        frame_sell_renewable.grid(sticky=W)
    if count == 6:
        var.set(question_fossil_fuel_use)
        frame_sell_renewable.grid_forget()
        frame_fuel_use.grid(sticky=W)
    if count == 7:
        var.set(question_npk_use)
        frame_fuel_use.grid_forget()
        frame_fertilizer_use.grid(sticky=W)
    if count == 8:
        var.set(question_substrate_use)
        frame_fertilizer_use.grid_forget()
        frame_substrate_use.grid(sticky=W)
    if count == 9:
        var.set(question_water_use)
        frame_substrate_use.grid_forget()
        frame_water_use.grid(sticky=W)
    if count == 10:
        var.set(question_pesticide_use)
        frame_water_use.grid_forget()
        frame_pesticide_use.grid(sticky=W)
    if count == 11:
        var.set(question_packaging_use)
        frame_pesticide_use.grid_forget()
        frame_packaging_use.grid(sticky=W)
    if count == 12:
        var.set(question_transport)
        frame_packaging_use.grid_forget()
        frame_transport.grid(sticky=W)
    if count == 13:
        var.set(question_finish)
        frame_transport.grid_forget()
        frame_finish.grid(sticky=W)
    if count > 12:
        count = 13
    return


# Closes the program
def quit1():
    root.destroy()
    return

def close_program():
    cal2()
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
    frame_start.pack_forget()
    frame_farm_name.pack(anchor=CENTER)

    return


# The command of 'Next' button after you input the farm name
def next2():
    global count
    frame_farm_name.pack_forget()
    frame_location.grid()
    frame_previous_next.grid()
    frame_location_extension.grid()
    count += 1
    # The code below is necessary in the last frame to instruct the user where he can find the results of the analysis.
    # Can only be created when the farm name is known.
    print_finish = 'When you click on the submit button, the questionnaire will close.'
    print_finish_2 = 'The results of the analysis can then be found in: '
    print_farm_name = farm_name.get() + '.xlsx.'
    Label(frame_finish, text=print_finish, justify=LEFT).grid(row=1, column=0, sticky=W, padx=10)
    Label(frame_finish, text=print_finish_2).grid(row=2, column=0, sticky=W, padx=10)
    Label(frame_finish, text=print_farm_name).grid(row=3, column=0, sticky=W, padx=10)
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
            if num < 4:
                list_ans[num].set(str(line.strip('\n')))
            if num >= 4:
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


# cal2 is a function that processes answers on Q2 into a dictionary for use in function 'worksheetoutput'.
# It is done when all answers are submit at the end of the questionnaire.
def cal2():
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
        dic_crops[list_crop_species[i]] = [frac_sur[i], frac_kg[i], kgVeg[i].get()]
    dic_crops = {x: y for x, y in dic_crops.items() if y != [0, 0, 0]}
    return dic_crops

def rid_of_zeros_sur(event, ans, sur):
    try:
        if sur.get() <= 0 and ans.get() == 1:
            sur.set('')
    except:
        sur.set(0)
    if ans.get() == 0:
        sur.set(0)
    return

def rid_of_zeros_kg(event, ans, kg):
    try:
        if kg.get() <= 0 and ans.get() == 1:
            kg.set('')
    except:
        kg.set(0)
    if ans.get() == 0:
        kg.set(0)
    return


def rid_of_zeros(event, answer):
    try:
        if answer.get() <= 0:
            answer.set('')
    except:
        answer.set(0)
    return

def van_owner():
    int_van = IntVar()
    if ans_van_own.get() == "Self":
        int_van = 2
    else:
        int_van = 1
    return int_van

def truck_owner():
    int_truck = IntVar()
    if ans_truck_own.get() == "Self":
        int_truck = 2
    else:
        int_truck = 1
    return int_truck


# ^^ End of functions for the program. Below, the GUI of the program is further developed.
# ------------------------------------------
# Here the start button at the beginning is created
startbutton = Button(frame_start, text='Start', command=start, font=12)
startbutton.pack(fill=X, side=BOTTOM, anchor=CENTER)

# The first page you see when starting the questionnaire
startlabel = Label(frame_start, text='© SFSF, 2019\n', font=12)
copyright_label = Label(frame_start, text='\nVertiCal, a sustainability calculator for vertical farms', font=12)
startlabel.pack(fill=BOTH, side=BOTTOM)
copyright_label.pack(fill=BOTH, side=BOTTOM)
my_image = PhotoImage(file = "avf logo nb.png") # your image
Label(frame_start, image = my_image).pack(side=BOTTOM)

# Enter farm's name
frame_start.pack(anchor=CENTER)
farm_name = StringVar()
Button(frame_farm_name, text='Next', command=next2).pack(fill=BOTH, side=BOTTOM, anchor=CENTER)
Entry(frame_farm_name, textvariable=farm_name).pack(fill=BOTH, side=BOTTOM, anchor=CENTER, pady=5)
Label(frame_farm_name, text='\n\n\n\nEnter the name of your farm:').pack(fill=BOTH, side=BOTTOM)

# Basic frame containing previous and next labels
button2 = Button(frame_previous_next, text=('Previous'), command=pre, padx=10)
button2.grid(row=0, column = 0, padx=10, sticky=W)
shitlabel = Label(frame_previous_next, text='                                   ').grid(row=0, column=1)
button1 = Button(frame_previous_next, text=('  Next  '), command=next1, padx=10)
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

#Here all questions for the questionnaire are defined
question_location = '1. In which country is your farm located?'
question_crop_types = '2. Which crops do you produce? \nWhat area is each crop grown on? \nHow many kilograms of each crop do you sell per year?'
question_buy_renewable = '3. How much renewable and non-renewable electricity (kWh) \ndo you buy per year?'
question_produce_renewable = '4. Do you produce your own renewable energy, \n and how much (kWh) do you produce per year?'
question_sell_electricity = '5. How much electricity (kWh) do you sell per year?'
question_fossil_fuel_use = "6. How much fossil fuel (excluding transportation) do you use per year?"
question_npk_use = "7. How many NPK chemicals (kg) do you buy per year?"
question_substrate_use = '8. Do you use any of the following substrates (kg)\n and how much per year?'
question_water_use = '9. How much water (L) do you buy per year?'
question_pesticide_use = '10. How much pesticides (kg) do you buy per year? '
question_packaging_use = '11. Is the product sold to the customer packaged? '
question_transport = '12. How far does your product travel to the distribution center?\n'\
                     'How are the products divided between the several transportation means?\n'\
                     'Who is the owner of the transportation means?'
question_finish = '13. This is the end of the questionnaire. \nPlease make sure that all questions are answered before you submit.'

# Question 1: Where is your farm located?
# (If you are in this frame, you can't go back and change your name)
# Q1 needs to be specified here because pre and next are not initialized yet
wb = xlrd.open_workbook('Database_full.xlsx')
var = StringVar()
var.set(question_location)
helloLabel = Label(frame_location, textvariable=var, justify=LEFT)
helloLabel.grid(row=0, column=0, padx=10, pady=10, sticky=W)
ans_country = StringVar()
sheet = wb.sheet_by_name('Energy (MJ)')
list_country = []
for i in range (0,sheet.ncols):
    if sheet.cell_value(0,i) == 'Parameter Name':
        for i in range (i,sheet.ncols):
            list_country += [sheet.cell_value(0, i + 1)]
            if sheet.cell_value(0, i+2) == 'World':
                break

country = ttk.Combobox(frame_location_extension, textvariable=ans_country, state='readonly')
country['values'] = list_country
country.current(0)
country.grid(padx=10)

# Here a list of all the possible crops a farmer can choose is read in. This is needed for Q2.
list_crop_species = []
database = wb.sheet_by_name('Crop parameters')
for i in range(1, len(database.col_values(0))):
    if database.col_values(0)[i] == "":
        break
    list_crop_species.append(database.col_values(0)[i])

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

Label(frame_crop_species, text='Crop [-]').grid(row=0, column=0, padx=10, sticky=W)
Label(frame_crop_species, text='Area [m2]').grid(row=0, column=1, padx=5, sticky=W)
Label(frame_crop_species, text='Sold products [kg/year]').grid(row=0, column=2, padx=5, sticky=W)

# In this for loop, the fields for Q2 are created
for i in range(0, len(list_crop_species)):
    Checkbutton(frame_crop_species, text=list_crop_species[i], variable=ansVeg[i]).grid(row=i + 1, column=0, sticky=W, padx=10)
    EntSur = Entry(frame_crop_species, textvariable=surVeg[i])
    EntSur.grid(row=i + 1, column=1, sticky=W, padx=5)
    Entkg = Entry(frame_crop_species, textvariable=kgVeg[i])
    Entkg.grid(row=i + 1, column=2, sticky=W, padx=5)
    EntSur.bind("<FocusIn>", lambda event,y=ansVeg[i], z=surVeg[i]: rid_of_zeros_sur(event,y, z))
    EntSur.bind("<FocusOut>", lambda event,y=ansVeg[i], z=surVeg[i]: rid_of_zeros_sur(event,y, z))
    Entkg.bind("<FocusIn>", lambda event,y=ansVeg[i], z=kgVeg[i]: rid_of_zeros_kg(event,y, z))
    Entkg.bind("<FocusOut>", lambda event,y=ansVeg[i], z=kgVeg[i]: rid_of_zeros_kg(event,y, z))


# Here the fields for question 3 (buying electricity) are created
ans_buy_renew = IntVar()
ans_buy_nonrenew = IntVar()
ans_check_buy_energy = IntVar()
greenlabel = Label(frame_buy_energy, text='Renewable').grid(row=1, column=0, padx=10, sticky=W)
greenentry = Entry(frame_buy_energy, width=10, textvariable=ans_buy_renew)
greenentry.grid(row=1, column=1)
greenentry.bind("<FocusIn>", lambda event,z = ans_buy_renew: rid_of_zeros(event,z))
greenentry.bind("<FocusOut>", lambda event,z = ans_buy_renew: rid_of_zeros(event,z))
greylabel = Label(frame_buy_energy, text='Non-renewable').grid(row=2, column=0, padx=10, sticky=W)
greyentry = Entry(frame_buy_energy, width=10, textvariable=ans_buy_nonrenew)
greyentry.grid(row=2, column=1)
greyentry.bind("<FocusIn>", lambda event,z = ans_buy_nonrenew: rid_of_zeros(event,z))
greyentry.bind("<FocusOut>", lambda event,z = ans_buy_nonrenew: rid_of_zeros(event,z))
Checkbutton(frame_buy_energy, text="I don't know", variable=ans_check_buy_energy).grid(row=3, column=0, sticky=W, padx=10)

# Here the fields for question 4 (creation of renewable energy) are created
ans_prod_solar = IntVar()
ans_prod_biomass = IntVar()
ans_prod_wind = IntVar()
ans_check_create_renewable = IntVar()
solarlabel = Label(frame_create_renewable, text='Solar energy').grid(row=1, column=0, padx=10, sticky=W)
solarentry = Entry(frame_create_renewable, width=10, textvariable=ans_prod_solar)
solarentry.grid(row=1, column=1)
solarentry.bind("<FocusIn>", lambda event,z = ans_prod_solar: rid_of_zeros(event,z))
solarentry.bind("<FocusOut>", lambda event,z = ans_prod_solar: rid_of_zeros(event,z))
biomasslabel = Label(frame_create_renewable, text='Biomass').grid(row=2, column=0, padx=10, sticky=W)
biomassentry = Entry(frame_create_renewable, width=10, textvariable=ans_prod_biomass)
biomassentry.grid(row=2, column=1)
biomassentry.bind("<FocusIn>", lambda event,z = ans_prod_biomass: rid_of_zeros(event,z))
biomassentry.bind("<FocusOut>", lambda event,z = ans_prod_biomass: rid_of_zeros(event,z))
windlabel = Label(frame_create_renewable, text='Windpower').grid(row=3, column=0, padx=10, sticky=W)
windentry = Entry(frame_create_renewable, width=10, textvariable=ans_prod_wind)
windentry.grid(row=3, column=1)
windentry.bind("<FocusIn>", lambda event,z = ans_prod_wind: rid_of_zeros(event,z))
windentry.bind("<FocusOut>", lambda event,z = ans_prod_wind: rid_of_zeros(event,z))
Checkbutton(frame_create_renewable, text="I don't know", variable=ans_check_create_renewable).grid(row=4, column=0, sticky=W, padx=10)

# Here the fields for Q5 (how electricity is used) are created
ans_sel_renew = IntVar()
ans_sel_non_renew = IntVar()
ans_check_sell_energy = IntVar()
SellRenLabel = Label(frame_sell_renewable, text='Selling renewable').grid(row=0, column=0, sticky=W, padx=10)
SellRenEntry = Entry(frame_sell_renewable, width=10, textvariable=ans_sel_renew)
SellRenEntry.grid(row=0, column=1)
SellRenEntry.bind("<FocusIn>", lambda event,z = ans_sel_renew: rid_of_zeros(event,z))
SellRenEntry.bind("<FocusOut>", lambda event,z = ans_sel_renew: rid_of_zeros(event,z))
SellNonrenLabel = Label(frame_sell_renewable, text='Selling non-renewable').grid(row=1, column=0, sticky=W, padx=10)
SellNonrenEntry = Entry(frame_sell_renewable, width=10, textvariable=ans_sel_non_renew)
SellNonrenEntry.grid(row=1, column=1)
SellNonrenEntry.bind("<FocusIn>", lambda event,z = ans_sel_non_renew: rid_of_zeros(event,z))
SellNonrenEntry.bind("<FocusOut>", lambda event,z = ans_sel_non_renew: rid_of_zeros(event,z))
Checkbutton(frame_sell_renewable, text='I don\'t know', variable=ans_check_sell_energy).grid(row=3, column=0, sticky=W, padx=10)

# Here the fields for Q6 (fossil fuel use) are created
ans_petrol_use = IntVar()
ans_diesel_use = IntVar()
ans_natural_gas_use = IntVar()
ans_oil_use = IntVar()
ans_check_fossil_fuel_use = IntVar()
petroll = Label(frame_fuel_use, text='Petrol (L)').grid(row=0, column=0, padx=10, sticky=W)
petroly = Entry(frame_fuel_use, width=5, textvariable=ans_petrol_use)
petroly.grid(row=0, column=1)
petroly.bind("<FocusIn>", lambda event,z = ans_petrol_use: rid_of_zeros(event,z))
petroly.bind("<FocusOut>", lambda event,z = ans_petrol_use: rid_of_zeros(event,z))
diesell = Label(frame_fuel_use, text='Diesel (L)').grid(row=1, column=0, padx=10, sticky=W)
diesely = Entry(frame_fuel_use, width=5, textvariable=ans_diesel_use)
diesely.grid(row=1, column=1)
diesely.bind("<FocusIn>", lambda event,z = ans_diesel_use: rid_of_zeros(event,z))
diesely.bind("<FocusOut>", lambda event,z = ans_diesel_use: rid_of_zeros(event,z))
Ngasl = Label(frame_fuel_use, text='Natural gas (m3)').grid(row=0, column=2, padx=10, sticky=W)
Ngasy = Entry(frame_fuel_use, width=5, textvariable=ans_natural_gas_use)
Ngasy.grid(row=0, column=3)
Ngasy.bind("<FocusIn>", lambda event,z = ans_natural_gas_use: rid_of_zeros(event,z))
Ngasy.bind("<FocusOut>", lambda event,z = ans_natural_gas_use: rid_of_zeros(event,z))
oill = Label(frame_fuel_use, text='Oil (L)').grid(row=1, column=2, padx=10, sticky=W)
oily = Entry(frame_fuel_use, width=5, textvariable=ans_oil_use)
oily.grid(row=1, column=3)
oily.bind("<FocusIn>", lambda event,z = ans_oil_use: rid_of_zeros(event,z))
oily.bind("<FocusOut>", lambda event,z = ans_oil_use: rid_of_zeros(event,z))
Checkbutton(frame_fuel_use, text="I don't know", variable=ans_check_fossil_fuel_use).grid(row=3, column=0, sticky=W, padx=10)


# Here the field for fertilizer use are created (Q7)
ans_ammonium_nitrate_use = IntVar()
ans_calcium_ammonium_nitrate_use = IntVar()
ans_ammonium_sulphate_use = IntVar()
ans_triple_super_phosphate_use = IntVar()
ans_single_super_phosphate_use = IntVar()
ans_ammonia_use = IntVar()
ans_limestone_use = IntVar()
ans_NPK_151515_use = IntVar()
ans_phosphoric_acid_use = IntVar()
ans_mono_ammonium_phosphate_use = IntVar()
ans_check_fertilizer_use = IntVar()
am_label = Label(frame_fertilizer_use, text='Ammoniumnitrate').grid(row=1, column=0, padx=10, sticky=W)
am_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_ammonium_nitrate_use)
am_entry.grid(row=1, column=1)
am_entry.bind("<FocusIn>", lambda event,z = ans_ammonium_nitrate_use: rid_of_zeros(event,z))
am_entry.bind("<FocusOut>", lambda event,z = ans_ammonium_nitrate_use: rid_of_zeros(event,z))
ca_label = Label(frame_fertilizer_use, text='Calciumammoniumnitrate').grid(row=2, column=0, padx=10, sticky=W)
ca_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_calcium_ammonium_nitrate_use)
ca_entry.grid(row=2, column=1)
ca_entry.bind("<FocusIn>", lambda event,z = ans_calcium_ammonium_nitrate_use: rid_of_zeros(event,z))
ca_entry.bind("<FocusOut>", lambda event,z = ans_calcium_ammonium_nitrate_use: rid_of_zeros(event,z))
amsu_label = Label(frame_fertilizer_use, text='Ammoniumsulphate').grid(row=3, column=0, padx=10, sticky=W)
amsu_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_ammonium_sulphate_use)
amsu_entry.grid(row=3, column=1)
amsu_entry.bind("<FocusIn>", lambda event,z = ans_ammonium_sulphate_use: rid_of_zeros(event,z))
amsu_entry.bind("<FocusOut>", lambda event,z = ans_ammonium_sulphate_use: rid_of_zeros(event,z))
tri_label = Label(frame_fertilizer_use, text='Triplesuperphosphate').grid(row=4, column=0, padx=10, sticky=W)
tri_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_triple_super_phosphate_use)
tri_entry.grid(row=4, column=1)
tri_entry.bind("<FocusIn>", lambda event,z = ans_triple_super_phosphate_use: rid_of_zeros(event,z))
tri_entry.bind("<FocusOut>", lambda event,z = ans_triple_super_phosphate_use: rid_of_zeros(event,z))
ssp_label = Label(frame_fertilizer_use, text='Single super phosphate').grid(row=5, column=0, padx=10, sticky=W)
ssp_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_single_super_phosphate_use)
ssp_entry.grid(row=5, column=1)
ssp_entry.bind("<FocusIn>", lambda event,z = ans_single_super_phosphate_use: rid_of_zeros(event,z))
ssp_entry.bind("<FocusOut>", lambda event,z = ans_single_super_phosphate_use: rid_of_zeros(event,z))
ammonia_label = Label(frame_fertilizer_use, text='Ammonia').grid(row=6, column=0, padx=10, sticky=W)
ammonia_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_ammonia_use)
ammonia_entry.grid(row=6, column=1)
ammonia_entry.bind("<FocusIn>", lambda event,z = ans_ammonia_use: rid_of_zeros(event,z))
ammonia_entry.bind("<FocusOut>", lambda event,z = ans_ammonia_use: rid_of_zeros(event,z))
lim_label = Label(frame_fertilizer_use, text='Limestone').grid(row=7, column=0, padx=10, sticky=W)
lim_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_limestone_use)
lim_entry.grid(row=7, column=1)
lim_entry.bind("<FocusIn>", lambda event,z = ans_limestone_use: rid_of_zeros(event,z))
lim_entry.bind("<FocusOut>", lambda event,z = ans_limestone_use: rid_of_zeros(event,z))
npk_label = Label(frame_fertilizer_use, text='NPK 15-15-15').grid(row=8, column=0, padx=10, sticky=W)
npk_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_NPK_151515_use)
npk_entry.grid(row=8, column=1)
npk_entry.bind("<FocusIn>", lambda event,z = ans_NPK_151515_use: rid_of_zeros(event,z))
npk_entry.bind("<FocusOut>", lambda event,z = ans_NPK_151515_use: rid_of_zeros(event,z))
pho_label = Label(frame_fertilizer_use, text='Phosphoric acid').grid(row=9, column=0, padx=10, sticky=W)
pho_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_phosphoric_acid_use)
pho_entry.grid(row=9, column=1)
pho_entry.bind("<FocusIn>", lambda event,z = ans_phosphoric_acid_use: rid_of_zeros(event,z))
pho_entry.bind("<FocusOut>", lambda event,z = ans_phosphoric_acid_use: rid_of_zeros(event,z))
mono_label = Label(frame_fertilizer_use, text='Mono-ammonium phosphate').grid(row=10, column=0, padx=10, sticky=W)
mono_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_mono_ammonium_phosphate_use)
mono_entry.grid(row=10, column=1)
mono_entry.bind("<FocusIn>", lambda event,z = ans_mono_ammonium_phosphate_use: rid_of_zeros(event,z))
mono_entry.bind("<FocusOut>", lambda event,z = ans_mono_ammonium_phosphate_use: rid_of_zeros(event,z))

Checkbutton(frame_fertilizer_use, text="I don't know", variable=ans_check_fertilizer_use).grid(row = 11, column = 0, padx=10, sticky=W)

# Here the fields for substrate use (Q8) are created
ans_rockwool_use = IntVar()
ans_perlite_use = IntVar()
ans_cocofiber_use = IntVar()
ans_hempfiber_use = IntVar()
ans_peat_use = IntVar()
ans_peatmoss_use = IntVar()
ans_check_substrate_use = IntVar()
roc_label = Label(frame_substrate_use, text='Rockwool').grid(row=1, column=0, padx=10, sticky=W)
roc_entry = Entry(frame_substrate_use, width=10, textvariable=ans_rockwool_use)
roc_entry.grid(row=1, column=1)
roc_entry.bind("<FocusIn>", lambda event,z = ans_rockwool_use: rid_of_zeros(event,z))
roc_entry.bind("<FocusOut>", lambda event,z = ans_rockwool_use: rid_of_zeros(event,z))
per_label = Label(frame_substrate_use, text='Perlite').grid(row=2, column=0, padx=10, sticky=W)
per_entry = Entry(frame_substrate_use, width=10, textvariable=ans_perlite_use)
per_entry.grid(row=2, column=1)
per_entry.bind("<FocusIn>", lambda event,z = ans_perlite_use: rid_of_zeros(event,z))
per_entry.bind("<FocusOut>", lambda event,z = ans_perlite_use: rid_of_zeros(event,z))
coc_label = Label(frame_substrate_use, text='Cocofiber').grid(row=1, column=2, padx=10, sticky=W)
coc_entry = Entry(frame_substrate_use, width=10, textvariable=ans_cocofiber_use)
coc_entry.grid(row=1, column=3)
coc_entry.bind("<FocusIn>", lambda event,z = ans_cocofiber_use: rid_of_zeros(event,z))
coc_entry.bind("<FocusOut>", lambda event,z = ans_cocofiber_use: rid_of_zeros(event,z))
hem_label = Label(frame_substrate_use, text='Hemp fiber').grid(row=2, column=2, padx=10, sticky=W)
hem_entry = Entry(frame_substrate_use, width=10, textvariable=ans_hempfiber_use)
hem_entry.grid(row=2, column=3)
hem_entry.bind("<FocusIn>", lambda event,z = ans_hempfiber_use: rid_of_zeros(event,z))
hem_entry.bind("<FocusOut>", lambda event,z = ans_hempfiber_use: rid_of_zeros(event,z))
pea_label = Label(frame_substrate_use, text='Peat').grid(row=3, column=0, padx=10, sticky=W)
pea_entry = Entry(frame_substrate_use, width=10, textvariable=ans_peat_use)
pea_entry.grid(row=3, column=1)
pea_entry.bind("<FocusIn>", lambda event,z = ans_peat_use: rid_of_zeros(event,z))
pea_entry.bind("<FocusOut>", lambda event,z = ans_peat_use: rid_of_zeros(event,z))
peaM_label = Label(frame_substrate_use, text='Peat Moss').grid(row=3, column=2, padx=10, sticky=W)
peaM_entry = Entry(frame_substrate_use, width=10, textvariable=ans_peatmoss_use)
peaM_entry.grid(row=3, column=3)
peaM_entry.bind("<FocusIn>", lambda event,z = ans_peatmoss_use: rid_of_zeros(event,z))
peaM_entry.bind("<FocusOut>", lambda event,z = ans_peatmoss_use: rid_of_zeros(event,z))
Checkbutton(frame_substrate_use, text="I don't know", variable=ans_check_substrate_use).grid(padx=10, row=4, column=0)

# Here the fields for water use (Q9) are created
ans_tap_water_use = IntVar()
ans_check_tap_water_use = IntVar()
water_label = Label(frame_water_use, text='Tap water').grid(row=1, column=0, padx=10, sticky=W)
water_entry = Entry(frame_water_use, width=10, textvariable=ans_tap_water_use)
water_entry.grid(row=1, column=1)
water_entry.bind("<FocusIn>", lambda event,z = ans_tap_water_use: rid_of_zeros(event,z))
water_entry.bind("<FocusOut>", lambda event,z = ans_tap_water_use: rid_of_zeros(event,z))
Checkbutton(frame_water_use, text="I don't know", variable=ans_check_tap_water_use).grid(sticky=W, padx=10, row=2, column=0)

# Here the fields for pesticide use (Q10) are created
ans_atrazine_use = IntVar()
ans_glyphosphate_use = IntVar()
ans_metolachlor_use = IntVar()
ans_herbicide_use = IntVar()
ans_insecticide_use = IntVar()
ans_check_pesticide_use = IntVar()
atr_label = Label(frame_pesticide_use, text='Atrazine').grid(row=1, column=0, padx=10, sticky=W)
atr_entry = Entry(frame_pesticide_use, width=10, textvariable=ans_atrazine_use)
atr_entry.grid(row=1, column=1)
atr_entry.bind("<FocusIn>", lambda event,z = ans_atrazine_use: rid_of_zeros(event,z))
atr_entry.bind("<FocusOut>", lambda event,z = ans_atrazine_use: rid_of_zeros(event,z))
gly_label = Label(frame_pesticide_use, text='Glyphosphate').grid(row=2, column=0, padx=10, sticky=W)
gly_entry = Entry(frame_pesticide_use, width=10, textvariable=ans_glyphosphate_use)
gly_entry.grid(row=2, column=1)
gly_entry.bind("<FocusIn>", lambda event,z = ans_glyphosphate_use: rid_of_zeros(event,z))
gly_entry.bind("<FocusOut>", lambda event,z = ans_glyphosphate_use: rid_of_zeros(event,z))
met_label = Label(frame_pesticide_use, text='Metolachlor').grid(row=3, column=0, padx=10, sticky=W)
met_entry = Entry(frame_pesticide_use, width=10, textvariable=ans_metolachlor_use)
met_entry.grid(row=3, column=1)
met_entry.bind("<FocusIn>", lambda event,z = ans_metolachlor_use: rid_of_zeros(event,z))
met_entry.bind("<FocusOut>", lambda event,z = ans_metolachlor_use: rid_of_zeros(event,z))
her_label = Label(frame_pesticide_use, text='Herbicide').grid(row=4, column=0, padx=10, sticky=W)
her_entry = Entry(frame_pesticide_use, width=10, textvariable=ans_herbicide_use)
her_entry.grid(row=4, column=1)
her_entry.bind("<FocusIn>", lambda event,z = ans_herbicide_use: rid_of_zeros(event,z))
her_entry.bind("<FocusOut>", lambda event,z = ans_herbicide_use: rid_of_zeros(event,z))
ins_label = Label(frame_pesticide_use, text='Insectiside').grid(row=5, column=0, padx=10, sticky=W)
ins_entry = Entry(frame_pesticide_use, width=10, textvariable=ans_insecticide_use)
ins_entry.grid(row=5, column=1)
ins_entry.bind("<FocusIn>", lambda event,z = ans_insecticide_use: rid_of_zeros(event,z))
ins_entry.bind("<FocusOut>", lambda event,z = ans_insecticide_use: rid_of_zeros(event,z))
Checkbutton(frame_pesticide_use, text="I don't know", variable=ans_check_pesticide_use).grid(sticky=W, padx=10, row=6, column=0)

# Here the fields for packaging (Q11) are created
ans_packaging = IntVar()
Radiobutton(frame_packaging_use, text='Yes, it is', variable=ans_packaging, value=1).grid(sticky=W, padx=10)
Radiobutton(frame_packaging_use, text='No, it isn\'t', variable=ans_packaging, value=0).grid(sticky=W, padx=10)

# Here the fields for transportation (Q12)are created
ans_van_use = IntVar()
ans_truck_use = IntVar()
ans_percentage_van_use = IntVar()
ans_percentage_truck_use = IntVar()
ans_check_transport = IntVar()
ans_van_own = StringVar()
ans_truck_own = StringVar()
van_label = Label(frame_transport, text='Van').grid(row=1, column=0, padx=10, sticky=W)
van_entry = Entry(frame_transport, width=10, textvariable=ans_van_use)
van_entry.grid(row=1, column=1)
van_entry.bind("<FocusIn>", lambda event,z = ans_van_use: rid_of_zeros(event,z))
van_entry.bind("<FocusOut>", lambda event,z = ans_van_use: rid_of_zeros(event,z))
tru_label = Label(frame_transport, text='Truck').grid(row=2, column=0, padx=10, sticky=W)
tru_entry = Entry(frame_transport, width=10, textvariable=ans_truck_use)
tru_entry.grid(row=2, column=1)
tru_entry.bind("<FocusIn>", lambda event,z = ans_truck_use: rid_of_zeros(event,z))
tru_entry.bind("<FocusOut>", lambda event,z = ans_truck_use: rid_of_zeros(event,z))
distance_label = Label(frame_transport, text='Distance [km]').grid(row=0, column=1, padx=5, sticky=W)
percent_label = Label(frame_transport, text ="Transported \nproducts [%]").grid(row=0, column=2, padx=5, sticky =W)
van_percent_entry = Entry(frame_transport, width=10, textvariable=ans_percentage_van_use)
van_percent_entry.grid(row=1, column=2)
van_percent_entry.bind("<FocusIn>", lambda event,z = ans_percentage_van_use: rid_of_zeros(event,z))
van_percent_entry.bind("<FocusOut>", lambda event,z = ans_percentage_van_use: rid_of_zeros(event,z))
van_percent_truck = Entry(frame_transport, width=10, textvariable=ans_percentage_truck_use)
van_percent_truck.grid(row=2, column=2)
van_percent_truck.bind("<FocusIn>", lambda event,z = ans_percentage_truck_use: rid_of_zeros(event,z))
van_percent_truck.bind("<FocusOut>", lambda event,z = ans_percentage_truck_use: rid_of_zeros(event,z))
Checkbutton(frame_transport, text="I don't know", variable=ans_check_transport).grid(sticky=W, padx=10, row=3, column=0)
own_label = Label(frame_transport, text ="Owner").grid(row=0, column=3, padx=5, sticky =W)
list_own = ['External', 'Self']
van_own = ttk.Combobox(frame_transport, textvariable=ans_van_own, state='readonly', width=10)
van_own['values'] = list_own
van_own.current(0)
van_own.grid(padx=5, row=1, column=3, pady=5)
truck_own = ttk.Combobox(frame_transport, textvariable=ans_truck_own, state='readonly', width=10)
truck_own['values'] = list_own
truck_own.current(0)
truck_own.grid(padx=5, row=2, column=3)

# Here fields for finishing the questionnaire are created
Button_finish = Button(frame_finish, text=('Submit!'), command=close_program, padx = 10, justify=RIGHT)
Button_finish.grid(row=4, column=0, padx=10)

# At the end, a list containing all the variables is created. It is needed to be able to load previously filled in results
list_ans = [farm_name, ans_country, ans_van_own, ans_truck_own, ansLet, ansEnd, ansSpi, ansBea, ansPar, ansKal, ansBas, ansRuc, ansMic, ansMin,
            surLet, surEnd, surSpi, surBea, surPar, surKal, surBas, surRuc, surMic, surMin, kgLet, kgEnd, kgSpi, kgBea,
            kgPar, kgKal, kgBas, kgRuc, kgMic, kgMin, ans_buy_renew, ans_buy_nonrenew, ans_check_buy_energy,ans_prod_solar,
            ans_prod_biomass, ans_prod_wind, ans_check_create_renewable,ans_sel_renew, ans_sel_non_renew, ans_check_sell_energy,
            ans_petrol_use, ans_diesel_use, ans_natural_gas_use, ans_oil_use, ans_check_fossil_fuel_use, ans_ammonium_nitrate_use,
            ans_calcium_ammonium_nitrate_use, ans_ammonium_sulphate_use, ans_triple_super_phosphate_use, ans_single_super_phosphate_use,
            ans_ammonia_use, ans_limestone_use, ans_NPK_151515_use, ans_phosphoric_acid_use,
            ans_mono_ammonium_phosphate_use, ans_check_fertilizer_use, ans_rockwool_use, ans_perlite_use,
            ans_cocofiber_use, ans_hempfiber_use, ans_peat_use, ans_peatmoss_use, ans_check_substrate_use,
            ans_tap_water_use, ans_check_tap_water_use, ans_atrazine_use, ans_glyphosphate_use,
            ans_metolachlor_use, ans_herbicide_use, ans_insecticide_use, ans_check_pesticide_use, ans_packaging,
            ans_van_use, ans_truck_use, ans_percentage_van_use, ans_percentage_truck_use, ans_check_transport]

# Important statement. If not placed here, program crashes. Assures that all information from above is in the program
root.mainloop()
