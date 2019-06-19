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
frame_previous_next = Frame(height=40, width=400)
frame_farm_name = Frame(height=65, width=400)
frame_location = Frame(height=75, width=400)
frame_location_extension = Frame(height=120, width=400)
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


# Function worksheet_output does all the calculations to perform the LCA and writes the result to an Excel sheet
def worksheet_output(dictionary_name):
    # Opening excel file in order to get parameters
    workbook = xlrd.open_workbook('Database_full.xlsx')
    dic_p = {}
    # Find the parameters of the LCA and place them in dic_p
    for tabs in workbook.sheet_names():
        if tabs != 'Crop parameters':
            sheet = workbook.sheet_by_name(tabs)
            for i in range(1, sheet.nrows):
                if sheet.cell_value(i, 2) != '':
                    for j in range(0, sheet.ncols):
                        if sheet.cell_value(0, j) == ans_country.get():
                            country_name = j
                            if sheet.cell_value(i, j) == '':
                                country_name = 9  # 9 is world
                            dic_p[sheet.cell_value(i, 2)] = sheet.cell_value(i, country_name)
                            if ans_packaging.get() == 0:
                                dic_p['Pac1'] = 0
                                dic_p['Pac2'] = 0

    # If choose 'I don't know’ option, set the value back zero. A message is composed to report in the output which
    # outputs are left out of the analysis.
    non_count = str()
    nr_do_not_know = 0
    if ans_check_buy_energy.get() == 1:
        ans_buy_renew.set(0)
        ans_buy_non_renew.set(0)
        non_count = 'bought electricity, '
        nr_do_not_know += 1
    if ans_check_create_renewable.get() == 1:
        ans_prod_biomass.set(0)
        ans_prod_solar.set(0)
        ans_prod_wind.set(0)
        non_count = (non_count + 'renewable energy production, ')
        nr_do_not_know += 1
    if ans_check_sell_energy.get() == 1:
        ans_sel_renew.set(0)
        ans_sel_non_renew.set(0)
        non_count = (non_count + 'sold electricity, ')
        nr_do_not_know += 1
    if ans_check_fossil_fuel_use.get() == 1:
        ans_oil_use.set(0)
        ans_natural_gas_use.set(0)
        ans_diesel_use.set(0)
        ans_petrol_use.set(0)
        non_count = (non_count + 'fossil fuel use, ')
        nr_do_not_know += 1
    if ans_check_fertilizer_use.get() == 1:
        ans_ammonium_nitrate_use.set(0)
        ans_calcium_ammonium_nitrate_use.set(0)
        ans_ammonium_sulphate_use.set(0)
        ans_triple_super_phosphate_use.set(0)
        ans_single_super_phosphate_use.set(0)
        ans_ammonia_use.set(0)
        ans_limestone_use.set(0)
        ans_NPK_151515_use.set(0)
        ans_phosphoric_acid_use.set(0)
        ans_mono_ammonium_phosphate_use.set(0)
        non_count = (non_count + 'NPK chemicals, ')
        nr_do_not_know += 1
    if ans_check_substrate_use.get() == 1:
        ans_rockwool_use.set(0)
        ans_perlite_use.set(0)
        ans_cocofiber_use.set(0)
        ans_hempfiber_use.set(0)
        ans_peat_use.set(0)
        ans_peatmoss_use.set(0)
        non_count = (non_count + 'substrate, ')
        nr_do_not_know += 1
    if ans_check_tap_water_use.get() == 1:
        ans_tap_water_use.set(0)
        non_count = (non_count + 'water, ')
        nr_do_not_know += 1
    if ans_check_pesticide_use.get() == 1:
        ans_atrazine_use.set(0)
        ans_glyphosphate_use.set(0)
        ans_metolachlor_use.set(0)
        ans_herbicide_use.set(0)
        ans_insecticide_use.set(0)
        non_count = (non_count + 'pesticides, ')
        nr_do_not_know += 1
    if ans_check_transport.get() == 1:
        ans_van_use.set(0)
        ans_truck_use.set(0)
        non_count = (non_count + 'transport')
        nr_do_not_know += 1

    # Create the output: an Excel file
    wb = xlsxwriter.Workbook(farm_name.get() + '.xlsx')
    sheet = workbook.sheet_by_name('Crop parameters')

    # Expand the existing dictionary with data about crops of the farmer with other parameters, such as energy content
    # and growth cycles.
    total_eoc = 0
    total_growth_cycles = 0
    for keys, values in dictionary_name.items():
        for i in range(1, len(list_crop_species) + 1):
            if keys == sheet.cell_value(i, 0):
                # Energy content
                dictionary_name[keys] += [sheet.cell_value(i, 4)]
                total_eoc += sheet.cell_value(i, 4)
                # Growth period
                dictionary_name[keys] += [365/sheet.cell_value(i, 5)]
                total_growth_cycles += (365/sheet.cell_value(i, 5))

    # The values for dictionary element 'Total' contain average values:
    average_eoc = total_eoc / (len(dictionary_name) - 1)
    average_growth_period = total_growth_cycles / (len(dictionary_name)-1)
    dictionary_name[list(dictionary_name.keys())[0]] += [average_eoc]
    dictionary_name[list(dictionary_name.keys())[0]] += [average_growth_period]

    # Calculation to find out the total combination of fraction surface and fraction growth period,
    # needed for the substrate calculations.
    # Take care: this for loop cannot be combined with the for loop behind it since it needs to run through this whole
    # loop, to find the final sum_fraction in order to properly do the calculations on substrate in the next for loop.
    sum_fraction = 0
    for keys, values in dictionary_name.items():
        if keys != 'Total':
            fraction_sur = values[0]
            growth_cycles = values[4]
            # Substrate calculations
            fraction_growth = growth_cycles / total_growth_cycles
            combine_fraction_sur_growth = fraction_growth * fraction_sur
            sum_fraction += combine_fraction_sur_growth

    # From here on, the real calculations on energy and co2 emissions are done.
    for keys, values in dictionary_name.items():
        crop_name = keys
        fraction_sur = values[0]
        # frac_kg = values[1]
        kg_prod = values[2]
        eoc = values[3]
        growth_cycles = values[4]

        # Calculation for total C02 of electricity usage
        eco2 = fraction_sur * (
                (dic_p['C1'] * ans_buy_renew.get()) + (dic_p['C3'] * ans_buy_non_renew.get()) +
                (dic_p['C5'] * ans_prod_solar.get()) + (dic_p['C7'] * ans_prod_wind.get()) +
                (dic_p['C9'] * ans_prod_biomass.get()) - (ans_sel_renew.get() * dic_p['C1']) -
                (ans_sel_non_renew.get() * dic_p['C3']))

        # Calculation for total energy of electricity usage
        e_energy = fraction_sur * (
                (dic_p['C2'] * ans_buy_renew.get()) + (dic_p['C4'] * ans_buy_non_renew.get()) +
                (dic_p['C6'] * ans_prod_solar.get()) + (dic_p['C8'] * ans_prod_wind.get()) +
                (dic_p['C10'] * ans_prod_biomass.get()) - (ans_sel_renew.get() * dic_p['C2']) -
                (ans_sel_non_renew.get() * dic_p['C4']))

        # Calculation for total Co2 of fossil fuels use
        f_co2 = fraction_sur * (
                (dic_p['Fo1'] * ans_petrol_use.get()) + (dic_p['Fo3'] * ans_diesel_use.get()) +
                (dic_p['Fo7'] * ans_natural_gas_use.get()) + (dic_p['Fo9'] * ans_oil_use.get()))

        # Calculation for total energy of fossil fuel use
        f_energy = fraction_sur * (
                (dic_p['Fo2'] * ans_petrol_use.get()) + (dic_p['Fo4'] * ans_diesel_use.get()) +
                (dic_p['Fo8'] * ans_natural_gas_use.get()) + (dic_p['Fo10'] * ans_oil_use.get()))

        # Calculation for total Co2 of fertilizers
        fer_co2 = fraction_sur * (
                (dic_p['Fe1'] * ans_ammonium_nitrate_use.get()) +
                (dic_p['Fe3'] * ans_calcium_ammonium_nitrate_use.get()) +
                (dic_p['Fe5'] * ans_ammonium_sulphate_use.get()) +
                (dic_p['Fe7'] * ans_triple_super_phosphate_use.get()) +
                (dic_p['Fe9'] * ans_single_super_phosphate_use.get()) + (dic_p['Fe11'] * ans_ammonia_use.get()) +
                (dic_p['Fe13'] * ans_limestone_use.get()) + (dic_p['Fe15'] * ans_NPK_151515_use.get()) +
                (dic_p['Fe21'] * ans_phosphoric_acid_use.get()) +
                (dic_p['Fe22'] * ans_mono_ammonium_phosphate_use.get()))

        # Calculation for total energy of fertilizers
        fer_energy = fraction_sur * (
                (dic_p['Fe2'] * ans_ammonium_nitrate_use.get()) +
                (dic_p['Fe4'] * ans_calcium_ammonium_nitrate_use.get()) +
                (dic_p['Fe6'] * ans_ammonium_sulphate_use.get()) +
                (dic_p['Fe8'] * ans_triple_super_phosphate_use.get()) +
                (dic_p['Fe10'] * ans_single_super_phosphate_use.get()) + (dic_p['Fe12'] * ans_ammonia_use.get()) +
                (dic_p['Fe14'] * ans_limestone_use.get()) + (dic_p['Fe16'] * ans_NPK_151515_use.get()) +
                (dic_p['Fe22'] * ans_phosphoric_acid_use.get()) +
                (dic_p['Fe24'] * ans_mono_ammonium_phosphate_use.get()))

        # Calculation in which growth period and fraction of surface are combined into one fraction for substrate
        # calculation.
        if keys != 'Total':
            fraction_growth = growth_cycles / total_growth_cycles
            combine_fraction_sur_growth = fraction_growth * fraction_sur
            fraction_substrate = combine_fraction_sur_growth/sum_fraction
        else:
            # fraction_substrate for total just need to be one, so is set to one here.
            fraction_substrate = 1

        # Calculation for total Co2 of substrates
        s_co2 = fraction_substrate * (
                (dic_p['S1'] * ans_rockwool_use.get()) + (dic_p['S3'] * ans_perlite_use.get()) +
                (dic_p['S5'] * ans_cocofiber_use.get()) + (dic_p['S7'] * ans_hempfiber_use.get()) +
                (dic_p['S9'] * ans_peat_use.get()) + (dic_p['S11'] * ans_peatmoss_use.get()))

        # Calculation for total energy of substrates
        s_energy = fraction_substrate * (
                (dic_p['S2'] * ans_rockwool_use.get()) + (dic_p['S4'] * ans_perlite_use.get()) +
                (dic_p['S6'] * ans_cocofiber_use.get()) + (dic_p['S8'] * ans_hempfiber_use.get()) +
                (dic_p['S10'] * ans_peat_use.get()) + (dic_p['S12'] * ans_peatmoss_use.get()))

        # Calculation for total Co2 of water
        w_co2 = fraction_sur * (
                dic_p['Wa1'] * ans_tap_water_use.get())

        # Calculation for total energy of water
        w_energy = fraction_sur * (
                dic_p['Wa2'] * ans_tap_water_use.get())

        # Calculation for total Co2 of pesticides
        p_co2 = fraction_sur * (
                (dic_p['P1'] * ans_atrazine_use.get()) + (dic_p['P3'] * ans_glyphosphate_use.get()) +
                (dic_p['P5'] * ans_metolachlor_use.get()) + (dic_p['P7'] * ans_herbicide_use.get()) +
                (dic_p['P9'] * ans_insecticide_use.get()))

        # Calculation for total energy of pesticides
        p_energy = fraction_sur * (
                (dic_p['P2'] * ans_atrazine_use.get()) + (dic_p['P4'] * ans_glyphosphate_use.get()) +
                (dic_p['P6'] * ans_metolachlor_use.get()) + +(dic_p['P8'] * ans_herbicide_use.get()) +
                (dic_p['P10'] * ans_insecticide_use.get()))

        # Scaling the percentages of transportation means. If no percentages are filled in, it is assumed that truck
        # and van both account for 50% of the rides.
        if ans_percentage_van_use.get() or ans_percentage_truck_use.get() > 0:
            truck_use_percent = ans_percentage_truck_use.get()/(ans_percentage_truck_use.get() +
                                                                ans_percentage_van_use.get())
            van_use_percent = ans_percentage_van_use.get()/(ans_van_use.get() + ans_percentage_truck_use.get())
        else:
            truck_use_percent = 50
            van_use_percent = 50

        # Calculation for total Co2 of transport
        t_co2 = kg_prod * (
                (dic_p['T3'] * ans_van_use.get() * van_use_percent * van_owner()) +
                (dic_p['T1'] * ans_truck_use.get() * truck_use_percent * truck_owner()))

        # Calculation for total energy of transport
        t_energy = kg_prod * (
                (dic_p['T4'] * ans_van_use.get() * van_use_percent * van_owner()) +
                (dic_p['T2'] * ans_truck_use.get() * truck_use_percent * truck_owner()))

        # Calculation for the total Co2 of packaging
        pac_co2 = kg_prod * dic_p['Pac1']

        # Calculation for the total energy of packaging
        pac_energy = kg_prod * dic_p['Pac2']

        # calculations for the total Co2 and energy
        total_co2 = eco2 + f_co2 + fer_co2 + s_co2 + w_co2 + p_co2 + t_co2 + pac_co2
        total_energy = e_energy + f_energy + fer_energy + s_energy + w_energy + p_energy + t_energy + pac_energy

        # Calculations for the total Co2 and energy per kg product
        total_co2_per_kg_product = total_co2 / kg_prod
        total_energy_per_kg_product = total_energy / kg_prod

        # Calculations for the total Co2 and energy per KJ product
        total_co2_per_kj_product = total_co2_per_kg_product / eoc
        total_energy_per_kj_product = total_energy_per_kg_product / eoc

        # Writing the outputs to the previously created Excel sheet
        ws = wb.add_worksheet(crop_name)
        cell_format_bold = wb.add_format({'bold': True,
                                          'align': 'right',
                                          'fg_color': '#cdcdcd'})
        cell_format_total = wb.add_format({'bold': True,
                                           'align': 'right',
                                           'top': 1,
                                           'fg_color': '#cdcdcd'})
        cell_format_header = wb.add_format({'bold': True,
                                            'font_size': 16,
                                            'align': 'center',
                                            'fg_color': '#cdcdcd'})
        cell_format_questions = wb.add_format({'bold': True,
                                               'font_size': 14})
        cell_format_expl_quest = wb.add_format({'bold': True,
                                                'align': 'right',
                                                'fg_color': '#cdcdcd',
                                                'top': 1})
        cell_format_top = wb.add_format({'bold': True,
                                         'top': 1})
        cell_format_align_r0 = wb.add_format({'align': 'right',
                                             'bg_color': '#e6e6e6'})
        cell_format_align_r1 = wb.add_format({'align': 'right',
                                             'bg_color': '#ffffff'})
        cell_format_ll = wb.add_format({'left': 1})
        cell_format_tl = wb.add_format({'top': 1})
        cell_format_bl = wb.add_format({'bottom': 1})
        cell_format_rll = wb.add_format({'right': 1,
                                         'left': 1})
        cell_format_background1 = wb.add_format({'bg_color': '#e6e6e6'})
        cell_format_background2 = wb.add_format({'bg_color': '#ffffff'})
        cell_formats = [cell_format_background1, cell_format_background2]

        # add SFSF logo to the top of the output
        ws.insert_image('A1', 'sfsf logo png.png', {'x_offset': 20, 'x_scale': 0.05, 'y_scale': 0.05})

        ws.merge_range('B1:C1', 'CO\u2082eq', cell_format_header)
        ws.write(1, 1, 'Total [kg]', cell_format_bold)
        ws.write(1, 2, 'Per kg crop [kg/kg]', cell_format_bold)
        ws.merge_range('D1:E1', 'Energy use', cell_format_header)
        ws.write(1, 3, 'Total [MJ]', cell_format_bold)
        ws.write(1, 4, 'Per kg crop [MJ/kg]', cell_format_bold)
        ws.set_column(2, 1, len('Per kg crop [kg/kg]'))
        ws.set_column(3, 2, len('Per kg crop [kg/kg]'))
        ws.set_column(4, 3, len('Per kg crop [kg/kg]'))
        ws.set_column(5, 4, len('Per kg crop [kg/kg]'))

        labels_output = ['Electricity', 'Fossil fuels', 'Fertilizer', 'Substrates', 'Water', 'Pesticides',
                         'Transport', 'Package']
        co2_emitted = [eco2, f_co2, fer_co2, s_co2, w_co2, p_co2, t_co2, pac_co2]
        co2_emitted_round = [round(elem, 0) for elem in co2_emitted]
        energy_used = [e_energy, f_energy, fer_energy, s_energy, w_energy, p_energy, t_energy, pac_energy]
        energy_used_round = [round(elem, 0) for elem in energy_used]
        co2_crop = []
        energy_crop = []
        sum_co2_per_crop = 0
        sum_energy_per_crop = 0

        for i in range(len(co2_emitted)):
            co2_crop += [co2_emitted[i] / dic_crops[crop_name][2]]
            energy_crop += [energy_used[i] / dic_crops[crop_name][2]]
            co2_crop_round = [round(elem, 3) for elem in co2_crop]
            energy_crop_round = [round(elem, 3) for elem in energy_crop]
            sum_co2_per_crop += co2_emitted[i] / dic_crops[crop_name][2]
            sum_energy_per_crop += energy_used_round[i] / dic_crops[crop_name][2]

        for x in range(len(labels_output)):
            if x % 2 == 0:
                i = 1
            else:
                i = 0
            ws.write(2 + x, 0, labels_output[x], cell_format_bold)
            ws.write(2 + x, 1, co2_emitted_round[x], cell_formats[i])
            ws.write(2 + x, 2, co2_crop_round[x], cell_formats[i])
            ws.write(2 + x, 3, energy_used_round[x], cell_formats[i])
            ws.write(2 + x, 4, energy_crop_round[x], cell_formats[i])

        ws.write(0, 0, '', cell_format_bold)
        ws.write(1, 0, '', cell_format_bold)
        ws.write(2 + len(labels_output), 0, 'Total', cell_format_total)
        ws.write(2 + len(labels_output), 1, round(sum(co2_emitted), 0), cell_format_top)
        ws.write(2 + len(labels_output), 2, round(sum(co2_crop), 3), cell_format_top)
        ws.write(2 + len(labels_output), 3, round(sum(energy_used), 0), cell_format_top)
        ws.write(2 + len(labels_output), 4, round(sum(energy_crop), 3), cell_format_top)
        ws.set_column(0, 0, 12)
        ws.set_row(0, 20)
        ws.set_row(1, 12)

        for i in range(0, 11):
            if i < 5:
                ws.write(11, i, '', cell_format_tl)
            ws.write(i, 5, '', cell_format_ll)

        ws.write(3, 6, '', cell_format_tl)
        ws.write(3, 7, '', cell_format_tl)
        ws.write(1, 8, '', cell_format_ll)
        ws.write(2, 8, '', cell_format_ll)
        ws.write(0, 6, '', cell_format_bl)
        ws.write(0, 7, '', cell_format_bl)
        ws.write(1, 5, '', cell_format_rll)
        ws.write(2, 5, '', cell_format_rll)

        labels_total = ['Total CO\u2082 emitted per KJ product [kg/KJ per year]',
                        'Total energy used per KJ product [KJ/KJ per year]']
        totals_output = [total_co2_per_kj_product, total_energy_per_kj_product]
        totals_output_round = [round(elem, 2) for elem in totals_output]
        for x in range(len(labels_total)):
            ws.write(1 + x, 6, labels_total[x], cell_format_background1)
            ws.write(1 + x, 7, totals_output_round[x], cell_format_background1)
        ws.set_column(6, 6, len('Total energy used per KJ product [KJ/KJ per year]'))

        if ans_check_buy_energy.get() == 1 or ans_check_create_renewable.get() == 1 \
                or ans_check_sell_energy.get() == 1 or ans_check_fossil_fuel_use.get() == 1 \
                or ans_check_fertilizer_use.get() == 1 or ans_check_substrate_use.get() == 1 or \
                ans_check_tap_water_use.get() == 1 or ans_check_pesticide_use.get() == 1 or \
                ans_check_transport.get() == 1:
            if nr_do_not_know <= 4:
                ws.write(12, 1,
                         "Specifications of " + non_count + ' are not taken into account because of lacking data.')
            else:
                warning_format = wb.add_format({'bold': True, 'font_size': 16})
                ws.write(12, 1,
                         "You have used the 'I don't know' button too often. The analysis is missing too much data to"
                         " show significant results. Please try again.", warning_format)

        # Creating bar charts
        chart_col = wb.add_chart({'type': 'column'})
        chart_col.add_series({
            'name': [crop_name, 0, 1],
            'categories': [crop_name, 2, 0, 9, 0],
            'values': [crop_name, 2, 1, 9, 1],
            'fill': {'color': 'black'}
        })
        chart_col.set_title({'name': 'Total CO\u2082eq from different sources',
                             'name_font': {'size': 12}})
        chart_col.set_y_axis({'name': 'CO\u2082eq[kg]',
                              'major_gridlines': {
                                  'visible': False
                              }})
        chart_col.set_x_axis({'name': 'Sources'})
        ws.insert_chart('A15', chart_col, {'x_offset': 20, 'y_offset': 8})

        chart_col = wb.add_chart({'type': 'column'})
        chart_col.add_series({
            'name': [crop_name, 0, 3],
            'categories': [crop_name, 2, 0, 9, 0],
            'values': [crop_name, 2, 3, 9, 3],
            'fill': {'color': 'black'}
        })
        chart_col.set_title({'name': 'Total energy used from different sources',
                             'name_font': {'size': 12}})
        chart_col.set_y_axis({'name': 'Energy [MJ]',
                              'major_gridlines': {
                                  'visible': False
                              }})
        chart_col.set_x_axis({'name': 'Sources'})
        ws.insert_chart('E15', chart_col, {'x_offset': 20, 'y_offset': 8})

        # In the tab total, several more graphs are created than in the other tabs.
        if crop_name == 'Total':
            chart_col = wb.add_chart({'type': 'column'})
            chart_col.add_series({
                'name': [crop_name, 0, 1],
                'categories': [crop_name, 2, 0, 9, 0],
                'values': [crop_name, 2, 1, 9, 1],
                'fill': {'color': 'black'}
            })
            chart_col.set_title({'name': 'Total CO\u2082eq from different sources',
                                 'name_font': {'size': 12}})
            chart_col.set_y_axis({'name': 'CO\u2082eq [kg]',
                                  'major_gridlines': {
                                      'visible': False
                                  }})
            chart_col.set_x_axis({'name': 'Sources', })
            ws.insert_chart('A15', chart_col, {'x_offset': 20, 'y_offset': 8})

            chart_col = wb.add_chart({'type': 'column'})
            chart_col.add_series({
                'name': [crop_name, 0, 4],
                'categories': [crop_name, 1, 0, 10, 0],
                'values': [crop_name, 1, 3, 10, 3],
                'fill': {'color': 'black'}
            })
            chart_col.set_title({'name': 'Total energy used from different sources',
                                 'name_font': {'size': 12}})
            chart_col.set_y_axis({'name': 'Energy [MJ]',
                                  'major_gridlines': {
                                      'visible': False
                                  }})
            chart_col.set_x_axis({'name': 'Sources'})

            chart_co2 = wb.add_chart({'type': 'column'})
            chart_co2.set_title({'name': 'CO\u2082eq per kg crop',
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

            chart_co2.set_y_axis({'name': 'Co\u2082eq [kg/kg]',
                                  'major_gridlines': {
                                      'visible': False
                                  }})
            chart_co2.set_x_axis({'name': 'Sources'})
            chart_energy.set_y_axis({'name': 'Energy [MJ/kg]',
                                     'major_gridlines': {
                                         'visible': False
                                     }})
            chart_energy.set_x_axis({'name': 'Sources'})
            # chart_co2.set_size({'width': 960, 'height': 285})
            ws.insert_chart('A30', chart_co2, {'x_offset': 20, 'y_offset': 8})
            ws.insert_chart('E30', chart_energy, {'x_offset': 20, 'y_offset': 8})

    # Write the raw data from the questionnaire to the Excel sheet
    ws = wb.add_worksheet("Raw data")
    ws.set_column(0, 2, len('Percentage of products [%]'))

    # Write question 1
    ws.write(0, 0, "Question 1", cell_format_questions)
    ws.write(1, 0, 'Country', cell_format_expl_quest)
    ws.write(1, 1, ans_country.get(), cell_format_align_r0)

    # Write question 2
    ws.write(3, 0, "Question 2", cell_format_questions)
    ws.write(4, 0, "Crop type", cell_format_expl_quest)
    ws.write(4, 1, "Area [m\u00b2]", cell_format_expl_quest)
    ws.write(4, 2, "Sold products [kg per year]", cell_format_expl_quest)
    for i in range(0, len(ansVeg)):
        if i % 2 == 0:
            x = 1
        else:
            x = 0
        ws.write(5 + i, 0, list_crop_species[i], cell_format_bold)
        ws.write(5 + i, 1, surVeg[i].get(), cell_formats[x])
        ws.write(5 + i, 2, kgVeg[i].get(), cell_formats[x])

    # Write question 3, 4 and 5
    ws.write(16, 0, "Question 3-5", cell_format_questions)
    ws.write(17, 0, "Electricity type", cell_format_expl_quest)
    ws.write(17, 1, "Amount [kWh per year]", cell_format_expl_quest)
    list_electricity = [ans_buy_renew.get(), ans_buy_non_renew.get(), ans_prod_solar.get(), ans_prod_biomass.get(),
                        ans_prod_wind.get(), ans_sel_renew.get(), ans_sel_non_renew.get()]
    list_electricity_names = ["Bought renewable", "Bought non-renewable", "Produced solar",
                              "Produced biomass", "Produced wind", "Sold renewable",
                              "Sold non-renewable"]
    for i in range(0, len(list_electricity)):
        if i % 2 == 0:
            x = 1
        else:
            x = 0
        ws.write(18 + i, 0, list_electricity_names[i], cell_format_bold)
        ws.write(18 + i, 1, list_electricity[i], cell_formats[x])

    # Question 6
    ws.write(26, 0, "Question 6", cell_format_questions)
    ws.write(27, 0, "Fossil fuel type", cell_format_expl_quest)
    ws.write(27, 1, "Consumption [per year]", cell_format_expl_quest)
    list_fuel = [ans_petrol_use.get(), ans_diesel_use.get(), ans_natural_gas_use.get(), ans_oil_use.get()]
    list_fuel_names = ["Petrol (L)", "Diesel (L)", "Oil (L)", "Natural gas (m\u00b3)"]
    for i in range(0, len(list_fuel)):
        if i % 2 == 0:
            x = 1
        else:
            x = 0
        ws.write(28 + i, 0, list_fuel_names[i], cell_format_bold)
        ws.write(28 + i, 1, list_fuel[i], cell_formats[x])

    # Question 7
    ws.write(33, 0, "Question 7", cell_format_questions)
    ws.write(34, 0, "Fertilizer type", cell_format_expl_quest)
    ws.write(34, 1, "Consumption [kg per year]", cell_format_expl_quest)
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
        if i % 2 == 0:
            x = 1
        else:
            x = 0
        ws.write(35 + i, 0, list_fertilizer_names[i], cell_format_bold)
        ws.write(35 + i, 1, list_fertilizers[i], cell_formats[x])

    # Question 8
    ws.write(46, 0, "Question 8", cell_format_questions)
    ws.write(47, 0, "Substrate type", cell_format_expl_quest)
    ws.write(47, 1, "Consumption [kg per year]", cell_format_expl_quest)
    list_substrates = [ans_rockwool_use.get(), ans_perlite_use.get(), ans_cocofiber_use.get(), ans_hempfiber_use.get(),
                       ans_peat_use.get(), ans_peatmoss_use.get()]
    list_substrates_names = ["Rockwool", "Perlite", "Cocofiber", "Hempfiber", 'Peat', "Peatmoss"]
    for i in range(0, len(list_substrates)):
        if i % 2 == 0:
            x = 1
        else:
            x = 0
        ws.write(48 + i, 0, list_substrates_names[i], cell_format_bold)
        ws.write(48 + i, 1, list_substrates[i], cell_formats[x])

    # Question 9
    ws.write(55, 0, "Question 9", cell_format_questions)
    ws.write(56, 0, "Water consumption:", cell_format_expl_quest)
    ws.write(56, 1, ans_tap_water_use.get(), cell_format_align_r0)
    ws.write(56, 2, "[L per year]", cell_format_align_r0)

    # Question 10
    ws.write(58, 0, "Question 10", cell_format_questions)
    ws.write(59, 0, "Pesticide type", cell_format_expl_quest)
    ws.write(59, 1, "Consumption [kg per year]", cell_format_expl_quest)
    list_pesticides = [ans_atrazine_use.get(), ans_glyphosphate_use.get(),
                       ans_metolachlor_use.get(), ans_herbicide_use.get(), ans_insecticide_use.get()]
    list_pesticides_names = ["Atrazine", "Glyphosphate", "Metolachlore", "Herbicide", "Insecticide"]
    for i in range(0, len(list_pesticides)):
        if i % 2 == 0:
            x = 1
        else:
            x = 0
        ws.write(60 + i, 0, list_pesticides_names[i], cell_format_bold)
        ws.write(60 + i, 1, list_pesticides[i], cell_formats[x])

    # Question 11
    ws.write(66, 0, "Question 11", cell_format_questions)
    ws.write(67, 0, "Packaged [Yes/No]", cell_format_expl_quest)
    if ans_packaging.get == 0:
        ws.write(67, 1, "No", cell_format_align_r0)
    else:
        ws.write(67, 1, "Yes", cell_format_align_r0)

    # Question 12
    ws.write(69, 0, "Question 12", cell_format_questions)
    ws.write(70, 0, "Transportation means", cell_format_expl_quest)
    ws.write(70, 1, "Average distance [km]", cell_format_expl_quest)
    ws.write(71, 0, "Van", cell_format_bold)
    ws.write(71, 1, ans_van_use.get(), cell_formats[0])
    ws.write(72, 0, "Truck", cell_format_bold)
    ws.write(72, 1, ans_truck_use.get(), cell_formats[1])
    ws.write(70, 2, "Percentage of products [%]", cell_format_expl_quest)
    ws.write(71, 2, ans_percentage_van_use.get(), cell_formats[0])
    ws.write(72, 2, ans_percentage_truck_use.get(), cell_formats[1])
    ws.write(70, 3, "Owner", cell_format_expl_quest)
    ws.write(72, 3, ans_truck_own.get(), cell_format_align_r1)

    # Adding border lines
    for i in range(0, 11):
        ws.write(4 + i, 3, '', cell_format_ll)
        ws.write(34 + i, 2, '', cell_format_ll)
        if i < 2:
            ws.write(2, i, '', cell_format_tl)
            ws.write(25, i, '', cell_format_tl)
            ws.write(32, i, '', cell_format_tl)
            ws.write(45, i, '', cell_format_tl)
            ws.write(54, i, '', cell_format_tl)
            ws.write(65, i, '', cell_format_tl)
            ws.write(68, i, '', cell_format_tl)
            ws.write(57, i, '', cell_format_ll)
        if i < 3:
            ws.write(15, i, '', cell_format_tl)
            ws.write(70 + i, 4, '', cell_format_ll)
            ws.write(57, i, '', cell_format_tl)
            ws.write(55, i, '', cell_format_bl)
        if i < 4:
            ws.write(73, i, '', cell_format_tl)
        if i < 5:
            ws.write(27 + i, 2, '', cell_format_ll)
        if i < 6:
            ws.write(59 + i, 2, '', cell_format_ll)
        if i < 7:
            ws.write(47 + i, 2, '', cell_format_ll)
        if i < 8:
            ws.write(17 + i, 2, '', cell_format_ll)
    ws.write(0, 1, '', cell_format_bl)
    ws.write(1, 2, '', cell_format_ll)
    ws.write(66, 1, '', cell_format_bl)
    ws.write(67, 2, '', cell_format_ll)
    ws.write(56, 3, '', cell_format_ll)

    # Close the workbook again
    wb.close()
    return
    # ^^ End of function worksheet output

# function pre() enables to go back to the previous question
# Initialize a new counter:
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
        keep_prev_empty.grid(row=0, column=0, padx=10, pady=0)
        button2.grid_remove()
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
        button1.grid(row=0, column=2, sticky=E, padx=10)
    return


# def next1() enables to go to the next question.
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
        keep_prev_empty.grid_remove()
        button2.grid(row=0, column=0, padx=10, sticky=W)
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
        button1.grid_remove()
    return


# Closes the program
def quit1():
    root.destroy()
    return


def close_program():
    cal2()
    worksheet_output(dic_crops)
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
    fraction_let_sur = fraction_end_sur = fraction_spi_sur = fraction_bea_sur = fraction_par_sur = fraction_kal_sur =\
        fraction_bas_sur = fraction_ruc_sur = fraction_mic_sur = fraction_min_sur = 0
    fraction_sur = [fraction_let_sur, fraction_end_sur, fraction_spi_sur, fraction_bea_sur, fraction_par_sur,
                    fraction_kal_sur, fraction_bas_sur, fraction_ruc_sur, fraction_mic_sur, fraction_min_sur]
    fraction_let_kg = fraction_end_kg = fraction_spi_kg = fraction_bea_kg = fraction_par_kg = fraction_kal_kg =\
        fraction_bas_kg = fraction_ruc_kg = fraction_mic_kg = fraction_min_kg = 0
    fraction_kg = [fraction_let_kg, fraction_end_kg, fraction_spi_kg, fraction_bea_kg, fraction_par_kg, fraction_kal_kg,
                   fraction_bas_kg, fraction_ruc_kg, fraction_mic_kg, fraction_min_kg]
    for i in range(0, len(fraction_sur)):
        fraction_sur[i] = surVeg[i].get() / total_area
        fraction_kg[i] = kgVeg[i].get() / total_kg

    # Creating a dictionary of all parameters: [fraction surface, fraction kg,kg vegetation]
    dic_crops = {}
    dic_crops['Total'] = [1, 1, total_kg]
    for i in range(0, len(fraction_sur)):
        dic_crops[list_crop_species[i]] = [fraction_sur[i], fraction_kg[i], kgVeg[i].get()]
    dic_crops = {x: y for x, y in dic_crops.items() if y != [0, 0, 0]}
    return dic_crops


# This function attempts to remove the zeros in the question on which crops a farmer grows, in the sur entries
def rid_of_zeros_sur(event, ans, sur):
    try:
        if sur.get() <= 0 and ans.get() == 1:
            sur.set('')
    except:
        sur.set(0)
    if ans.get() == 0:
        sur.set(0)
    return

# This function attempts to remove the zeros in the question on which crops a farmer grows, in the seeding entries
def rid_of_zeros_seedlings(event, ans, seedlings):
    try:
        if seedlings.get() <= 0 and ans.get() == 1:
            seedlings.set('')
    except:
        seedlings.set(0)
    if ans.get() == 0:
        seedlings.set(0)
    return


# This function attempts to remove the zeros in the question on which crops a farmer grows, in the kg entries
def rid_of_zeros_kg(event, ans, kg):
    try:
        if kg.get() <= 0 and ans.get() == 1:
            kg.set('')
    except:
        kg.set(0)
    if ans.get() == 0:
        kg.set(0)
    return


# This function attempts to remove the zeros in all entries as soon as they are clicked
def rid_of_zeros(event,answer):
    try:
        if answer.get() <= 0:
            answer.set('')
    except:
        answer.set(0)
    return


# This function checks who owns the van in the transport question, necessary for transport calculations
def van_owner():
    if ans_van_own.get() == "Self":
        int_van = 2
    else:
        int_van = 1
    return int_van


# This function checks who owns the truck in the transport questionn, necessary for transport calculations
def truck_owner():
    if ans_truck_own.get() == "Self":
        int_truck = 2
    else:
        int_truck = 1
    return int_truck


# ^^ End of functions for the program. Below, the GUI of the program is further developed.
# ------------------------------------------
# Here the start button at the beginning is created
start_button = Button(frame_start, text='Start', command=start, font=12)
start_button.pack(fill=X, side=BOTTOM, anchor=CENTER)

# The first page you see when starting the questionnaire
start_label = Label(frame_start, text='© SFSF, 2019\n', font=12)
copyright_label = Label(frame_start, text='\nVertiCal, a sustainability calculator for vertical farms', font=12)
start_label.pack(fill=BOTH, side=BOTTOM)
copyright_label.pack(fill=BOTH, side=BOTTOM)
my_image = PhotoImage(file="avf logo nb.png")
Label(frame_start, image=my_image).pack(side=BOTTOM)

# Enter farm's name
frame_start.pack(anchor=CENTER)
farm_name = StringVar()
Button(frame_farm_name, text='Next', command=next2).pack(fill=BOTH, side=BOTTOM, anchor=CENTER)
Entry(frame_farm_name, textvariable=farm_name).pack(fill=BOTH, side=BOTTOM, anchor=CENTER, pady=5)
Label(frame_farm_name, text='\n\n\n\nEnter the name of your farm:').pack(fill=BOTH, side=BOTTOM)

# Basic frame containing previous and next labels
button2 = Button(frame_previous_next, text='Previous', command=pre, padx=10)
keep_prev_empty = Label(frame_previous_next, text='                       ')
keep_prev_empty.grid(row=0, column=0, padx=10)
space_between_prev_next = Label(frame_previous_next, text='                                   ').grid(row=0, column=1)
button1 = Button(frame_previous_next, text='  Next  ', command=next1, padx=10)
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
file_menu = Menu(menu, tearoff=0)
file_menu.add_command(label='Load', command=file_open)
file_menu.add_command(label='Save', command=file_save)
file_menu.add_command(label='Quit', command=root.quit)
menu.add_cascade(label='File', menu=file_menu)
root.config(menu=menu)

# Here all questions for the questionnaire are defined
question_location = '1. In which country is your farm located? '
question_crop_types = '2. Which crops do you produce? \nWhat area is each crop grown on? \nHow many kg of seedlings ' \
                      'do you buy per year?\nHow many kilograms of each crop do you sell per year?'
question_buy_renewable = '3. How much renewable and non-renewable electricity (kWh) \ndo you buy per year?'
question_produce_renewable = '4. Do you produce your own renewable energy, \n and how much (kWh) do you produce ' \
                             'per year?'
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
question_finish = '13. This is the end of the questionnaire. \nPlease make sure that all questions are answered ' \
                  'before you submit.'

# Question 1: Where is your farm located?
# Q1 needs to be specified here because pre and next are not initialized yet
wb = xlrd.open_workbook('Database_full.xlsx')
var = StringVar()
var.set(question_location)
helloLabel = Label(frame_location, textvariable=var, justify=LEFT)
helloLabel.grid(row=0, column=0, padx=10, pady=10, sticky=W)
ans_country = StringVar()
sheet_q1 = wb.sheet_by_name('Energy (MJ)')

# Load in the list of countries a user can choose from
list_country = []
for i in range(0, sheet_q1.ncols):
    if sheet_q1.cell_value(0, i) == 'Parameter Name':
        for i in range(i, sheet_q1.ncols):
            list_country += [sheet_q1.cell_value(0, i + 1)]
            if sheet_q1.cell_value(0, i+2) == 'World':
                break

country_q1 = ttk.Combobox(frame_location_extension, textvariable=ans_country, state='readonly')
country_q1['values'] = list_country
country_q1.current(0)
country_q1.grid(padx=10)

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

# Initialize variables for buying seeds
seedLet = IntVar()
seedEnd = IntVar()
seedSpi = IntVar()
seedBea = IntVar()
seedPar = IntVar()
seedKal = IntVar()
seedBas = IntVar()
seedRuc = IntVar()
seedMic = IntVar()
seedMin = IntVar()
seedVeg = [seedLet, seedEnd, seedSpi, seedBea, seedPar, seedKal, seedBas, seedRuc, seedMic, seedMin]

Label(frame_crop_species, text='Crop [-]').grid(row=0, column=0, padx=10, sticky=W)
Label(frame_crop_species, text='Area [m\u00b2]').grid(row=0, column=1, padx=5, sticky=W)
Label(frame_crop_species, text='Seedlings\n[kg/year]').grid(row=0, column=2, padx=5, sticky=W)
Label(frame_crop_species, text='Sold products\n[kg/year]').grid(row=0, column=3, padx=5, sticky=W)

# In this for loop, the fields for Q2 are created
for i in range(0, len(list_crop_species)):
    Checkbutton(frame_crop_species, text=list_crop_species[i], variable=ansVeg[i]).grid(row=i + 1, column=0, sticky=W,
                                                                                        padx=10)
    surface_entry = Entry(frame_crop_species, textvariable=surVeg[i], width=12)
    surface_entry.grid(row=i + 1, column=1, sticky=W, padx=5)
    seed_entry = Entry(frame_crop_species, textvariable =seedVeg[i], width=12)
    seed_entry.grid(row=i + 1, column=2, sticky=W, padx=5)
    kg_entry = Entry(frame_crop_species, textvariable=kgVeg[i], width=12)
    kg_entry.grid(row=i + 1, column=3, sticky=W, padx=5)
    surface_entry.bind("<FocusIn>", lambda event, y=ansVeg[i], z=surVeg[i]: rid_of_zeros_sur(event, y, z))
    surface_entry.bind("<FocusOut>", lambda event, y=ansVeg[i], z=surVeg[i]: rid_of_zeros_sur(event, y, z))
    seed_entry.bind("<FocusIn>", lambda event, y=ansVeg[i], z=seedVeg[i]: rid_of_zeros_seedlings(event, y, z))
    seed_entry.bind("<FocusOut>", lambda event, y=ansVeg[i], z=seedVeg[i]: rid_of_zeros_seedlings(event, y, z))
    kg_entry.bind("<FocusIn>", lambda event, y=ansVeg[i], z=kgVeg[i]: rid_of_zeros_kg(event, y, z))
    kg_entry.bind("<FocusOut>", lambda event, y=ansVeg[i], z=kgVeg[i]: rid_of_zeros_kg(event, y, z))

# Here the fields for question 3 (buying electricity) are created
ans_buy_renew = IntVar()
ans_buy_non_renew = IntVar()
ans_check_buy_energy = IntVar()
renewable_label = Label(frame_buy_energy, text='Renewable').grid(row=1, column=0, padx=10, sticky=W)
renewable_entry = Entry(frame_buy_energy, width=10, textvariable=ans_buy_renew)
renewable_entry.grid(row=1, column=1)
renewable_entry.bind("<FocusIn>", lambda event, z=ans_buy_renew: rid_of_zeros(event, z))
renewable_entry.bind("<FocusOut>", lambda event, z=ans_buy_renew: rid_of_zeros(event, z))
non_renewable_label = Label(frame_buy_energy, text='Non-renewable').grid(row=2, column=0, padx=10, sticky=W)
non_renewable_entry = Entry(frame_buy_energy, width=10, textvariable=ans_buy_non_renew)
non_renewable_entry.grid(row=2, column=1)
non_renewable_entry.bind("<FocusIn>", lambda event, z=ans_buy_non_renew: rid_of_zeros(event, z))
non_renewable_entry.bind("<FocusOut>", lambda event, z=ans_buy_non_renew: rid_of_zeros(event, z))
Checkbutton(frame_buy_energy, text="I don't know", variable=ans_check_buy_energy).grid(row=3, column=0, sticky=W,
                                                                                       padx=10)

# Here the fields for question 4 (creation of renewable energy) are created
ans_prod_solar = IntVar()
ans_prod_biomass = IntVar()
ans_prod_wind = IntVar()
ans_check_create_renewable = IntVar()
solar_label = Label(frame_create_renewable, text='Solar energy').grid(row=1, column=0, padx=10, sticky=W)
solar_entry = Entry(frame_create_renewable, width=10, textvariable=ans_prod_solar)
solar_entry.grid(row=1, column=1)
solar_entry.bind("<FocusIn>", lambda event, z=ans_prod_solar: rid_of_zeros(event, z))
solar_entry.bind("<FocusOut>", lambda event, z=ans_prod_solar: rid_of_zeros(event, z))
biomass_label = Label(frame_create_renewable, text='Biomass').grid(row=2, column=0, padx=10, sticky=W)
biomass_entry = Entry(frame_create_renewable, width=10, textvariable=ans_prod_biomass)
biomass_entry.grid(row=2, column=1)
biomass_entry.bind("<FocusIn>", lambda event, z=ans_prod_biomass: rid_of_zeros(event, z))
biomass_entry.bind("<FocusOut>", lambda event, z=ans_prod_biomass: rid_of_zeros(event, z))
wind_label = Label(frame_create_renewable, text='Windpower').grid(row=3, column=0, padx=10, sticky=W)
wind_entry = Entry(frame_create_renewable, width=10, textvariable=ans_prod_wind)
wind_entry.grid(row=3, column=1)
wind_entry.bind("<FocusIn>", lambda event, z=ans_prod_wind: rid_of_zeros(event, z))
wind_entry.bind("<FocusOut>", lambda event, z=ans_prod_wind: rid_of_zeros(event, z))
Checkbutton(frame_create_renewable, text="I don't know", variable=ans_check_create_renewable).grid(row=4, column=0,
                                                                                                   sticky=W, padx=10)

# Here the fields for Q5 (how electricity is used) are created
ans_sel_renew = IntVar()
ans_sel_non_renew = IntVar()
ans_check_sell_energy = IntVar()
sell_renewable_label = Label(frame_sell_renewable, text='Selling renewable').grid(row=0, column=0, sticky=W, padx=10)
sell_renewable_entry = Entry(frame_sell_renewable, width=10, textvariable=ans_sel_renew)
sell_renewable_entry.grid(row=0, column=1)
sell_renewable_entry.bind("<FocusIn>", lambda event, z=ans_sel_renew: rid_of_zeros(event, z))
sell_renewable_entry.bind("<FocusOut>", lambda event, z=ans_sel_renew: rid_of_zeros(event, z))
sell_non_renewable_label = Label(frame_sell_renewable, text='Selling non-renewable').grid(row=1, column=0, sticky=W,
                                                                                          padx=10)
sell_non_renewable_entry = Entry(frame_sell_renewable, width=10, textvariable=ans_sel_non_renew)
sell_non_renewable_entry.grid(row=1, column=1)
sell_non_renewable_entry.bind("<FocusIn>", lambda event, z=ans_sel_non_renew: rid_of_zeros(event, z))
sell_non_renewable_entry.bind("<FocusOut>", lambda event, z=ans_sel_non_renew: rid_of_zeros(event, z))
Checkbutton(frame_sell_renewable, text='I don\'t know', variable=ans_check_sell_energy).grid(row=3, column=0, sticky=W,
                                                                                             padx=10)

# Here the fields for Q6 (fossil fuel use) are created
ans_petrol_use = IntVar()
ans_diesel_use = IntVar()
ans_natural_gas_use = IntVar()
ans_oil_use = IntVar()
ans_check_fossil_fuel_use = IntVar()
petrol_label = Label(frame_fuel_use, text='Petrol (L)').grid(row=0, column=0, padx=10, sticky=W)
petrol_entry = Entry(frame_fuel_use, width=5, textvariable=ans_petrol_use)
petrol_entry.grid(row=0, column=1)
petrol_entry.bind("<FocusIn>", lambda event, z=ans_petrol_use: rid_of_zeros(event, z))
petrol_entry.bind("<FocusOut>", lambda event, z=ans_petrol_use: rid_of_zeros(event, z))
diesel_label = Label(frame_fuel_use, text='Diesel (L)').grid(row=1, column=0, padx=10, sticky=W)
diesel_entry = Entry(frame_fuel_use, width=5, textvariable=ans_diesel_use)
diesel_entry.grid(row=1, column=1)
diesel_entry.bind("<FocusIn>", lambda event, z=ans_diesel_use: rid_of_zeros(event, z))
diesel_entry.bind("<FocusOut>", lambda event, z=ans_diesel_use: rid_of_zeros(event, z))
gas_label = Label(frame_fuel_use, text="Natural gas (m\u00b3)").grid(row=0, column=2, padx=10, sticky=W)
gas_entry = Entry(frame_fuel_use, width=5, textvariable=ans_natural_gas_use)
gas_entry.grid(row=0, column=3)
gas_entry.bind("<FocusIn>", lambda event, z=ans_natural_gas_use: rid_of_zeros(event, z))
gas_entry.bind("<FocusOut>", lambda event, z=ans_natural_gas_use: rid_of_zeros(event, z))
oil_label = Label(frame_fuel_use, text='Oil (L)').grid(row=1, column=2, padx=10, sticky=W)
oil_entry = Entry(frame_fuel_use, width=5, textvariable=ans_oil_use)
oil_entry.grid(row=1, column=3)
oil_entry.bind("<FocusIn>", lambda event, z=ans_oil_use: rid_of_zeros(event, z))
oil_entry.bind("<FocusOut>", lambda event, z=ans_oil_use: rid_of_zeros(event, z))
Checkbutton(frame_fuel_use, text="I don't know", variable=ans_check_fossil_fuel_use).grid(row=3, column=0, sticky=W,
                                                                                          padx=10)
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
am_entry.bind("<FocusIn>", lambda event, z=ans_ammonium_nitrate_use: rid_of_zeros(event, z))
am_entry.bind("<FocusOut>", lambda event, z=ans_ammonium_nitrate_use: rid_of_zeros(event, z))
ca_label = Label(frame_fertilizer_use, text='Calciumammoniumnitrate').grid(row=2, column=0, padx=10, sticky=W)
ca_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_calcium_ammonium_nitrate_use)
ca_entry.grid(row=2, column=1)
ca_entry.bind("<FocusIn>", lambda event, z=ans_calcium_ammonium_nitrate_use: rid_of_zeros(event, z))
ca_entry.bind("<FocusOut>", lambda event, z=ans_calcium_ammonium_nitrate_use: rid_of_zeros(event, z))
am_su_label = Label(frame_fertilizer_use, text='Ammoniumsulphate').grid(row=3, column=0, padx=10, sticky=W)
am_su_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_ammonium_sulphate_use)
am_su_entry.grid(row=3, column=1)
am_su_entry.bind("<FocusIn>", lambda event, z=ans_ammonium_sulphate_use: rid_of_zeros(event, z))
am_su_entry.bind("<FocusOut>", lambda event, z=ans_ammonium_sulphate_use: rid_of_zeros(event, z))
tri_label = Label(frame_fertilizer_use, text='Triplesuperphosphate').grid(row=4, column=0, padx=10, sticky=W)
tri_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_triple_super_phosphate_use)
tri_entry.grid(row=4, column=1)
tri_entry.bind("<FocusIn>", lambda event, z=ans_triple_super_phosphate_use: rid_of_zeros(event, z))
tri_entry.bind("<FocusOut>", lambda event, z=ans_triple_super_phosphate_use: rid_of_zeros(event, z))
ssp_label = Label(frame_fertilizer_use, text='Single super phosphate').grid(row=5, column=0, padx=10, sticky=W)
ssp_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_single_super_phosphate_use)
ssp_entry.grid(row=5, column=1)
ssp_entry.bind("<FocusIn>", lambda event, z=ans_single_super_phosphate_use: rid_of_zeros(event, z))
ssp_entry.bind("<FocusOut>", lambda event, z=ans_single_super_phosphate_use: rid_of_zeros(event, z))
ammonia_label = Label(frame_fertilizer_use, text='Ammonia').grid(row=6, column=0, padx=10, sticky=W)
ammonia_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_ammonia_use)
ammonia_entry.grid(row=6, column=1)
ammonia_entry.bind("<FocusIn>", lambda event, z=ans_ammonia_use: rid_of_zeros(event, z))
ammonia_entry.bind("<FocusOut>", lambda event, z=ans_ammonia_use: rid_of_zeros(event, z))
lim_label = Label(frame_fertilizer_use, text='Limestone').grid(row=7, column=0, padx=10, sticky=W)
lim_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_limestone_use)
lim_entry.grid(row=7, column=1)
lim_entry.bind("<FocusIn>", lambda event, z=ans_limestone_use: rid_of_zeros(event, z))
lim_entry.bind("<FocusOut>", lambda event, z=ans_limestone_use: rid_of_zeros(event, z))
npk_label = Label(frame_fertilizer_use, text='NPK 15-15-15').grid(row=8, column=0, padx=10, sticky=W)
npk_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_NPK_151515_use)
npk_entry.grid(row=8, column=1)
npk_entry.bind("<FocusIn>", lambda event, z=ans_NPK_151515_use: rid_of_zeros(event, z))
npk_entry.bind("<FocusOut>", lambda event, z=ans_NPK_151515_use: rid_of_zeros(event, z))
pho_label = Label(frame_fertilizer_use, text='Phosphoric acid').grid(row=9, column=0, padx=10, sticky=W)
pho_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_phosphoric_acid_use)
pho_entry.grid(row=9, column=1)
pho_entry.bind("<FocusIn>", lambda event, z=ans_phosphoric_acid_use: rid_of_zeros(event, z))
pho_entry.bind("<FocusOut>", lambda event, z=ans_phosphoric_acid_use: rid_of_zeros(event, z))
mono_label = Label(frame_fertilizer_use, text='Mono-ammonium phosphate').grid(row=10, column=0, padx=10, sticky=W)
mono_entry = Entry(frame_fertilizer_use, width=10, textvariable=ans_mono_ammonium_phosphate_use)
mono_entry.grid(row=10, column=1)
mono_entry.bind("<FocusIn>", lambda event, z=ans_mono_ammonium_phosphate_use: rid_of_zeros(event, z))
mono_entry.bind("<FocusOut>", lambda event, z=ans_mono_ammonium_phosphate_use: rid_of_zeros(event, z))
Checkbutton(frame_fertilizer_use, text="I don't know", variable=ans_check_fertilizer_use).grid(row=11, column=0,
                                                                                               padx=10, sticky=W)

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
roc_entry.bind("<FocusIn>", lambda event, z=ans_rockwool_use: rid_of_zeros(event, z))
roc_entry.bind("<FocusOut>", lambda event, z=ans_rockwool_use: rid_of_zeros(event, z))
per_label = Label(frame_substrate_use, text='Perlite').grid(row=2, column=0, padx=10, sticky=W)
per_entry = Entry(frame_substrate_use, width=10, textvariable=ans_perlite_use)
per_entry.grid(row=2, column=1)
per_entry.bind("<FocusIn>", lambda event, z=ans_perlite_use: rid_of_zeros(event, z))
per_entry.bind("<FocusOut>", lambda event, z=ans_perlite_use: rid_of_zeros(event, z))
coc_label = Label(frame_substrate_use, text='Cocofiber').grid(row=1, column=2, padx=10, sticky=W)
coc_entry = Entry(frame_substrate_use, width=10, textvariable=ans_cocofiber_use)
coc_entry.grid(row=1, column=3)
coc_entry.bind("<FocusIn>", lambda event, z=ans_cocofiber_use: rid_of_zeros(event, z))
coc_entry.bind("<FocusOut>", lambda event, z=ans_cocofiber_use: rid_of_zeros(event, z))
hem_label = Label(frame_substrate_use, text='Hemp fiber').grid(row=2, column=2, padx=10, sticky=W)
hem_entry = Entry(frame_substrate_use, width=10, textvariable=ans_hempfiber_use)
hem_entry.grid(row=2, column=3)
hem_entry.bind("<FocusIn>", lambda event, z=ans_hempfiber_use: rid_of_zeros(event, z))
hem_entry.bind("<FocusOut>", lambda event, z=ans_hempfiber_use: rid_of_zeros(event, z))
pea_label = Label(frame_substrate_use, text='Peat').grid(row=3, column=0, padx=10, sticky=W)
pea_entry = Entry(frame_substrate_use, width=10, textvariable=ans_peat_use)
pea_entry.grid(row=3, column=1)
pea_entry.bind("<FocusIn>", lambda event, z=ans_peat_use: rid_of_zeros(event, z))
pea_entry.bind("<FocusOut>", lambda event, z=ans_peat_use: rid_of_zeros(event, z))
peaM_label = Label(frame_substrate_use, text='Peat Moss').grid(row=3, column=2, padx=10, sticky=W)
peaM_entry = Entry(frame_substrate_use, width=10, textvariable=ans_peatmoss_use)
peaM_entry.grid(row=3, column=3)
peaM_entry.bind("<FocusIn>", lambda event, z=ans_peatmoss_use: rid_of_zeros(event, z))
peaM_entry.bind("<FocusOut>", lambda event, z=ans_peatmoss_use: rid_of_zeros(event, z))
Checkbutton(frame_substrate_use, text="I don't know", variable=ans_check_substrate_use).grid(padx=10, row=4, column=0)

# Here the fields for water use (Q9) are created
ans_tap_water_use = IntVar()
ans_check_tap_water_use = IntVar()
water_label = Label(frame_water_use, text='Tap water').grid(row=1, column=0, padx=10, sticky=W)
water_entry = Entry(frame_water_use, width=10, textvariable=ans_tap_water_use)
water_entry.grid(row=1, column=1)
water_entry.bind("<FocusIn>", lambda event, z=ans_tap_water_use: rid_of_zeros(event, z))
water_entry.bind("<FocusOut>", lambda event, z=ans_tap_water_use: rid_of_zeros(event, z))
Checkbutton(frame_water_use, text="I don't know", variable=ans_check_tap_water_use).grid(sticky=W, padx=10, row=2,
                                                                                         column=0)

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
atr_entry.bind("<FocusIn>", lambda event, z=ans_atrazine_use: rid_of_zeros(event, z))
atr_entry.bind("<FocusOut>", lambda event, z=ans_atrazine_use: rid_of_zeros(event, z))
gly_label = Label(frame_pesticide_use, text='Glyphosphate').grid(row=2, column=0, padx=10, sticky=W)
gly_entry = Entry(frame_pesticide_use, width=10, textvariable=ans_glyphosphate_use)
gly_entry.grid(row=2, column=1)
gly_entry.bind("<FocusIn>", lambda event, z=ans_glyphosphate_use: rid_of_zeros(event, z))
gly_entry.bind("<FocusOut>", lambda event, z=ans_glyphosphate_use: rid_of_zeros(event, z))
met_label = Label(frame_pesticide_use, text='Metolachlor').grid(row=3, column=0, padx=10, sticky=W)
met_entry = Entry(frame_pesticide_use, width=10, textvariable=ans_metolachlor_use)
met_entry.grid(row=3, column=1)
met_entry.bind("<FocusIn>", lambda event, z=ans_metolachlor_use: rid_of_zeros(event, z))
met_entry.bind("<FocusOut>", lambda event, z=ans_metolachlor_use: rid_of_zeros(event, z))
her_label = Label(frame_pesticide_use, text='Other herbicides').grid(row=4, column=0, padx=10, sticky=W)
her_entry = Entry(frame_pesticide_use, width=10, textvariable=ans_herbicide_use)
her_entry.grid(row=4, column=1)
her_entry.bind("<FocusIn>", lambda event, z=ans_herbicide_use: rid_of_zeros(event, z))
her_entry.bind("<FocusOut>", lambda event, z=ans_herbicide_use: rid_of_zeros(event, z))
ins_label = Label(frame_pesticide_use, text='Other insectisides').grid(row=5, column=0, padx=10, sticky=W)
ins_entry = Entry(frame_pesticide_use, width=10, textvariable=ans_insecticide_use)
ins_entry.grid(row=5, column=1)
ins_entry.bind("<FocusIn>", lambda event, z=ans_insecticide_use: rid_of_zeros(event, z))
ins_entry.bind("<FocusOut>", lambda event, z=ans_insecticide_use: rid_of_zeros(event, z))
Checkbutton(frame_pesticide_use, text="I don't know", variable=ans_check_pesticide_use).grid(sticky=W, padx=10, row=6,
                                                                                             column=0)

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
van_entry.bind("<FocusIn>", lambda event, z=ans_van_use: rid_of_zeros(event, z))
van_entry.bind("<FocusOut>", lambda event, z=ans_van_use: rid_of_zeros(event, z))
tru_label = Label(frame_transport, text='Truck').grid(row=2, column=0, padx=10, sticky=W)
tru_entry = Entry(frame_transport, width=10, textvariable=ans_truck_use)
tru_entry.grid(row=2, column=1)
tru_entry.bind("<FocusIn>", lambda event, z=ans_truck_use: rid_of_zeros(event, z))
tru_entry.bind("<FocusOut>", lambda event, z=ans_truck_use: rid_of_zeros(event, z))
distance_label = Label(frame_transport, text='Distance [km]').grid(row=0, column=1, padx=5, sticky=W)
percent_label = Label(frame_transport, text="Transported \nproducts [%]").grid(row=0, column=2, padx=5, sticky=W)
van_percent_entry = Entry(frame_transport, width=10, textvariable=ans_percentage_van_use)
van_percent_entry.grid(row=1, column=2)
van_percent_entry.bind("<FocusIn>", lambda event, z=ans_percentage_van_use: rid_of_zeros(event, z))
van_percent_entry.bind("<FocusOut>", lambda event, z=ans_percentage_van_use: rid_of_zeros(event, z))
truck_percent_entry = Entry(frame_transport, width=10, textvariable=ans_percentage_truck_use)
truck_percent_entry.grid(row=2, column=2)
truck_percent_entry.bind("<FocusIn>", lambda event, z=ans_percentage_truck_use: rid_of_zeros(event, z))
truck_percent_entry.bind("<FocusOut>", lambda event, z=ans_percentage_truck_use: rid_of_zeros(event, z))
Checkbutton(frame_transport, text="I don't know", variable=ans_check_transport).grid(sticky=W, padx=10, row=3, column=0)
own_label = Label(frame_transport, text="Owner").grid(row=0, column=3, padx=5, sticky=W)
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
Button_finish = Button(frame_finish, text='Submit!', command=close_program, padx=10, justify=RIGHT)
Button_finish.grid(row=4, column=0, padx=10)

# At the end, a list containing all the variables is created. Needed to be able to load previously filled in results.
list_ans = [farm_name, ans_country, ans_van_own, ans_truck_own, ansLet, ansEnd, ansSpi, ansBea, ansPar, ansKal, ansBas,
            ansRuc, ansMic, ansMin, surLet, surEnd, surSpi, surBea, surPar, surKal, surBas, surRuc, surMic, surMin,
            kgLet, kgEnd, kgSpi, kgBea, kgPar, kgKal, kgBas, kgRuc, kgMic, kgMin, ans_buy_renew, ans_buy_non_renew,
            ans_check_buy_energy, ans_prod_solar, ans_prod_biomass, ans_prod_wind, ans_check_create_renewable,
            ans_sel_renew, ans_sel_non_renew, ans_check_sell_energy, ans_petrol_use, ans_diesel_use,
            ans_natural_gas_use, ans_oil_use, ans_check_fossil_fuel_use, ans_ammonium_nitrate_use,
            ans_calcium_ammonium_nitrate_use, ans_ammonium_sulphate_use, ans_triple_super_phosphate_use,
            ans_single_super_phosphate_use, ans_ammonia_use, ans_limestone_use, ans_NPK_151515_use,
            ans_phosphoric_acid_use, ans_mono_ammonium_phosphate_use, ans_check_fertilizer_use, ans_rockwool_use,
            ans_perlite_use, ans_cocofiber_use, ans_hempfiber_use, ans_peat_use, ans_peatmoss_use,
            ans_check_substrate_use, ans_tap_water_use, ans_check_tap_water_use, ans_atrazine_use, ans_glyphosphate_use,
            ans_metolachlor_use, ans_herbicide_use, ans_insecticide_use, ans_check_pesticide_use, ans_packaging,
            ans_van_use, ans_truck_use, ans_percentage_van_use, ans_percentage_truck_use, ans_check_transport]

# Important statement. If not placed here, program crashes. Assures that all information from above is in the program
root.mainloop()
