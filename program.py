from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import xlrd
import xlsxwriter
import tkinter.filedialog

# Building the Gui including all the frames
root = Tk()
root.title('Life cyle assesment for vertical farms')
root.wm_iconbitmap('sfsf logo.ico')
root.geometry('440x440+500+200')

# Setting heights and widths for every frame.
frame0 = Frame(height=65, width=400)
frame00 = Frame(height=65, width=400)
frame1 = Frame(height=75, width=400)
frame2 = Frame(height=40, width=400)
frame3 = Frame(height=120, width=400)
frame4 = Frame(height=120, width=400)
frame5 = Frame(height=500, width=500)
frame6 = Frame(height=120, width=400)
frame7 = Frame(height=120, width=400)
frame8 = Frame(height=120, width=400)
frame9 = Frame(height=120, width=400)
frame10 = Frame(height=90, width=400)
frame100 = Frame(height=140, width=400)
frame11 = Frame(height=70, width=400)
frame110 = Frame(height=140, width=400)
frame111 = Frame(height=40, width=400)
frame12 = Frame(height=120, width=400)
frame13 = Frame(height=120, width=400)
frame14 = Frame(height=290, width=400)
frame140 = Frame(height=120, width=400)
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
frame22 = Frame(height=90, width=400)
frame220 = Frame(height=120, width=400)
frame23 = Frame(height=50, width=400)
frame230 = Frame(height=50, width=400)
frame24 = Frame(height=75, width=400)
frame_finish = Frame(height=75, width=400)
all_frames = [frame1, frame2, frame3, frame4, frame5, frame6, frame7, frame8, frame9, frame10, frame100, frame11,
              frame12, frame13, frame14, frame140, frame15, frame16, frame160, frame17, frame170, frame18,
              frame180, frame19, frame190, frame20, frame200, frame21, frame210, frame22, frame220, frame23, frame230, frame24]

# Set all the frames to a certain size
for frame in all_frames:
    frame.grid_propagate(0)

# ----------------------------------------
# Below, all functions of the program are created

# Function worksheetoutput does all the calculations to perform the LCA and writes the result to an Excel sheet
def worksheetoutput(dictionary_name):
    # Opening excel file in order to get parameters
    workbook = xlrd.open_workbook('Country database.xlsx')
    sheet = workbook.sheet_by_name(ans1.get())

    # Package.
    # 1 = yes, it is packaged on the farm
    # 2 = No, it is not packaged.
    if ans171.get() == 1:
        Pac1 = float(sheet.cell_value(1, 39))  # Packaging Co2 equivalent/kg
        Pac2 = float(sheet.cell_value(1, 40))  # Packaging energy equivalent/kg
    else:
        Pac1 = 0
        Pac2 = 0

    # C1 - C10 are emission values for electricity, retrieved from an Excel sheet
    # eventually create a loop for this.
    C1 = float(sheet.cell_value(1, 1))  # co2 equivalent green electricty from net
    C2 = float(sheet.cell_value(1, 2))  # energy equivalent green electricty from net
    C3 = float(sheet.cell_value(2, 1))  # co2 equivalent gray electricity from net
    C4 = float(sheet.cell_value(2, 2))  # energy equivalent gray electricity from net
    C5 = float(sheet.cell_value(4, 1))  # Co2 equivalent solar electricity
    C6 = float(sheet.cell_value(4, 2))  # energy equivalent solar electricity
    C7 = float(sheet.cell_value(5, 1))  # Co2 equivalent wind electricity
    C8 = float(sheet.cell_value(5, 2))  # energy equivalent wind electricity
    C9 = float(sheet.cell_value(6, 1))  # Co2 equivalent biomass electricity
    C10 = float(sheet.cell_value(6, 2))  # energy equivalent buimass electricity

    # Fo1- Fo12 are emission values for fossil fuels
    Fo1 = sheet.cell_value(1, 5)  # petrol Co2 equivalent/L
    Fo2 = sheet.cell_value(1, 6)  # petrol Energy equivalent/L
    Fo3 = sheet.cell_value(2, 5)  # diesel Co2 equivalent/L
    Fo4 = sheet.cell_value(2, 6)  # diesel Energy equivalent/L
    Fo7 = sheet.cell_value(4, 5)  # natural gas Co2 equivalent/L
    Fo8 = sheet.cell_value(4, 6)  # natural gas Energy equivalent/L
    Fo9 = sheet.cell_value(5, 5)  # oil Co2 equivalent/L
    Fo10 = sheet.cell_value(5, 6)  # oil Energy equivalent/L
    Fo11 = sheet.cell_value(6, 6)  # hard coal Co2 equivalent/L
    F012 = sheet.cell_value(6, 6)  # hard coal Energy equivalent/L

    # Fe1-Fe24 are emission values for fertilizers
    Fe1 = sheet.cell_value(1, 9)  # amonium nitrate Co2 equivalent/L
    Fe2 = sheet.cell_value(1, 10)  # amonium nitrate energy equivalent/L
    Fe3 = sheet.cell_value(2, 9)  # Calcium Ammonium Nitrate Co2 equivalent/L
    Fe4 = sheet.cell_value(2, 10)  # Calcium Ammonium Nitrate energy equivalent/L
    Fe5 = sheet.cell_value(3, 9)  # Ammonium Sulphate Co2 equivalent/L
    Fe6 = sheet.cell_value(3, 10)  # Ammonium Sulphate energy equivalent/L
    Fe7 = sheet.cell_value(4, 9)  # Triple Superphosphate Co2 equivalent/L
    Fe8 = sheet.cell_value(4, 10)  # Triple Superphosphate energy equivalent/L
    Fe9 = sheet.cell_value(5, 9)  # Single super phosphate Co2 equivalent/L
    Fe10 = sheet.cell_value(5, 10)  # Single super phosphate energy equivalent/L
    Fe11 = sheet.cell_value(6, 9)  # Ammonia Co2 equivalent/L
    Fe12 = sheet.cell_value(6, 10)  # Ammonia energy equivalent/L
    Fe13 = sheet.cell_value(7, 9)  # limestone Co2 equivalent/L
    Fe14 = sheet.cell_value(7, 10)  # limestone energy equivalent/L
    Fe15 = sheet.cell_value(8, 9)  # NPK 15-15-15 Co2 equivalent/L
    Fe16 = sheet.cell_value(8, 10)  # NPK 15-15-15 energy equivalent/L
    Fe17 = sheet.cell_value(9, 9)  # Urea Co2 equivalent/L
    Fe18 = sheet.cell_value(9, 10)  # Urea energy equivalent/L
    Fe19 = sheet.cell_value(10, 9)  # cow manure Co2 equivalent/L
    Fe20 = sheet.cell_value(10, 10)  # cow manure energy equivalent/L
    Fe21 = sheet.cell_value(11, 9)  # phosphoric acid Co2 equivalent/L
    Fe22 = sheet.cell_value(11, 10)  # phosphoric acid energy equivalent/L
    Fe23 = sheet.cell_value(12, 9)  # Mono-ammonium phosphate Co2 equivalent/L
    Fe24 = sheet.cell_value(12, 10)  # Mono-ammonium phosphate energy equivalent/L

    # S1-S12 are emission values for substrates
    S1 = sheet.cell_value(1, 13)  # Rockwool Co2 equivalent/
    S2 = sheet.cell_value(1, 14)  # Rockwool energy equivalent/L
    S3 = sheet.cell_value(2, 13)  # Perlite Co2 equivalent/L
    S4 = sheet.cell_value(2, 14)  # Perlite energy equivalent/L
    S5 = sheet.cell_value(3, 13)  # Coco Fiber (coir pith) Co2 equivalent/L
    S6 = sheet.cell_value(3, 14)  # Coco Fiber (coir pith) energy equivalent/L
    S7 = sheet.cell_value(4, 13)  # Hemp fiber Co2 equivalent/L
    S8 = sheet.cell_value(4, 14)  # Hemp fiber energy equivalent/L
    S9 = sheet.cell_value(5, 13)  # Peat Co2 equivalent/L
    S10 = sheet.cell_value(5, 14)  # Peat energy equivalent/L
    S11 = sheet.cell_value(6, 13)  # Peat moss Co2 equivalent/L
    S12 = sheet.cell_value(6, 14)  # Peat moss energy equivalent/L

    # W1 and W2 are emission values of water
    W1 = sheet.cell_value(1, 17)  # Tapwater Co2 equivalent/L
    W2 = sheet.cell_value(1, 18)  # Tapwater energy equivalent/L

    # P1-P10 are emission values of pesticides
    P1 = sheet.cell_value(1, 21)  # Atrazine water Co2 equivalent/L
    P2 = sheet.cell_value(1, 22)  # Atrazine energy equivalent/L
    P3 = sheet.cell_value(2, 21)  # Glyphosphate water Co2 equivalent/L
    P4 = sheet.cell_value(2, 22)  # Glyphosphate energy equivalent/L
    P5 = sheet.cell_value(3, 21)  # Metolachlor Co2 equivalent/L
    P6 = sheet.cell_value(3, 22)  # Metolachlor energy equivalent/L
    P7 = sheet.cell_value(4, 21)  # Herbicide Co2 equivalent/L
    P8 = sheet.cell_value(4, 22)  # Herbicide energy equivalent/L
    P9 = sheet.cell_value(5, 21)  # Insectiside Co2 equivalent/L
    P10 = sheet.cell_value(5, 22)  # Insectiside energy equivalent/L

    # W1 - W6 are emission values for waste
    W1 = sheet.cell_value(1, 25)  # green waste Co2 equivalent/kg
    W2 = sheet.cell_value(1, 26)  # green waste energy equivalent/L
    W3 = sheet.cell_value(2, 25)  # other waste Co2 equivalent/kg
    W4 = sheet.cell_value(2, 26)  # other waste energy equivalent/L
    W5 = sheet.cell_value(3, 25)  # paper waste Co2 equivalent/kg
    W6 = sheet.cell_value(3, 26)  # paper waste energy equivalent/L

    # Here the emission for transport by a plane is calculated
    Tvp1 = sheet.cell_value(1, 29)  # regression factor plane
    Tvp2 = sheet.cell_value(2, 29)  # regression factor plane
    Tvp3 = sheet.cell_value(4, 29)  # extra distance travelled plane
    Tvp4 = sheet.cell_value(5, 29)  # extra langdings made plane
    Tvp5 = sheet.cell_value(6, 29)  # landing take of kerosene usage plane
    Tvp6 = sheet.cell_value(7, 29)  # co2 emissions 1 L of kerosene (Co2-eq/kg) plane
    Tvp7 = sheet.cell_value(8, 29)  # radiative forcing factor plane
    Tvp8 = sheet.cell_value(9, 29)  # possible cargo plane
    Tvp9 = sheet.cell_value(10, 29)  # average percent full plane
    T1 = ((Tvp1 * (ans201.get() * Tvp3) ** 2 + Tvp2 * (ans201.get() * Tvp3) * Tvp6 * Tvp7) + Tvp4 * Tvp5 * Tvp6) / (
            ans201.get() + 0.001) / (Tvp8 * Tvp9)  # plane Co2 equivalent/(Ton*km)
    T2 = sheet.cell_value(1, 32)  # Plane energy equivalent/(ton*Km)

    # Here the emission for transport by a truck is calculated
    Tvt1 = sheet.cell_value(12, 29)  # regression factor Truck
    Tvt2 = sheet.cell_value(13, 29)  # regression factor Truck
    Tvt3 = sheet.cell_value(14, 29)  # regression factor Truck
    Tvt4 = sheet.cell_value(16, 29)  # extra distance travelled Truck
    Tvt5 = sheet.cell_value(17, 29)  # regression factor Truck
    Tvt6 = sheet.cell_value(18, 29)  # Co2 emission 1 L of diesel (Co2-eq/kg) Truck
    Tvt7 = sheet.cell_value(19, 29)  # max cargo (ton) Truck
    Tvt8 = sheet.cell_value(20, 29)  # average percent full Truck
    T3 = ((Tvt1 * Tvt7 + Tvt2 * Tvt3 * ans202.get() * Tvt4) * Tvt5 * Tvt6) / (ans202.get() + 0.001) / (
            Tvt8 * Tvt7)  # truck C02 equivalent/(Ton*km)
    T4 = sheet.cell_value(12, 32)  # Truck energy equivalent/(ton*Km)

    # Here the emission for transport by a ship is calculated
    Tp1 = sheet.cell_value(22, 29)  # regression factor plane
    Tp2 = sheet.cell_value(23, 29)  # regression factor plane
    Tp3 = sheet.cell_value(25, 29)  # extra distance travelled plane
    Tp4 = sheet.cell_value(26, 29)  # regression factor plane
    Tp5 = sheet.cell_value(27, 29)  # litre of oil used in harbours
    Tp6 = sheet.cell_value(28, 29)  # extra stops see harbours
    Tp7 = sheet.cell_value(29, 29)  # Co2 emission 1 L of Oil (Co2-eq/kg)
    Tp8 = sheet.cell_value(30, 29)  # Cargo (ton)
    Tp9 = sheet.cell_value(31, 29)  # average percentage full
    T5 = (((Tp1 * Tp8 + Tp2 * Tp4 * ans203.get() * Tp3 + (Tp5 * (Tp6 + 1)) * Tp7) / (ans203.get() + 0.001))) / (
            Tp8 * Tp9)  # ship C02 equivalent/(Ton*km)
    T6 = sheet.cell_value(22, 32)  # Ship energy equivalent/(ton*Km)

    # Add Energy content to an earlier build dictionary
    workbook = xlrd.open_workbook('Crops energy content.xlsx')
    sheet = workbook.sheet_by_name('Basic database')
    count = 1
    for keys, values in dictionary_name.items():
        if keys == sheet.cell_value(count, 2):
            dictionary_name[keys] += [sheet.cell_value(count, 3)]
        count += 1

    print(dictionary_name)

    non_count = str()
    # If choose 'I don't know’ option, set the value back to zero
    if ans890.get() == 1:
        ans81.set(0);
        ans82.set(0);
        ans83.set(0);
        ans84.set(0);
        ans85.set(0);
        ans86.set(0);
        ans87.set(0);
        ans88.set(0);
        ans89.set(0)
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
        ans129.set(0);
        ans130.set(0);
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
    if ans184.get() == 1:
        ans181.set(0);
        ans182.set(0);
        ans183.set(0)
        non_count = (non_count + 'Waste,')
    if ans204.get() == 1:
        ans201.set(0);
        ans202.set(0);
        ans203.set(0)
        non_count = (non_count + 'NPK chemicals,')
    if ans212.get() == 1:
        ans211.set(0)
        non_count = (non_count + 'Waste during transportation')

    # Create the output: an Excel file
    wb = xlsxwriter.Workbook(farm_name.get() + '.xlsx')
    Total_Eoc = 0
    for keys, values in dictionary_name.items():
        for i in range(1, len(lis)):
            if keys == sheet.cell_value(i, 2):
                dictionary_name[keys] += [sheet.cell_value(i, 3)]
                Total_Eoc += sheet.cell_value(i, 3)
    Average_Eoc = Total_Eoc / (len(dictionary_name) - 1)
    dictionary_name[list(dictionary_name.keys())[0]] += [Average_Eoc]

    for keys, values in dictionary_name.items():
        cropname = keys
        kg_prod = values[2]
        frac_surf = values[0]
        frac_kg = values[1]
        Eoc = values[3]

        # Calculation for total C02 of electricity usage
        Eco2 = frac_surf * ((C1 * ans61.get()) + (C3 * ans62.get()) + (C5 * ans71.get()) + (C7 * ans73.get()) + (
                C9 * ans72.get()) - (ans87.get() * C1) - (ans88.get() * C3))

        # Calculation for total energy of electricity usage
        Eenergy = frac_surf * ((C2 * ans61.get()) + (C4 * ans62.get()) + (C6 * ans71.get()) + (C8 * ans73.get()) + (
                C10 * ans72.get()) - (ans87.get() * C2) - (ans88.get() * C4))

        # Calculation for total Co2 of fossil fuels use
        Fco2 = frac_surf * ((Fo1 * ans91.get()) + (Fo3 * ans92.get()) + (Fo7 * ans94.get()) + (Fo9 * ans95.get()))

        # Calculation for total energy of fossil fuel use
        Fenergy = frac_surf * (
                (Fo2 * ans91.get()) + (Fo4 * ans92.get()) + (Fo8 * ans94.get()) + (Fo10 * ans95.get()))

        # Calculation for total Co2 of fertilizers
        FERco2 = frac_surf * ((
                (Fe1 * ans121.get()) + (Fe3 * ans122.get()) + (Fe5 * ans123.get()) + (Fe7 * ans124.get()) + (
                Fe9 * ans125.get()) + (Fe11 * ans126.get()) + (Fe13 * ans127.get()) + (
                        Fe15 * ans128.get()) + (Fe17 * ans129.get()) + (Fe19 * ans130.get()) + (
                        Fe21 * ans131.get()) + (Fe22 * ans132.get())))

        # Calculation for total energy of fertilizers
        FERenergy = frac_surf * ((
                (Fe2 * ans121.get()) + (Fe4 * ans122.get()) + (Fe6 * ans123.get()) + (Fe8 * ans124.get()) + (
                Fe10 * ans125.get()) + (Fe12 * ans126.get()) + (Fe14 * ans127.get()) + (
                        Fe16 * ans128.get()) + (Fe18 * ans129.get()) + (Fe20 * ans130.get()) + (
                        Fe22 * ans131.get()) + (Fe24 * ans132.get())))

        # Calculation for total Co2 of substrates
        Sco2 = frac_surf * (
                (S1 * ans141.get()) + (S3 * ans142.get()) + (S5 * ans143.get()) + (S7 * ans144.get()) + (
                S9 * ans145.get()) + (S11 * ans146.get()))

        # Calculation for total energy of substrates
        Senergy = frac_surf * (
                (S2 * ans141.get()) + (S4 * ans142.get()) + (S6 * ans143.get()) + (S8 * ans144.get()) + (
                S10 * ans145.get()) + (S12 * ans146.get()))

        # Calculation for total Co2 of water
        Wco2 = frac_surf * (W1 * ans151.get())

        # Calculation for total energy of water
        Wenergy = frac_surf * (W2 * ans151.get())

        # Calculation for total Co2 of pesticides
        Pco2 = frac_surf * (
                (P1 * ans161.get()) + (P3 * ans162.get()) + (P5 * ans163.get()) + (P7 * ans164.get()) + (
                P9 * ans165.get()))

        # Calculation for total energy of pesticides
        Penergy = frac_surf * (
                (P2 * ans161.get()) + (P4 * ans162.get()) + +(P6 * ans163.get()) + +(P8 * ans164.get()) + (
                P10 * ans165.get()))

        # Calculation for total Co2 of transport
        Tco2 = frac_kg * ((T1 * ans201.get()) + (T3 * ans202.get()) + (T5 * ans203.get()))

        # Calculation for total energy of transport
        Tenergy = frac_kg * ((T2 * ans201.get()) + (T4 * ans202.get()) + (T6 * ans203.get()))

        # Calculation for the total Co2 of packaging
        Pacco2 = (kg_prod * (100 - ans211.get()) / 100) * Pac1

        # Calculation for the total energy of packaging
        Pacenergy = (kg_prod * (100 - ans211.get()) / 100) * Pac2

        # calculations for the total Co2 and energy
        Totalco2 = Eco2 + Fco2 + FERco2 + Sco2 + Wco2 + Pco2 + Tco2 + Pacco2
        Totalenergy = Eenergy + Fenergy + FERenergy + Senergy + Wenergy + Penergy + Tenergy + Pacenergy

        # calculations for the total Co2 and energy per kg product #ans5 moet kg worden
        Totalco2_per_kg_product = Totalco2 / (kg_prod * (100 - ans211.get()) / 100)
        Totalenergy_per_kg_product = Totalenergy / (kg_prod * (100 - ans211.get()) / 100)

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

        labels_diff_aspects = ['Heating', 'Cooling', 'Electricity', 'Tillage', 'Sowing', 'Weeding', 'Harvest',
                               'Fertilizer', 'Irrigation', 'Pesticide', 'Other']
        for x in range(len(labels_diff_aspects)):
            ws.write(43 + x, 1, labels_diff_aspects[x])
            ws.write(43 + x, 2, list_ans[22 + x].get()) * frac_surf

        ws.write(43, 9, 'Heating')
        ws.write(44, 9, 'Cooling')
        ws.write(45, 9, 'Ventilation')
        ws.write(46, 9, 'Lighting')
        ws.write(47, 9, 'Machinery')
        ws.write(48, 9, 'Storage')
        ws.write(49, 9, 'Selling')
        ws.write(50, 9, 'Other')

        ws.write(43, 10, ans81.get()) * frac_surf
        ws.write(44, 10, ans82.get()) * frac_surf
        ws.write(45, 10, ans83.get()) * frac_surf
        ws.write(46, 10, ans84.get()) * frac_surf
        ws.write(47, 10, ans85.get()) * frac_surf
        ws.write(48, 10, ans86.get()) * frac_surf
        ws.write(49, 10, ans87.get()) * frac_surf
        ws.write(50, 10, ans88.get()) * frac_surf

        if ans890.get() == 1 or ans131.get() == 1 or ans147.get() == 1 or ans152.get() == 1 or ans166.get() == 1 or ans184.get() == 1 or ans204.get() == 1:
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

        chart_col = wb.add_chart({'type': 'pie'})
        chart_col.add_series({
            'name': 'Fossil fuels used',
            'categories': [cropname, 43, 1, 53, 1],
            'values': [cropname, 43, 2, 53, 2],
            'points': [{'fill': {'color': 'blue'}},
                       {'fill': {'color': 'yellow'}},
                       {'fill': {'color': 'red'}},
                       {'fill': {'color': 'gray'}},
                       {'fill': {'color': 'black'}},
                       {'fill': {'color': 'purple'}},
                       {'fill': {'color': 'pink'}},
                       {'fill': {'color': 'cyan'}},
                       {'fill': {'color': 'magenta'}},
                       {'fill': {'color': 'brown'}},
                       ],
        })
        chart_col.set_title({'name': 'Fossil fuels used for different aspects'})
        chart_col.set_style(5)
        ws.insert_chart('A43', chart_col, {'x_offset': 25, 'y_offset': 10})

        chart_col = wb.add_chart({'type': 'pie'})
        chart_col.add_series({
            'name': 'Electricity used',
            'categories': [cropname, 43, 9, 50, 9],
            'values': [cropname, 43, 10, 50, 10],
            'points': [{'fill': {'color': 'blue'}},
                       {'fill': {'color': 'yellow'}},
                       {'fill': {'color': 'red'}},
                       {'fill': {'color': 'gray'}},
                       {'fill': {'color': 'black'}},
                       {'fill': {'color': 'purple'}},
                       {'fill': {'color': 'pink'}},
                       {'fill': {'color': 'cyan'}},

                       ],
        })
        chart_col.set_title({'name': 'Electricity used for different aspects'})
        chart_col.set_style(4)
        ws.insert_chart('I43', chart_col, {'x_offset': 25, 'y_offset': 10})

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
        var.set('2. What crop do you produce?')
        frame8.grid_forget()
        frame5.grid(sticky=W)
    if count == 3:
        var.set(
            '3. How much renewable and non-renewable electricity \ndo you buy per year? \n (own production not included)')
        frame9.grid_forget()
        frame8.grid(sticky=W)
    if count == 4:
        var.set('4. Do you produce your own renewable energy and how much \ndo you produce?')
        frame10.grid_forget()
        frame100.grid_forget()
        frame9.grid(sticky=W)
    if count == 5:
        var.set("5. Can you specify what and how much the electricity is spend on? \nif you can't fill in zeros")
        frame11.grid_forget()
        frame110.grid_forget()
        frame111.grid_forget()
        frame10.grid(sticky=W)
        frame100.grid(sticky=W)
    if count == 6:
        var.set(
            "6. Do you use any fossil fuels(excluding transportation), \nand how much do you use(if you don't know fill in zero)")
        frame14.grid_forget()
        frame140.grid_forget()
        frame11.grid(sticky=W)
        frame111.grid(sticky=W)
        frame110.grid(sticky=W)
    if count == 7:
        var.set('7. How much kilograms do you use of the following \nNPK chemicals per year?')
        frame16.grid_forget()
        frame160.grid_forget()
        frame14.grid(sticky=W)
        frame140.grid(sticky=W)
    if count == 8:
        var.set('8. Do you use substrate and how much per year? (kg)')
        frame17.grid_forget()
        frame170.grid_forget()
        frame16.grid(sticky=W)
        frame160.grid(sticky=W)
    if count == 9:
        var.set('9. How much water do you use?')
        frame18.grid_forget()
        frame180.grid_forget()
        frame17.grid(sticky=W)
        frame170.grid(sticky=W)
    if count == 10:
        var.set('10. How much pesticides do you use? ')
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
        var.set('12. How much waste do you produce? ')
        frame22.grid_forget()
        frame220.grid_forget()
        frame20.grid(sticky=W)
        frame200.grid(sticky=W)
    if count == 13:
        var.set('13. How far does your product travel to the distribution center \non average? ')
        frame23.grid_forget()
        frame230.grid_forget()
        frame22.grid(sticky=W)
        frame220.grid(sticky=W)
    if count == 14:
        var.set('14. How much of the product does not survive the transport \nstage to the store? ')
        frame_finish.grid_forget()
        frame23.grid(sticky=W)
        frame230.grid(sticky=W)
    return


# def next1() enables to go to the next question.
# i.e. forgetting the current frames and introducing new frames
v = IntVar()


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
        var.set('2. What crop do you produce?')
        frame3.grid_forget()
        frame5.grid(sticky=W)
    if count == 3:
        var.set(
            '3. How much renewable and non-renewable electricity \ndo you buy per year? \n (own production not included)')
        frame5.grid_forget()
        frame8.grid(sticky=W)
    if count == 4:
        var.set('4. Do you produce your own renewable energy and how much \ndo you produce?')
        frame8.grid_forget()
        frame9.grid(sticky=W)
    if count == 5:
        var.set("5. Can you specify what and how much the electricity is spend on? \nif you can't fill in zeros")
        frame9.grid_forget()
        frame10.grid(sticky=W)
        frame100.grid(sticky=W)
    if count == 6:
        var.set(
            "6. Do you use any fossil fuels(excluding transportation), \nand how much do you use (if you don't know fill in zero)")
        frame10.grid_forget()
        frame100.grid_forget()
        frame11.grid(sticky=W)
        frame111.grid(sticky=W)
        frame110.grid(sticky=W)
    if count == 7:
        var.set('7. How much kilograms do you use of the following \nNPK chemicals per year?')
        other = 100 - ans916.get() - ans915.get() - ans914.get() - ans913.get() - ans912.get() - ans911.get() - ans910.get() - ans99.get() - ans98.get() - ans97.get()
        ans917.set(other)
        frame11.grid_forget()
        frame111.grid_forget()
        frame110.grid_forget()
        frame14.grid(sticky=W)
        frame140.grid(sticky=W)
    if count == 8:
        var.set('8. Do you use substrate and how much per year? (kg)?')
        frame14.grid_forget()
        frame140.grid_forget()
        frame16.grid(sticky=W)
        frame160.grid(sticky=W)
    if count == 9:
        var.set('9. How much water do you use?')
        frame16.grid_forget()
        frame160.grid_forget()
        frame17.grid(sticky=W)
        frame170.grid(sticky=W)
    if count == 10:
        var.set('10. How much pesticides do you use?')
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
        var.set('12. How much waste do you produce? ')
        frame19.grid_forget()
        frame190.grid_forget()
        frame20.grid(sticky=W)
        frame200.grid(sticky=W)
    if count == 13:
        var.set('13. How far does your produce travel to the distribution center \non average? ')
        frame20.grid_forget()
        frame200.grid_forget()
        frame22.grid(sticky=W)
        frame220.grid(sticky=W)
    if count == 14:
        var.set('14. How much of the product does not survive the transport \nstage to the store? ')
        frame22.grid_forget()
        frame220.grid_forget()
        frame23.grid(sticky=W)
        frame230.grid(sticky=W)
    if count == 15:
        var.set('This is was the questionnaire, are you finished?')
        frame23.grid_forget()
        frame230.grid_forget()
        frame_finish.grid(sticky=W)
        shitlabel2 = Label(frame_finish, text='                                                            ').grid(row=0, column=0)
        Button_finish = Button(frame_finish, text=('finish'), command=close_program)
        Button_finish.grid(row=0, column=1, padx=10)

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
    fracLetsur = fracEndsur = fracSpisur = fracBeasur = fracParsur = fracKalsur = fracBassur = fracRucsur = fracMicsur = 0
    frac_sur = [fracLetsur, fracEndsur, fracSpisur, fracBeasur, fracParsur, fracKalsur, fracBassur, fracRucsur,
                fracMicsur]
    fracLetkg = fracEndkg = fracSpikg = fracBeakg = fracParkg = fracKalkg = fracBaskg = fracRuckg = fracMickg = 0
    frac_kg = [fracLetkg, fracEndkg, fracSpikg, fracBeakg, fracParkg, fracKalkg, fracBaskg, fracRuckg, fracMickg]
    for i in range(0, len(frac_sur)):
        frac_sur[i] = surVeg[i].get() / total_area
        frac_kg[i] = kgVeg[i].get() / total_kg

    # Creating a dictionary of all parameters: [fraction surface, fraction kg,kg vegetation]
    dic_crops = {}
    dic_crops['Total'] = [1, 1, total_kg]
    for i in range(0, len(frac_sur)):
        dic_crops[lis[i]] = [frac_sur[i], frac_kg[i], kgVeg[i].get()]
    dic_crops = {x: y for x, y in dic_crops.items() if y != [0, 0, 0]}
    print(dic_crops)
    return dic_crops


# The function cal checks whether a percentage value (Q6) is between 0 and 100
# It works but you can also ignore the notifications and go to the next question
dd = 2


def cal(event):
    try:
        if 0 <= ans916.get() <= 100 and 0 <= ans915.get() <= 100 and 0 <= ans914.get() <= 100 and 0 <= ans913.get() <= 100 and 0 <= ans912.get() <= 100 and 0 <= ans911.get() <= 100 and 0 <= ans910.get() <= 100 and 0 <= ans99.get() <= 100 and 0 <= ans98.get() <= 100 and 0 <= ans97.get() <= 100:
            other = 100 - ans916.get() - ans915.get() - ans914.get() - ans913.get() - ans912.get() - ans911.get() - ans910.get() - ans99.get() - ans98.get() - ans97.get()
            if other >= 0:
                ans917.set(other)
            else:
                messagebox.showinfo('Notification', 'The range of the number should be (0,100)')
        else:
            global dd
            dd += 1
            if dd % 3 == 0:
                messagebox.showinfo('Notification', 'The range of the number should be (0,100)')

    except TclError:
        for i in range(len(list_ans)):  # if there is no value in Entry, make it back to 0
            try:
                if i != 0 or 2 or 1:
                    list_ans[i].get() != ''

            except TclError:
                list_ans[i].set(00)
    return

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

# Enter farm's name
frame0.pack(anchor=CENTER)
farm_name = StringVar()
Button(frame00, text='Next', command=next2).pack(fill=BOTH, side=BOTTOM, anchor=CENTER)
Entry(frame00, textvariable=farm_name).pack(fill=BOTH, side=BOTTOM, anchor=CENTER)
Label(frame00, text='\n\n\n\nEnter the name of your farm').pack(fill=BOTH, side=BOTTOM)

# Basic frame containing previous and next labels
button2 = Button(frame2, text=('previous'), command=pre).grid(row=0, column=0, padx=10)
shitlabel = Label(frame2, text='                                   ').grid(row=0, column=1)
button1 = Button(frame2, text=('next'), command=next1).grid(row=0, column=2, sticky=E,padx=10)
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
country['values'] = ('Netherlands', 'China', 'Germany')
country.current(0)
country.grid(padx=10)

# Here a list of all the possible crops a farmer can choose is read in. This is needed for Q2.
wb = xlrd.open_workbook('Crops energy content.xlsx')
lis = []
database = wb.sheet_by_name('Basic database')
for i in range(1, len(database.col_values(2))):
    if database.col_values(2)[i] == "":
        break
    lis.append(database.col_values(2)[i])

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
ansVeg = [ansLet, ansEnd, ansSpi, ansBea, ansPar, ansKal, ansBas, ansRuc, ansMic]

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
surVeg = [surLet, surEnd, surSpi, surBea, surPar, surKal, surBas, surRuc, surMic]

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
kgVeg = [kgLet, kgEnd, kgSpi, kgBea, kgPar, kgKal, kgBas, kgRuc, kgMic]

Label(frame5, text='Crop [-]').grid(row=0, column=0, padx=5, sticky=W)
Label(frame5, text='Surface [m2]').grid(row=0, column=1, padx=5, sticky=W)
Label(frame5, text='Sold products [kg/year]').grid(row=0, column=2, padx=5, sticky=W)

# In this for loop, the fields for Q2 are created
for i in range(0, len(lis)):
    Checkbutton(frame5, text=lis[i], variable=ansVeg[i]).grid(row=i + 1, column=0, sticky=W, padx=5)
    EntSur = Entry(frame5, textvariable=surVeg[i])
    EntSur.grid(row=i + 1, column=1, sticky=W, padx=5, pady=5)
    Entkg = Entry(frame5, textvariable=kgVeg[i])
    Entkg.grid(row=i + 1, column=2, sticky=W, padx=5, pady=5)
    EntSur.bind('<FocusOut>', cal2)
    Entkg.bind('<FocusOut>', cal2)
    EntSur.bind("<Button-1>", rid_of_zeros_sur)
    Entkg.bind("<Button-1>", rid_of_zeros_kg)

# Here the fields for question 3 (buying electricity) are created
ans61 = IntVar()
ans62 = IntVar()
greenlabel = Label(frame8, text='Renewable (kWh)').grid(row=1, column=0, padx=20, sticky=W)
greenentry = Entry(frame8, width=10, textvariable=ans61).grid(row=1, column=1)
greylabel = Label(frame8, text='Non-renewable(kWh)').grid(row=2, column=0, padx=20, sticky=W)
greyentry = Entry(frame8, width=10, textvariable=ans62).grid(row=2, column=1)

# Here the fields for question 4 (creation of renewable energy) are created
ans71 = IntVar()
ans72 = IntVar()
ans73 = IntVar()
solarlabel = Label(frame9, text='Solar energy (kWh)').grid(row=1, column=0, padx=20, sticky=W)
solarentry = Entry(frame9, width=10, textvariable=ans71).grid(row=1, column=1)
biomasslabel = Label(frame9, text='Biomass (kWh)').grid(row=2, column=0, padx=20, sticky=W)
biomassentry = Entry(frame9, width=10, textvariable=ans72).grid(row=2, column=1)
windlabel = Label(frame9, text='Windpower (kWh)').grid(row=3, column=0, padx=20, sticky=W)
windentry = Entry(frame9, width=10, textvariable=ans73).grid(row=3, column=1)

# Here the fields for Q5 (how electricity is used) are created
ans81 = IntVar()
ans82 = IntVar()
ans83 = IntVar()
ans84 = IntVar()
ans85 = IntVar()
ans86 = IntVar()
ans87 = IntVar()
ans88 = IntVar()
ans89 = IntVar()
ans890 = IntVar()
heatingl = Label(frame10, text='Heating (kWh)').grid(row=1, column=0, padx=5, sticky=W)
heatingy = Entry(frame10, width=10, textvariable=ans81).grid(row=1, column=1)
coolingl = Label(frame10, text='Cooling (kWh)').grid(row=2, column=0, padx=5, sticky=W)
coolingy = Entry(frame10, width=10, textvariable=ans82).grid(row=2, column=1)
ventillationl = Label(frame10, text='Ventillation (kWh)').grid(row=3, column=0, padx=5, sticky=W)
ventillationy = Entry(frame10, width=10, textvariable=ans83).grid(row=3, column=1)
lightingl = Label(frame10, text='Lighting (kWh)').grid(row=1, column=2, padx=5, sticky=W)
lightingy = Entry(frame10, width=10, textvariable=ans84).grid(row=1, column=3)
machineryl = Label(frame10, text='Machinery (kWh)').grid(row=2, column=2, padx=5, sticky=W)
machineryy = Entry(frame10, width=10, textvariable=ans85).grid(row=2, column=3)
storagel = Label(frame10, text='Storage (kWh)').grid(row=3, column=2, padx=5, sticky=W)
storagey = Entry(frame10, width=10, textvariable=ans86).grid(row=3, column=3)
Label(frame100, text='Selling renewable(kWh)').grid(row=0, column=0, sticky=W, padx=5)
Entry(frame100, width=10, textvariable=ans87).grid(row=0, column=1)
Label(frame100, text='Selling non-renewable(kWh)').grid(row=1, column=0, sticky=W, padx=5)
Entry(frame100, width=10, textvariable=ans88).grid(row=1, column=1)
Label(frame100, text='Other(kWh)').grid(row=2, column=0, sticky=W, padx=5)
Entry(frame100, width=10, textvariable=ans89).grid(row=2, column=1)
Checkbutton(frame100, text='I don\'t know', variable=ans890).grid(row=3, column=0, sticky=W, padx=5)

# Here the fields for Q6 (fossil fuel use) are created
ans91 = IntVar()
ans92 = IntVar()
ans93 = IntVar()
ans94 = IntVar()
ans95 = IntVar()
ans96 = IntVar()
ans97 = IntVar()
ans98 = IntVar()
ans99 = IntVar()
ans910 = IntVar()
ans911 = IntVar()
ans912 = IntVar()
ans913 = IntVar()
ans914 = IntVar()
ans915 = IntVar()
ans916 = IntVar()
ans917 = IntVar()
# 'other' is also calculated in line 216 and 424...
other = 100 - ans916.get() - ans915.get() - ans914.get() - ans913.get() - ans912.get() - ans911.get() - ans910.get() - ans99.get() - ans98.get() - ans97.get()
ans917.set(other)
petroll = Label(frame11, text='Petrol (L)').grid(row=0, column=0, padx=5, sticky=W)
petroly = Entry(frame11, width=5, textvariable=ans91).grid(row=0, column=1)
diesell = Label(frame11, text='Diesel (L)').grid(row=1, column=0, padx=5, sticky=W)
diesely = Entry(frame11, width=5, textvariable=ans92).grid(row=1, column=1)
Ngasl = Label(frame11, text='Natural gas (L)').grid(row=0, column=2, padx=10, sticky=W)
Ngasy = Entry(frame11, width=5, textvariable=ans94).grid(row=0, column=3)
oill = Label(frame11, text='Oil (L)').grid(row=1, column=2, padx=10, sticky=W)
oily = Entry(frame11, width=5, textvariable=ans95).grid(row=1, column=3)

# Here, additional fields for Q6 (estimating the percentages of fossil fuel use) are build
Label(frame111, text='Estimate in percentages what the fossil fuels are used for').grid(row=0, column=0, sticky=W,
                                                                                        padx=5)
Label(frame110, text='Heating').grid(row=1, column=0, sticky=W, padx=5)
q = Entry(frame110, width=5, textvariable=ans97)
q.grid(row=1, column=1)
Label(frame110, text='%').grid(row=1, column=2, sticky=W)
q.bind('<FocusOut>', cal)
Label(frame110, text='Cooling').grid(row=2, column=0, sticky=W, padx=5)
w = Entry(frame110, width=5, textvariable=ans98)
w.grid(row=2, column=1)
Label(frame110, text='%').grid(row=2, column=2, sticky=W)
w.bind('<FocusOut>', cal)
Label(frame110, text='Electricity').grid(row=3, column=0, sticky=W, padx=5)
e = Entry(frame110, width=5, textvariable=ans99)
e.grid(row=3, column=1)
Label(frame110, text='%').grid(row=3, column=2, sticky=W)
e.bind('<FocusOut>', cal)
Label(frame110, text='Tillage').grid(row=4, column=0, sticky=W, padx=5)
a = Entry(frame110, width=5, textvariable=ans910)
a.grid(row=4, column=1)
Label(frame110, text='%').grid(row=4, column=2, sticky=W)
a.bind('<FocusOut>', cal)
Label(frame110, text='Sowing').grid(row=5, column=0, sticky=W, padx=5)
s = Entry(frame110, width=5, textvariable=ans911)
s.grid(row=5, column=1)
Label(frame110, text='%').grid(row=5, column=2, sticky=W)
s.bind('<FocusOut>', cal)
Label(frame110, text='Weeding').grid(row=5, column=3, sticky=W, padx=20)
d = Entry(frame110, width=5, textvariable=ans912)
d.grid(row=5, column=4)
Label(frame110, text='%').grid(row=5, column=5, sticky=W)
d.bind('<FocusOut>', cal)
Label(frame110, text='Harvest').grid(row=1, column=3, sticky=W, padx=20)
z = Entry(frame110, width=5, textvariable=ans913)
z.grid(row=1, column=4)
Label(frame110, text='%').grid(row=1, column=5, sticky=W)
z.bind('<FocusOut>', cal)
Label(frame110, text='Fertilize').grid(row=2, column=3, sticky=W, padx=20)
x = Entry(frame110, width=5, textvariable=ans914)
x.grid(row=2, column=4)
Label(frame110, text='%').grid(row=2, column=5, sticky=W)
x.bind('<FocusOut>', cal)
Label(frame110, text='Irrigation').grid(row=3, column=3, sticky=W, padx=20)
c = Entry(frame110, width=5, textvariable=ans915)
c.grid(row=3, column=4)
Label(frame110, text='%').grid(row=3, column=5, sticky=W)
c.bind('<FocusOut>', cal)
Label(frame110, text='Pesticide').grid(row=4, column=3, sticky=W, padx=20)
r = Entry(frame110, width=5, textvariable=ans916)
r.grid(row=4, column=4)
Label(frame110, text='%').grid(row=4, column=5, sticky=W)
r.bind('<FocusOut>', cal)
Label(frame110, text='Other').grid(row=6, column=0, sticky=W, padx=5)
Entry(frame110, width=5, textvariable=ans917).grid(row=6, column=1)
Label(frame110, text='%').grid(row=6, column=2, sticky=W)

# Here the field for fertilizer use are created (Q7)
ans121 = IntVar()
ans122 = IntVar()
ans123 = IntVar()
ans124 = IntVar()
ans125 = IntVar()
ans126 = IntVar()
ans127 = IntVar()
ans128 = IntVar()
ans129 = IntVar()
ans130 = IntVar()
ans131 = IntVar()
ans132 = IntVar()
ans133 = IntVar()
Label(frame14, text='Ammoniumnitrate (kg)').grid(row=1, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans121).grid(row=1, column=1)
Label(frame14, text='Calciumammoniumnitrate (kg)').grid(row=2, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans122).grid(row=2, column=1)
Label(frame14, text='Ammoniumsulphate (kg)').grid(row=3, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans123).grid(row=3, column=1)
Label(frame14, text='Triplesuperphosphate (kg)').grid(row=4, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans124).grid(row=4, column=1)
Label(frame14, text='Single super phosphate (kg)').grid(row=5, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans125).grid(row=5, column=1)
Label(frame14, text='Ammonia(kg)').grid(row=6, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans126).grid(row=6, column=1)
Label(frame14, text='limestone (kg)').grid(row=7, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans127).grid(row=7, column=1)
Label(frame14, text='NPK 15-15-15(kg)').grid(row=8, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans128).grid(row=8, column=1)
Label(frame14, text='Urea(kg)').grid(row=9, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans129).grid(row=9, column=1)
Label(frame14, text='Manure(kg)').grid(row=10, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans130).grid(row=10, column=1)
Label(frame14, text='Phosphoric acid(kg)').grid(row=11, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans131).grid(row=11, column=1)
Label(frame14, text='mono-ammonium phosphate(kg)').grid(row=12, column=0, padx=5, sticky=W)
Entry(frame14, width=10, textvariable=ans132).grid(row=12, column=1)
Checkbutton(frame140, text='I don\'t know', variable=ans133).grid(padx=5)

# Here the fields for substrate use (Q8) are created
ans141 = IntVar()
ans142 = IntVar()
ans143 = IntVar()
ans144 = IntVar()
ans145 = IntVar()
ans146 = IntVar()
ans147 = IntVar()
Label(frame16, text='Rockwool(kg)').grid(row=1, column=0, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans141).grid(row=1, column=1)
Label(frame16, text='Perlite(kg)').grid(row=2, column=0, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans142).grid(row=2, column=1)
Label(frame16, text='Cocofiber(kg)').grid(row=1, column=2, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans143).grid(row=1, column=3)
Label(frame16, text='hemp fiber(kg)').grid(row=2, column=2, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans144).grid(row=2, column=3)
Label(frame16, text='Peat(kg)').grid(row=3, column=0, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans145).grid(row=3, column=1)
Label(frame16, text='Peat Moss Kg)').grid(row=3, column=2, padx=5, sticky=W)
Entry(frame16, width=10, textvariable=ans146).grid(row=3, column=3)
Checkbutton(frame160, text='No substrate is used', variable=ans147).grid(padx=5)

# Here the fields for water use (Q9) are created
ans151 = IntVar()
ans152 = IntVar()
Label(frame17, text='Tap water(L)').grid(row=1, column=0, padx=5, sticky=W)
Entry(frame17, width=10, textvariable=ans151).grid(row=1, column=1)
Checkbutton(frame170, text='I don\'t know', variable=ans152).grid(sticky=W, padx=5)

# Here the fields for pesticide use (Q10) are created
ans161 = IntVar()
ans162 = IntVar()
ans163 = IntVar()
ans164 = IntVar()
ans165 = IntVar()
ans166 = IntVar()
Label(frame18, text='Atrazine(kg)').grid(row=1, column=0, padx=5, sticky=W)
Entry(frame18, width=10, textvariable=ans161).grid(row=1, column=1)
Label(frame18, text='Glyphosphate(kg)').grid(row=2, column=0, padx=5, sticky=W)
Entry(frame18, width=10, textvariable=ans162).grid(row=2, column=1)
Label(frame18, text='Metolachlor(kg)').grid(row=3, column=0, padx=5, sticky=W)
Entry(frame18, width=10, textvariable=ans163).grid(row=3, column=1)
Label(frame18, text='Herbicide(kg)').grid(row=4, column=0, padx=5, sticky=W)
Entry(frame18, width=10, textvariable=ans164).grid(row=4, column=1)
Label(frame18, text='Insectiside(kg)').grid(row=5, column=0, padx=5, sticky=W)
Entry(frame18, width=10, textvariable=ans165).grid(row=5, column=1)
Checkbutton(frame180, text='I don\'t know', variable=ans166).grid(sticky=W, padx=5)

# Here the fields for packaging (Q11) are created
ans171 = IntVar()
ans172 = IntVar()
ans173 = IntVar()
Radiobutton(frame19, text='Yes, it is', variable=ans171, value=1).grid(sticky=W, padx=5)
Radiobutton(frame19, text='No, it isn\'t', variable=ans171, value=2).grid(sticky=W, padx=5)

# Here the fields for waste production (Q12) are created
ans181 = IntVar()
ans182 = IntVar()
ans183 = IntVar()
ans184 = IntVar()
Label(frame20, text='Green waste(kg)').grid(row=1, column=0, padx=5, sticky=W)
Entry(frame20, width=10, textvariable=ans181).grid(row=1, column=1)
Label(frame20, text='Gray waste(kg)').grid(row=2, column=0, padx=5, sticky=W)
Entry(frame20, width=10, textvariable=ans182).grid(row=2, column=1)
Label(frame20, text='Paper(kg)').grid(row=3, column=0, padx=5, sticky=W)
Entry(frame20, width=10, textvariable=ans183).grid(row=3, column=1)
Checkbutton(frame200, text='I don\'t know', variable=ans184).grid(sticky=W, padx=5)

# Here the fields for transportation (Q13)are created
ans201 = IntVar()
ans202 = IntVar()
ans203 = IntVar()
ans204 = IntVar()
Label(frame22, text='Plane (km)').grid(row=1, column=0, padx=40, sticky=W)
Entry(frame22, width=10, textvariable=ans201).grid(row=1, column=1)
Label(frame22, text='Truck (km)').grid(row=2, column=0, padx=40, sticky=W)
Entry(frame22, width=10, textvariable=ans202).grid(row=2, column=1)
Label(frame22, text='Ship (km)').grid(row=3, column=0, padx=40, sticky=W)
Entry(frame22, width=10, textvariable=ans203).grid(row=3, column=1)
Checkbutton(frame220, text='I don\'t know', variable=ans204).grid(sticky=W, padx=40)

# Here the fields for transport losses are created
ans211 = IntVar()
ans212 = IntVar()
Entry(frame23, width=5, textvariable=ans211).grid(row=0, column=0, padx=20, pady=20)
Label(frame23, text='%').grid(row=0, column=1, sticky=W)
Checkbutton(frame230, text='I don\'t know', variable=ans212).grid(sticky=W, padx=40)

# At the end, a list containing all the variables is created. It is needed to be able to load previously filled in results
list_ans = [farm_name, ans1, v, ans61, ans62, ans71, ans72, ans73, ans81, ans82, ans83, ans84, ans85, ans86, ans87,
            ans88, ans89, ans890, ans91, ans92, ans94, ans95, ans97, ans98, ans99, ans910, ans911, ans912, ans913,
            ans914, ans915, ans916, ans917, ans121, ans122, ans123, ans124, ans125, ans126, ans127, ans128, ans129,
            ans130, ans131, ans132, ans133, ans141, ans142, ans143, ans144, ans145, ans146, ans147, ans151, ans152,
            ans161, ans162, ans163, ans164, ans165, ans166, ans171, ans172, ans181, ans182, ans183, ans184, ans201,
            ans202, ans203, ans204, ans211, ans212]

# Important statement. If not placed here, program crashes. Assures that all information from above is in the program
root.mainloop()
