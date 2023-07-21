"""
NESC Clearance calculation based on C2-2017
Joseph Dong
"""
import os, math, json, copy, re
from datetime import datetime

from .go_95 import go_95_main
from ..commons import report_xlsx_general, common_utils


class GlobalVar():
    color_bkg_header = 'E0E0E0'
    color_bkg_data_1 = 'CCE5FF'
    color_bkg_data_2 = 'CCFFFF'
    color_bkg_special = 'FF0000'


class DataClearance(object):
    dict_help = {
        'description': 'List of tables: table_name <-> table_description',
        'url_format': '<base>/api/tools/clearances/table_name',
        'method': 'GET',
        'list_of_tables': [
            {'T232-1': 'Vertical clearance above ground, roadway, rail, or water surfaces (C2-2017 NESC RULES 232 A, B, C)'},
            {'T233-1 (Vertival)': 'Vertical clearance between wires/conductors carried on different supporting structures (C2-2017 NESC RULES 233A, C)'},
            {'T233-1 (Horizontal)': 'Horizontal clearance between wires/conductors carried on different supporting structures (C2-2017 NESC RULES 233A, B)'},
            {'T234-1': 'Clearance from buildings and other installations except bridges (C2-2017 NESC RULES 234 A, C, G)'},
            {'T234-2': 'Clearance from bridges and over or near swimming pools (C2-2017 NESC RULES 234 A, D, E, G)'},
            {'T234-Other': 'Clearance from other supporting structures, grain bins and rail cars (C2-2017 NESC RULES 234 A, B, F, G, I)'},
            {'T235-6 (Structure)': ' Clearances from any direction to structure components (C2-2017 NESC RULES 235-A, E1, E2, I)'},
            {'T232 Alternate': 'Vertical clearance above ground, etc. by using the alternate method (C2-2017 NESC RULES 232A, D)'},
            {'T233 Alternate (Vert.)': 'Vertical clearance between wires/conductors carried on different supporting structures by using the alternate method (C2-2017 NESC RULES 233C3b & assC3c)'},
            {'T233 Alternate (Horiz.)': 'Horizontal clearance between wires/conductors carried on different supporting structures by using the alternate method (C2-2017 NESC RULES 233-B2 & 235-B.3.a, 235-B.3.b)'},
            {'T234 Alternate': 'Clearance from objects (e.g. buildings, bridges) by using the alternate method (C2-2017 NESC RULES 234H)'},
            {'T235-6 Alternate': 'Clearances from any direction to structure components (Alternate Method)'},
            {'Table 1 (Rule 37)': 'Basic Minimum Allowable Vertical Clearance of Wires above Railroads, Thoroughfares, '
                                  'Ground or Water Surfaces; Also Clearances from Poles, Buildings, Structures or Other Objects'},
            {'Table 2 (Rule 38)': 'Basic Minimum Allowable Clearance of Wires from Other Wires at Crossings, '
                                  'in Midspans and at Supports (Clearances are in feet); '
                                  'Non-High Fire Threat District (10% Reduction for Reduced Design Values)'},
            {'Table 2 (Rule 38 High Fire Threat)': 'Basic Minimum Allowable Clearance of Wires from Other Wires at '
                                                   'Crossings, in Midspans and at Supports (Clearances are in feet); '
                                                   'High Fire Threat District (5% Reduction for Reduced Design Values)'},
            {'Table 2A (Rule 39)': 'Minimum Clearances of Wires from Signs Mounted on Buildings and '
                                   'Isolated Structures; All Clearances are in feet'},
            {'swimming Pools': 'Minimum Vertical and Redial clearances over swimming pools'}
        ]
    }

    table_input = {
        'system_type': '0',
        'system_voltage': '500',
        'factor_voltage': '1.10',
        'terrain_ele': '4500',
        'margin_vert': '3.0',
        'margin_hori': '2.0',
        'round_upto': '0.5',
        'mot_temperature': '202',
        'factor_switching': '2.2',
        'form_id': 'form_clearance_input',
        'item_name': None
    }

    rule235_input = {
        "angle_swing": "30.0",
        "circuit1": "500",
        "circuit2": "500",
        "elevation_ter": "4500",
        "factor_experience": "0",
        "factor_pu": "2.2",
        "len_insulator": "6.0",
        "same_circuit": "0",
        "same_utility": "0",
        "type_system1": "0",
        "type_system2": "0",
        "value_sag": "20.0",
        "vol_multipler1": "1.1",
        "vol_multipler2": "1.1"
    }

    proj_info = {
    "date_check": "",
    "date_design": "",
    "form_id": "form_clearance_proj",
    "proj_checker": "",
    "proj_engineer": "",
    "proj_name": "test",
    "proj_note": "",
    "proj_numb": "12345",
    "proj_subj": "",
    }

    params = {}

    table_header_nesc = ['Nature of Surface',
                         '<html>NESC<BR>Base (ft)</html>',
                         '<html>Voltage<BR>Adder (ft)</html>',
                         '<html>Altitude<BR>Adder (ft)</html>',
                         '<html>NESC<BR>Total (ft)</html>',
                         '<html>Design<BR>Total (ft)</html>']
    table_header_nesc_alt = ['Nature of Surface',
                             '<html>Reference<BR>Height<BR>(ft)</html>',
                             '<html>Electrical<BR>Component<BR>(ft)</html>',
                             '<html>Altitude<BR>Adder<BR>(ft)</html>',
                             '<html>NESC<BR>Total<BR>(ft)</html>',
                             '<html>Design<BR>Total<BR>(ft)</html>']

    def __init__(self, params=None):
        if params:
            self.params = params

    def get_params(self):
        return self.params


    def retrieve_form_data_to_dict(self, form_data):
        form_id = form_data['form_id']

        if 'form_nesc235' in form_id:
            params = _retrieve_form_data_nesc235(form_data)
        elif 'form_clearance_proj' in form_id:
            params = _retrieve_form_data_proj(form_data=form_data)
        elif 'form_clearance_input' in form_id:
            params = _retrieve_form_data_input(form_data=form_data)
        else:
            params = None

        return params

    def get_table_clearance(self, params, table_desc=None):
        if table_desc is None:
            table_desc = params['item_name']

        # conditions for using alternate method
        voltage_sys = get_float(params['system_voltage'], value_default=0.0)
        system_type = params['system_type']  # 0-AC, 1-DC
        factor_vol = get_float(params['factor_voltage'], value_default=1.0)
        voltage_max = _get_voltage_max(system_voltage=voltage_sys, factor_vol=factor_vol, is_dc=(system_type==1))
        voltage_gnd = voltage_max / math.sqrt(3) if params['system_type'] == '0' else voltage_max / math.sqrt(2)

        if voltage_gnd > 470:
            alternate_method = 2
        elif voltage_gnd > 98:
            alternate_method = 1
        else:
            alternate_method = 0

        if 't232-1' in table_desc.lower():
            return get_table_232_norm(self.table_header_nesc, params, alternate_method=alternate_method)
        elif 't233-1 (vert' in table_desc.lower():
            return get_table_233_normVH(self.table_header_nesc, params, is_horiz=False, alternate_method=alternate_method)
        elif 't233-1 (hori' in table_desc.lower():
            return get_table_233_normVH(self.table_header_nesc, params, is_horiz=True, alternate_method=alternate_method)
        elif 't234-1' in table_desc.lower():
            return get_table_234_1_norm(self.table_header_nesc, params, alternate_method=alternate_method)
        elif 't234-2' in table_desc.lower():
            return get_table_234_2_norm(self.table_header_nesc, params, alternate_method=alternate_method)
        elif 't234-other' in table_desc.lower():
            return get_table_234_other_norm(self.table_header_nesc, params, alternate_method=alternate_method)
        elif 't235-6 (structure)' in table_desc.lower():
            return get_table_235_6_norm(self.table_header_nesc, params, alternate_method=alternate_method)
        elif 't232 alt' in table_desc.lower():
            return get_table_232_alt(self.table_header_nesc_alt, params, alternate_method=alternate_method)
        elif 't233 alternate (vert.)' in table_desc.lower():
            return get_table_233_alter(self.table_header_nesc_alt, params, alternate_method=alternate_method)
        elif 't233 alternate (horiz.)' in table_desc.lower():
            return get_table_233_alter(self.table_header_nesc_alt, params, is_hori=True, alternate_method=alternate_method)
        elif 't234 alternate' in table_desc.lower():
            return get_table_234_alter(self.table_header_nesc_alt, params, alternate_method=alternate_method)
        else:
            return None

    def get_tables_spreadsheet(self, path_params):
        filename_xlsx = 'tmp/' + datetime.utcnow().strftime('utc_%Y%m%d_%H%M%S') + '_clearance.xlsx'

        with open(path_params, 'r') as fh:
            params_dict = json.load(fh)
            list_err_input = _verify_json_spreadsheet(params_dict=params_dict)
            if list_err_input:
                return {'error_msg': list_err_input,
                        'filename': None}

            ws_cover = get_ws_coversheet(params_dict=params_dict)
            list_ws_nesc = get_ws_spreadsheet_nesc(self, params_dict=params_dict)
            list_ws_go95 = get_ws_spreadsheet_go95(self, params_dict=params_dict)
            list_ws = list_ws_nesc + list_ws_go95
            list_ws.insert(0, ws_cover)
            report_xlsx_general.create_workbook(workbook_content=list_ws, filename=filename_xlsx,
                                                add_footer_index=False)

        return {'filename': filename_xlsx}


def get_ws_coversheet(params_dict):
    color_bkg_header = GlobalVar.color_bkg_header
    color_bkg_data1 = GlobalVar.color_bkg_data_1
    color_bkg_data2 = GlobalVar.color_bkg_data_2

    ws_data = [['CLEARANCE REPORTS'], ['For Selected Tables (NESC and/or GO-95)']]
    ws_data.append(['Selected NESC Tables\n' + ', '.join(params_dict.get('list_select_nesc'))])
    ws_data.append(['Selected GO-95 Tables\n' + ', '.join(params_dict.get('list_select_go95'))])
    ws_data.append(['Inputs Utilized for Clearance Calculations Are As Follows'])
    system_type = ['AC', 'DC']
    system_voltage = params_dict.get('system_voltage') + ' (' + system_type[int(params_dict.get('system_type'))] + ')'
    ws_data.append(['System Voltage: ', system_voltage])
    ws_data.append(['Voltage Multiplier: ', params_dict.get('factor_voltage')])
    ws_data.append(['Terrain Elevation: ', params_dict.get('terrain_ele')])
    ws_data.append(['Vertical Margin: ', params_dict.get('margin_vert')])
    ws_data.append(['Horizontal Margin: ', params_dict.get('margin_hori')])
    ws_data.append(['Round Value Upto: ', params_dict.get('round_upto')])
    ws_data.append(['Max. Operating Temperature: ', params_dict.get('mot_temperature')])
    ws_data.append(['Switching Factor: ', params_dict.get('factor_switching')])

    n_row = len(ws_data)
    range_merge = [(i+1, 1, i+1, 2, 'center') for i in range(5)]
    range_color = [(1, 1, 1, 2, color_bkg_header)]
    for i in range(1, n_row):
        if i % 2 == 0:
            range_color.append([i+1, 1, i+1, 2, color_bkg_data1])
        else:
            range_color.append([i + 1, 1, i + 1, 2, color_bkg_data2])

    range_font_bold = [(1, 1, 1, 2)]
    range_font_size = [(1, 1, 1, 2, 16)]
    range_border = [(1, 1, n_row, 2)]
    range_alignment = [(6, 1, n_row, 1, 'right')]

    row_height = [(1, 25), (3, 50), (4, 30), (5, 25)]
    row_height += [(i+1, 20) for i in range(5, n_row)]
    col_width = [(1, 50), (2, 50)]
    cell_range_style = {
        'range_merge': range_merge,
        'range_border': range_border,
        'range_color': range_color,
        'range_font_bold': range_font_bold,
        'range_font_size': range_font_size,
        'range_align': range_alignment,
        'row_height': row_height,
        'column_width': col_width
    }

    ws = {
        'ws_name': 'Cover',
        'ws_content': ws_data,
        'cell_range_style': cell_range_style
    }
    return ws


def get_ws_spreadsheet_nesc(self, params_dict):
    '''
    INPUT: params_dict - it includes the list that defines tables selected for exporting
    RETURN: a list of worksheet content
    '''
    color_bkg_header = GlobalVar.color_bkg_header
    color_bkg_data1 = GlobalVar.color_bkg_data_1
    color_bkg_data2 = GlobalVar.color_bkg_data_2

    list_nesc = params_dict.get('list_select_nesc')
    list_ws_nesc = []
    if list_nesc is None:
        return None
    for table_desc in list_nesc:
        table_content = self.get_table_clearance(params=params_dict, table_desc=table_desc)
        if table_content:
            ws = _convert_table_content_to_ws(table_content=table_content, table_desc=table_desc)
            list_ws_nesc.append(ws)

    return list_ws_nesc


def get_ws_spreadsheet_go95(self, params_dict):
    list_ws_go95 = []
    list_go95 = params_dict.get('list_select_go95')
    if list_go95 is not None:
        for table_desc in list_go95:
            params_dict.update({'item_name': table_desc})
            user_input = {'table_input': params_dict}
            table_content = go_95_main.get_go_95_table(user_input=user_input)
            if table_content:
                ws = _convert_table_content_to_ws(table_content=table_content, table_desc=table_desc)
                list_ws_go95.append(ws)
    return list_ws_go95


def get_table_232_norm(table_header, params, alternate_method=0):
    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    terrain_ele = get_float(params['terrain_ele'], 0.0)
    round_upto = get_float(params['round_upto'], 0.5)
    margin_v = get_float(params['margin_vert'], 0.0)

    title = ['Table 232-1. Vertical clearance above ground, roadway, rail, or water surfaces '
             '(C2-2017 NESC RULES 232 A, B, C)']

    items = ["<html>1. Track rails of railroads</html>",
             "<html>2. Roads, streets, and other areas subject to truck traffic</html>",
             "<html>3. Driveways, parking lots, and alleys</html>",
             "<html>4. Other land traversed by vehicles (cultivated, grazing, forest, orchards, etc)</html>",
             "<html>5. Spaces and ways subject to pedestrians or restricted traffic only</html>",
             "<html>6. Water areas not suitable for sailboating or where it is prohibited</html>",
             "<html>7. Water areas suitable for sailboating with unobstructed area of "
             "(see Note 1):<BR>"
             "&nbsp;&nbsp;&nbsp;&nbsp; a. Less than 20 acres (0.08 km<sup>2</sup>)</html>",
             "<html>&nbsp;&nbsp;&nbsp;&nbsp; b. 20 - 200 acres (0.08 - 0.8 km<sup>2</sup>)</html>",
             "<html>&nbsp;&nbsp;&nbsp;&nbsp; c. 200 - 2000 acres (0.8 - 8 km<sup>2</sup>)</html>",
             "<html>&nbsp;&nbsp;&nbsp;&nbsp; d. over 2000 acres (8 km<sup>2</sup>)</html>",
             '<html>Where wires run along within the limits of highways or '
             'other road right-of-way but do not overhang the roadway:',

             "<html>9. Roads, streets, or alleys</html>",
             "<html>10. Roads where it is unlikely that vehicles<BR>will be crossing under the line</html>"]

    clearance_base = [26.5, 18.5, 18.5, 18.5, 14.5, 17.0, 20.5, 28.5, 34.5, 40.5, 0, 18.5, 16.5]
    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')
    voltage_gnd = voltage_max / math.sqrt(3) if params['system_type'] == '0' else voltage_max / math.sqrt(2)
    adder_vol22 = get_adder_vol_ft(voltage_gnd, 22)
    adder_altitude = get_adder_altitude(terrain_ele, adder_vol22, voltage_max)

    table_footer = ['<html><B>Notes:</B></html>',
                    'For water area with signs for rigging and launching, the clearance values shall be 5 ft greater '
                    'than that shown in Item 7.',
                    'Refer to NESC C2-2017 for clarifying notes, exceptions, and variations to the above mentioned '
                    'areas and clearances.'
                    ]
    if voltage_sys == 69:
        table_footer.append('Maximum voltage is used for 69kV to conform to industrial conservative practice although its line-to-ground voltage is less than 50kV.')

    if alternate_method == 1:
        table_footer.append('Per Rule 232-D, alternate method may be used with the specified voltage.')
    elif alternate_method == 2:
        table_footer.append("Per Rule 232-C.1, the clearance <font color='red'>SHALL BE DETERMINED BY THE "
                            "<B>ALTERNATE</B> METHOD</font>. See Rule 232D and T232-Alternate for details.")
    if alternate_method > 0 and params['system_type'] == '0':
        table_footer.append('Per Rule 232-C.1.c or 232-D.3.c, the 5mA rms rule shall be checked.')

    # add footer index
    _table_footer_index(table_footer)

    data = []
    for i, item in enumerate(items):
        base = clearance_base[i]
        text_base = '{:.2f}'.format(base) if base > 0 else ''
        text_adder22 = '{:.2f}'.format(adder_vol22) if base > 0 else ''
        text_adder_alt = '{:.2f}'.format(adder_altitude) if base > 0 else ''
        clearance_min = base + adder_vol22 + adder_altitude
        clearance_min = get_rounded(clearance_min, n_decimal=2, val_nearest=round_upto)
        clearance_design = clearance_min + margin_v
        # clearance_design = _get_rounded(clearance_design, n_decimal=2, val_nearest=round_upto)
        text_min = '{:.2f}'.format(clearance_min) if base > 0 else ''
        text_design = '{:.2f}'.format(clearance_design) if base > 0 else ''
        data.append([item, text_base, text_adder22, text_adder_alt, text_min, text_design])

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }
    return table_content


def get_table_233_normVH(table_header, params, is_horiz=False, alternate_method=0):
    '''

    :param table_header:
    :param params:
    :param is_horiz:
    :param alternate_method: eligibility of using alternate method according to voltage level
    :return:
    '''
    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    terrain_ele = get_float(params['terrain_ele'], 0.0)
    round_upto = get_float(params['round_upto'], 0.5)
    margin = get_float(params['margin_vert'], 0.0)
    margin_h = get_float(params['margin_hori'], 0.0)

    title = ['Table 233-1V. Vertical clearance between wires/conductors carried on '
             'different supporting structures (C2-2017 NESC RULES 233A, C)']
    items = ['<html>Overhead shield wires, supply guys, span wires & messengers</html>',
             '<html>Communication guys, messengers, conductors etc.</html>',
             '<html>Trolley and electrified railroad conductors</html>',
             '<html>Open supply lines that are effectively grounded or have ground fault relaying:<BR>'
             '&nbsp;&nbsp;&nbsp;&nbsp;  22 kV (phase-to-ground) and below</html>',
             '<html>&nbsp;&nbsp;&nbsp;&nbsp;  46 kV (phase-to-phase)</html>',
             '<html>&nbsp;&nbsp;&nbsp;&nbsp;  69 kV (phase-to-phase)</html>',
             '<html>&nbsp;&nbsp;&nbsp;&nbsp;  115 kV (phase-to-phase)</html>',
             '<html>&nbsp;&nbsp;&nbsp;&nbsp;  138 kV (phase-to-phase)</html>',
             '<html>&nbsp;&nbsp;&nbsp;&nbsp;  161 kV (phase-to-phase)</html>',
             '<html>&nbsp;&nbsp;&nbsp;&nbsp;  230 kV (phase-to-phase)</html>',
             '<html>&nbsp;&nbsp;&nbsp;&nbsp;  345 kV (phase-to-phase)</html>',
             '<html>&nbsp;&nbsp;&nbsp;&nbsp;  500 kV (phase-to-phase)</html>']

    table_footer = ['<html><B>Notes:</B></html>',
                    'Refer to NESC C2-2017 for more details including footnotes, exceptions, and variations associated '
                    'with the specific table.',
                    'Maximum voltage is used for 69kV to conform to industrial conservative practice although '
                    'its line-to-ground voltage is less than 50kV.',
                    'Multiplier used for crossing circuit is 1.05 when voltage <= 345kV and 1.1 when voltage = 500 kV.'
                    ]

    if is_horiz:
        if alternate_method > 0:
            table_footer.append(
                'Per Rule 233-C.2(2) Exception 1, alternate method may be used with the specified voltage')
    else:
        if alternate_method == 1:
            table_footer.append(
                'Per Rule 233-C.2(2) Exception 1, alternate method may be used with the specified voltage')
        elif alternate_method == 2:
            table_footer.append(
                "Per Rule 233-C.2(2) Exception 2, the clearance <font color='red'>SHALL BE DETERMINED BY THE "
                "<B>ALTERNATE</B> METHOD</font>.")

    clearance_base = [2.0, 5.0, 6.0, 2.0, 2.0, 2.0, 2.0, 2.0, 2.0, 2.0, 2.0, 2.0]
    voltage_xing = [0, 0, 0, 22 * math.sqrt(3), 46, 69*1.05, 115 * 1.05, 138 * 1.05, 161 * 1.05, 230 * 1.05, 345 * 1.05,
                    500 * 1.10]

    if is_horiz:
        clearance_base = [5.0, 5.0, 5.0, 5.0, 5.0, 5.0, 5.0, 5.0, 5.0, 5.0, 5.0, 5.0]
        margin = margin_h
        title = ['<html>Table 233-1H. Horizontal clearance between wires/conductors '
                 'carried on different supporting structures (C2-2017 NESC RULES 233A, B)</html>']

    data = []

    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')
    voltage_gnd = voltage_max / math.sqrt(3) if params['system_type'] == '0' else voltage_max / math.sqrt(2)

    red_font_note = False
    for i, item in enumerate(items):
        voltage_total = voltage_gnd + voltage_xing[i] / math.sqrt(3)
        adder_vol22 = get_adder_vol_ft(voltage_total, 22)
        adder_altitude = get_adder_altitude(terrain_ele, adder_vol22, voltage_max)

        base = clearance_base[i]
        text_base = '{:.2f}'.format(base) if base > 0 else ''
        text_adder22 = '{:.2f}'.format(adder_vol22) if base > 0 else ''
        text_adder_alt = '{:.2f}'.format(adder_altitude) if base > 0 else ''
        clearance_min = base + adder_vol22 + adder_altitude
        clearance_min = get_rounded(clearance_min, n_decimal=2, val_nearest=round_upto)
        clearance_design = clearance_min + margin
        text_min = '{:.2f}'.format(clearance_min) if base > 0 else ''
        text_design = '{:.2f}'.format(clearance_design) if base > 0 else ''

        if voltage_total > 470:
            red_font_note = True
            text_adder22 = '<html><span style="color:red;">' + text_adder22 + '</span></html>'
            text_adder_alt = '<html><span style="color:red;">' + text_adder_alt + '</span></html>'
            text_min = '<html><span style="color:red;">' + text_min + '</span></html>'
            text_design = '<html><span style="color:red;">' + text_design + '</span></html>'

        data.append([item, text_base, text_adder22, text_adder_alt, text_min, text_design])

    if red_font_note:
        table_footer.append('<html><span style="color:red;">Red font</span></html> '
                            'indicates values for reference only. Alternate method should be utilized.')

    # add footer index
    _table_footer_index(table_footer)

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }
    return table_content


def get_table_234_1_norm(table_header, params, alternate_method=0):
    '''
    :param alternate_method:
    :param table_header:
    :param params:
    :return:

    ENSURE length of first column not less than 79 chars to get correct column width when displaying
    '''
    space4 = ''
    space60 = ''
    for i in range(4):
        space4 += '&nbsp;'

    for i in range(60):
        space60 += '&nbsp;'

    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    terrain_ele = get_float(params['terrain_ele'], 0.0)
    round_upto = get_float(params['round_upto'], 0.5)
    margin_v = get_float(params['margin_vert'], 0.0)
    margin_h = get_float(params['margin_hori'], 0.0)

    title = ['Table 234-1. Clearance from buildings and other installations except bridges '
             '(C2-2017 NESC RULES 234 A, C, G)']

    items = ["<html><b>1. Buildings</b></html>",
        "<html>" + space4 + "<b>a. Horizontal</b><BR>"
                 + space4 + "(1) To Walls, projections & guarded windows" + space60 + "</html>",
        "<html>" + space4 + "(2) Unguarded windows</html>",
        "<html>" + space4 + "(3) Balconies & areas> accessible to ped.</html>",
        "<html>" + space4 + "(4) Under 6 psf wind displacement</html>",
        "<html>" + space4 + "<b>b. Vertical</b><BR>"
                 + space4 + "(1) Over/under roofs not accessible to ped.</html>",
        "<html>" + space4 + "(2) Over/under balconies/roofs accessible to ped.</html>",
        "<html>" + space4 + "(3) Over roofs accessible to non-truck vehicles</html>",
        "<html>" + space4 + "(4) Over roofs accessible to truck traffic</html>",
        "<html><b>2. Signs, chimneys, billboards, antennas, tanks etc.</b></html>",
        "<html>" + space4 + "<b>a. Horizontal</b><BR>"
                 + space4 + "(1) Portions accessible to ped.</b></html>",
        "<html>" + space4 + "(2) Portions not accessible to ped.</html>",
        "<html>" + space4 + "(3) Under 6 psf wind displacement</html>",
        "<html>" + space4 + "<b>b. Vertical</b><BR>"
                 + space4 + "(1) Over/under surfaces upon which personnel walk</html>",
        "<html>" + space4 + "(2) Over/under other portions</html>"]

    #clearance_base = [0.0, 7.5, 7.5, 7.5, 4.5, 12.5, 13.5, 13.5, 18.5, 0.0, 7.5, 7.5, 4.5, 13.5, 8.0]
    # Update based on NESC-2023: 13.5 ft => 14.5 ft
    clearance_base = [0.0, 7.5, 7.5, 7.5, 4.5, 12.5, 14.5, 14.5, 18.5, 0.0, 7.5, 7.5, 4.5, 14.5, 8.0]
    is_vert = [False, False, False, False, False, True, True, True, True,
            False, False, False, False, True, True]

    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')
    voltage_gnd = voltage_max / math.sqrt(3) if params['system_type'] == '0' else voltage_max / math.sqrt(2)
    adder_vol22 = get_adder_vol_ft(voltage_gnd, 22)
    adder_altitude = get_adder_altitude(terrain_ele, adder_vol22, voltage_max)

    table_footer = ['<html><B>Notes:</B></html>',
                    'Refer to NESC C2-2017 for more details including footnotes, exceptions, and variations associated '
                    'with the specific table.',
                    '<span style="color:red;">'
                    'According to NESC C2-2023, base clearance for items 1b(2), 1b(3), and 2b(1) is changed from 13.5 ft to 14.5 ft </span>'
                    ]
    if alternate_method  == 1:
        table_footer.append('Per Rule 234-G.1 Exception, alternate method may be used with the specified voltage. See Rule 234H and T234-Alternate for details.')

    if voltage_sys == 69:
        table_footer.append('Maximum voltage is used for 69kV to conform to industrial conservative practice although its line-to-ground voltage is less than 50kV.')

    elif alternate_method == 2:
        table_footer.append("Per Rule 234-G.1, the clearance <font color='red'>SHALL BE DETERMINED BY THE "
                            "<B>ALTERNATE</B> METHOD</font>. See Rule 234H and T234-Alternate for details.")
    if alternate_method > 0 and params['system_type'] == '0':
        table_footer.append('Per Rule 234-G.3, the 5mA rms rule shall be checked.')

    # add footer index
    _table_footer_index(table_footer)

    data = []
    for i, item in enumerate(items):
        base = clearance_base[i]

        text_base = '{:.2f}'.format(base) if base > 0 else ''
        text_adder22 = '{:.2f}'.format(adder_vol22) if base > 0 else ''
        text_adder_alt = '{:.2f}'.format(adder_altitude) if base > 0 else ''
        clearance_min = base + adder_vol22 + adder_altitude
        clearance_min = get_rounded(clearance_min, n_decimal=2, val_nearest=round_upto)
        clearance_design = clearance_min + (margin_v if is_vert[i] else margin_h)
        clearance_design = get_rounded(clearance_design, n_decimal=2, val_nearest=round_upto)
        text_min = '{:.2f}'.format(clearance_min) if base > 0 else ''
        text_design = '{:.2f}'.format(clearance_design) if base > 0 else ''
        
        # red font for update according to NESC C2-2023
        if i == 6 or i == 7 or i == 13:
            text_base = '<span style="color:red;">' + text_base + '</span>'
        data.append([item, text_base, text_adder22, text_adder_alt, text_min, text_design])

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }
    return table_content


def get_table_234_2_norm(table_header, params, alternate_method=0):
    '''
        :param alternate_method:
        :param table_header:
        :param params:
        :return:

        ENSURE length of first column not less than 79 chars to get correct column width when displaying
        '''
    space4 = ''
    space60 = ''
    for i in range(4):
        space4 += '&nbsp;'

    for i in range(60):
        space60 += '&nbsp;'

    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    terrain_ele = get_float(params['terrain_ele'], 0.0)
    round_upto = get_float(params['round_upto'], 0.5)
    margin_v = get_float(params['margin_vert'], 0.0)
    margin_h = get_float(params['margin_hori'], 0.0)

    title = ['Table 234-2/3. Clearance from bridges and over or near swimming pools '
             '(C2-2017 NESC RULES 234 A, D, E, G)']

    items = [ "<html><b>1. Over bridges</b></html>",
        "<html>" + space4 + "a. Attached</html>",
        "<html>" + space4 + "b. Not attached</html>",
        "<html><b>2. Beside, under or within bridge structure</b></html>",
        "<html>" + space4 + "a. Readily accessible portions"
                + "" + space4 + "(1) Attached</html>",
        "<html>" + space4 + "(2) Not attached</html>",
        "<html>" + space4 + "b. Ordinarily inaccessible portions of bridges and from abutments<BR>"
                + "" + space4 + "(1) Attached</html>",
        "<html>" + space4 + "(2) Not attached</html>",
        "<html><b>3. Over/near swimming pools</b></html>",
        "<html>" + space4 + "A. Any direction from water level, edge of pool, base of diving platform, or anchored raft</html>",
        "<html>" + space4 + "B. Any direction to diving platform, tower, water slide, or other fixed, pool-related structures</html>"]

    clearance_base = [0.0, 5.5, 12.5, 0.0, 5.5, 7.5, 5.5, 6.5, 0.0, 25.0, 17.0]
    margin_ids = [0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 2]
    margins = []
    for margin_id in margin_ids:
        # margin_id: 0-Vertical, 1-Horizontal, 2-max
        if margin_id == 0:
            margins.append(margin_v)
        elif margin_id == 1:
            margins.append(margin_h)
        else:
            margins.append(max(margin_v, margin_h))

    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')
    voltage_gnd = voltage_max / math.sqrt(3) if params['system_type'] == '0' else voltage_max / math.sqrt(2)
    adder_vol22 = get_adder_vol_ft(voltage_gnd, 22)
    adder_altitude = get_adder_altitude(terrain_ele, adder_vol22, voltage_max)

    table_footer = ['<html><B>Notes:</B></html>',
                    'Refer to NESC C2-2017 for more details including footnotes, exceptions, and variations associated '
                    'with the specific table.',
                    ]
    if alternate_method == 1:
        table_footer.append(
            'Per Rule 234-G.1 Exception, alternate method may be used with the specified voltage. See Rule 234H and T234-Alternate for details.')
    if (voltage_sys == 69):
        table_footer.append('Maximum voltage is used for 69kV to conform to industrial conservative practice although its line-to-ground voltage is less than 50kV.')

    elif alternate_method == 2:
        table_footer.append(
            "Per Rule 234-G.1, the clearance <font color='red'>SHALL BE DETERMINED BY THE "
            "<B>ALTERNATE</B> METHOD</font>. See Rule 234H and T234-Alternate for details.")
    if alternate_method > 0 and params['system_type'] == '0':
        table_footer.append('Per Rule 234-G.3, the 5mA rms rule shall be checked.')

    # add footer index
    _table_footer_index(table_footer)

    data = []
    for i, item in enumerate(items):
        base = clearance_base[i]

        text_base = '{:.2f}'.format(base) if base > 0 else ''
        text_adder22 = '{:.2f}'.format(adder_vol22) if base > 0 else ''
        text_adder_alt = '{:.2f}'.format(adder_altitude) if base > 0 else ''
        clearance_min = base + adder_vol22 + adder_altitude
        clearance_min = get_rounded(clearance_min, n_decimal=2, val_nearest=round_upto)
        clearance_design = clearance_min + margins[i]
        clearance_design = get_rounded(clearance_design, n_decimal=2, val_nearest=round_upto)
        text_min = '{:.2f}'.format(clearance_min) if base > 0 else ''
        text_design = '{:.2f}'.format(clearance_design) if base > 0 else ''
        data.append([item, text_base, text_adder22, text_adder_alt, text_min, text_design])

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }
    return table_content


def get_table_234_other_norm(table_header, params, alternate_method=0):
    '''
        :param alternate_method:
        :param table_header:
        :param params:
        :return:

        ENSURE length of first column not less than xx chars to get correct column width when displaying
    '''
    space4 = ''
    spacemany = ''
    for i in range(4):
        space4 += '&nbsp;'

    for i in range(85):
        spacemany += '&nbsp;'

    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    terrain_ele = get_float(params['terrain_ele'], 0.0)
    round_upto = get_float(params['round_upto'], 0.5)
    margin_v = get_float(params['margin_vert'], 0.0)
    margin_h = get_float(params['margin_hori'], 0.0)

    title = ['Table 234-Other. Clearance from other supporting structures, grain bins and rail cars '
             '(C2-2017 NESC RULES 234 A, B, F, G, I)']

    items = [        "<html><b>234B. Other supporting structures</b></html>",
        "<html>" + space4 + "Horizontal @ rest (see Note 3)" + spacemany + "</html>",
        "<html>" + space4 + "Horizontal @ 6 psf wind</html>",
        "<html>" + space4 + "Vertical</html>",
        "<html><b>234F. Grain bins</b></html>",
        "<html>" + space4 + "<B>1</B>. loaded by permanently installed augers, conveyors, or elevators:<BR>"
                + "" + space4 + "a). All directions above bin from each probe port in roof</html>",
        "<html>" + space4 + "b). Horizontal</html>",
        "<html>" + space4 + "<B>2</B>. loaded by portable augers, conveyors, or elevators:<BR>"
                 + space4 + "a). Vertical above highest filling</html>",
        "<html>" + space4 + "b). Horizontal from loading side: "
                + "<b>height plus 5.5m (18ft)</b></html>",
        "<html>" + space4 + "c). Horizontal from nonloading side</html>",
        "<html><b>243I. Rail cars</b></html>",
        "<html>" + space4 + "a). Vertical (Figure 234-5)</html>",
        "<html>" + space4 + "b). Horizontal (Figure 234-5)</html>"]

    clearance_base = [0.0, 5.0, 4.5, 4.5, 0.0, 18.0, 15.0, 18.0, 0.0, 15.0, 0.0, 6.5, 11.5]
    margin_ids = [0, 1, 1, 0, 0, 2, 1, 0, 1, 1, 0, 0, 1]
    margins = []
    for margin_id in margin_ids:
        # margin_id: 0-Vertical, 1-Horizontal, 2-max
        if margin_id == 0:
            margins.append(margin_v)
        elif margin_id == 1:
            margins.append(margin_h)
        else:
            margins.append(max(margin_v, margin_h))

    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')
    voltage_gnd = voltage_max / math.sqrt(3) if params['system_type'] == '0' else voltage_max / math.sqrt(2)
    adder_vol22 = get_adder_vol_ft(voltage_gnd, 22)
    adder_altitude = get_adder_altitude(terrain_ele, adder_vol22, voltage_max)

    table_footer = ['<html><B>Notes:</B></html>',
                    'Refer to NESC C2-2017 for more details including footnotes, exceptions, and variations associated '
                    'with the specific table.',
                    ]
    if (voltage_sys == 69):
        table_footer.append('Maximum voltage is used for 69kV to conform to industrial conservative practice although its line-to-ground voltage is less than 50kV.')

    if alternate_method == 1:
        table_footer.append(
            'Per Rule 234-G.1 Exception, alternate method may be used with the specified voltage. See Rule 234H and T234-Alternate for details.')
    elif alternate_method == 2:
        table_footer.append(
            "Per Rule 234-G.1, the clearance <font color='red'>SHALL BE DETERMINED BY THE "
            "<B>ALTERNATE</B> METHOD</font>. See Rule 234H and T234-Alternate for details.")
    if alternate_method > 0 and params['system_type'] == '0':
        table_footer.append('Per Rule 234-G.3, the 5mA rms rule shall be checked.')

    # add footer index
    _table_footer_index(table_footer)

    data = []
    for i, item in enumerate(items):
        base = clearance_base[i]

        text_base = '{:.2f}'.format(base) if base > 0 else ''
        text_adder22 = '{:.2f}'.format(adder_vol22) if base > 0 else ''
        text_adder_alt = '{:.2f}'.format(adder_altitude) if base > 0 else ''
        clearance_min = base + adder_vol22 + adder_altitude
        clearance_min = get_rounded(clearance_min, n_decimal=2, val_nearest=round_upto)
        clearance_design = clearance_min + margins[i]
        clearance_design = get_rounded(clearance_design, n_decimal=2, val_nearest=round_upto)
        text_min = '{:.2f}'.format(clearance_min) if base > 0 else ''
        text_design = '{:.2f}'.format(clearance_design) if base > 0 else ''
        data.append([item, text_base, text_adder22, text_adder_alt, text_min, text_design])

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }
    return table_content


def get_table_235_6_norm(table_header, params, alternate_method=0):
    space4 = ''
    spacemany = ''
    for i in range(4):
        space4 += '&nbsp;'

    for i in range(80):
        spacemany += '&nbsp;'

    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    terrain_ele = get_float(params['terrain_ele'], 0.0)
    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')

    title = ['Table 235-6. Clearances from any direction to structure components '
             '(C2-2017 NESC RULES 235-A, E1, E2, I)']
    table_header = ["Description", "NESC<BR>Basic (in)", "Voltage<BR>Adder (in)", "Altitude<BR>Adder (in)",
                    "NESC<BR>Total (in)", "Design*<BR>Total (in)"]
    items = ["1. Vertical/Lateral at support" + spacemany,
             space4 + "a. Same circuit",
             space4 + "b. Other circuit (Note 4)",
             space4 + "c. Communication",
			 "2. Guy wires attached to same structure", space4 + "a. Parallel to line", space4 + "b. Anchor guys", space4 + "c. All other",
			 "3. Surface of support arms",
			 "4. Surface of structures", space4 + "a. jointly used structure", space4 + "b. All other",
			 "5. Service drops", space4 + "a. Communication", space4 + "b. Supply"]

    clearance_base87 = [0,  3,  6,  6, 0, 12,  6,  6,  3, 0,  5, 3, 0,  30, 12]
    clearance_base50 = [0, 13.3, 23, 23, 0, 29, 16, 23, 11, 0, 13, 11, 0, 47, 29]
    factor_adder = [0, 0.25, 0.4, 0.4, 0, 0.4, 0.25, 0.4, 0.2, 0, 0.2, 0.2, 0, 0.4, 0.4]

    table_footer = ['<html><B>Notes:</B></html>',
                    'Design clearance is not included because various design margins may apply.',
                    'Depending on given voltage, the base clearance may be associated with either 8.7kV or 50 kV.',
                    'If suspension insulator is used, clearance in this table shall be adjust to consider swinging.',
                    'See footnote 12 of NESC Table 235-6 for determination of voltage between the conductor involved. '
                    ]
    if voltage_sys >= 50:
        table_footer.append('Item 1a is extended to beyond 50kV but not specified by NESC 235.')

    table_footer.append('According to Table 7-1 of RUS1724E200, this clearance applies under 6psf wind displacement.')

    # add footer index
    _table_footer_index(table_footer)

    data = []
    for i, item in enumerate(items):
        base = clearance_base87[i] if voltage_sys < 50 else clearance_base50[i]
        adder_vol = (voltage_max - 8.7) * factor_adder[i] if voltage_sys < 50 else (voltage_max-50) * factor_adder[i]
        adder_vol = math.ceil(adder_vol)
        adder_alt = get_adder_altitude(terrain_ele, adder_vol, voltage_max)

        text_base = '{:.2f}'.format(base) if base > 0 else ''
        text_adder_vol = '{:.2f}'.format(adder_vol) if base > 0 else ''
        text_adder_alt = '{:.2f}'.format(adder_alt) if base > 0 else ''
        clearance_min = base + adder_vol + adder_alt
        clearance_min = get_rounded(clearance_min, n_decimal=2, val_nearest=0.1)
        text_min = '{:.2f}'.format(clearance_min) if base > 0 else ''
        text_design = ''
        if i == 0:
            text_design = '(Note 1)'
        if i == 1 and voltage_sys >= 50:
            text_design = '(Note 5)'
        if i == 8 or i == 9:
            if voltage_sys >= 50:
                text_design = '(Note 6)'
            else:
                text_design = '(Note 5)'

        data.append([item, text_base, text_adder_vol, text_adder_alt, text_min, text_design])

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }
    return table_content


def get_table_232_alt(table_header, params, alternate_method=1):
    '''
        :param alternate_method:
        :param table_header:
        :param params:
        :return:

        ENSURE length of first column not less than xx chars to get correct column width when displaying
    '''
    space4 = ''
    spacemany = ''
    for i in range(4):
        space4 += '&nbsp;'

    for i in range(85):
        spacemany += '&nbsp;'

    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    terrain_ele = get_float(params['terrain_ele'], 0.0)
    round_upto = get_float(params['round_upto'], 0.5)
    margin_v = get_float(params['margin_vert'], 0.0)
    margin_h = get_float(params['margin_hori'], 0.0)
    factor_switching = get_float(params['factor_switching'], value_default=1.0)

    title = ['T232-Alternate. Vertical clearance above ground, etc. by using the alternate method '
             '(C2-2017 NESC RULES 232A, D)']

    items = [
        "<html>a. Track rails of railroads (except electrified railroads using overhead trolley conductors</html>",
        "<html>b. Streets, alleys, roads, driveways, and parking lots</html>",
        "<html>c. Spaces & ways subject to pedestrians or restricted traffic only</html>",
        "<html>d. Other land, such as cultivated, grazing, forest, or orchard, that is traversed by vehicles</html>",
        "<html>e. Water areas not suitable for sailboating or where it is prohibited</html>",
        "<html>f. Water areas suitable for sailboating with unobstructed area of:<BR> " +
        "" + space4 + "(1). Less than 20 acres (0.08 km<sup>2</sup>)</html>",
        "<html>" + space4 + "(2). 20 - 200 acres (0.08 - 0.8 km<sup>2</sup>)</html>",
        "<html>" + space4 + "(3). 200 - 2000 acres (0.8 - 8 km<sup>2</sup>)</html>",
        "<html>" + space4 + "(4). over 2000 acres (8 km<sup>2</sup>)</html>",
        "<html>g. In public or private land and water areas posted for rigging or launching sailboats, the "
        + "reference height shall be 1.5 m (5 ft) greater than in f above.</html>"
    ]

    clearance_base = [22.0, 14.0, 10.0, 14.0, 12.5, 16.0, 24.0, 30.0, 36.0, 0]
    margin_ids = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    margins = []
    for margin_id in margin_ids:
        # margin_id: 0-Vertical, 1-Horizontal, 2-max
        if margin_id == 0:
            margins.append(margin_v)
        elif margin_id == 1:
            margins.append(margin_h)
        else:
            margins.append(max(margin_v, margin_h))

    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')
    if alternate_method == 0:
        adder_elect = -100
    else:
        adder_elect = get_adder_alternate_ft(voltage_max, factor_switching, a=1.15, b=1.03, c=1.2, k=1.15,
                                           vol_l=0, is_dc1=(params['system_type'] == '1'))
    adder_altitude = get_adder_altitude(terrain_ele, adder_elect, voltage_max, base_ele=1500)

    table_footer = ['<html><B>Notes:</B></html>',
                    'Refer to NESC C2-2017 for more details including footnotes, exceptions, and variations associated '
                    'with the specific table.',
                    ]
    if alternate_method == 0:
        table_footer.append("<font color='red'>Per Rule 232-D, the alternate method cannot be used with the specified voltage.</font>")

    elif alternate_method == 1:
        table_footer.append("Per NESC Rule 232D4, the alternate clearance shall be not less than the clearance given"
                            " in Table 232-1 or 232-2 computed for 98 kV ac to ground in accordance with Rule 232C.")
        table_footer.append(
            'Per Rule 232-C.1a Exception, alternate method may be used with the specified voltage. See Rule 234H and T234-Alternate for details.')
    elif alternate_method == 2:
        table_footer.append("Per NESC Rule 232D4, the alternate clearance shall be not less than the clearance given"
                            " in Table 232-1 or 232-2 computed for 98 kV ac to ground in accordance with Rule 232C.")
        table_footer.append(
            'Per Rule 232-C.1a, the clearance SHALL BE DETERMINED BY THE <B>ALTERNATE</B> METHOD. See Rule 234H and T234-Alternate for details.')
    if alternate_method > 0 and params['system_type'] == '0':
        # AC voltage; alternate method eligible (vol_gnd exceeds 98 kv)
        table_footer.append('Per Rule 232-C.1c, the 5mA rms rule shall be checked.')

    # add footer index
    _table_footer_index(table_footer)

    data = []
    for i, item in enumerate(items):
        base = clearance_base[i]
        text_base = '{:.2f}'.format(base) if base > 0 else ''
        text_adder = '{:.2f}'.format(adder_elect) if base > 0 else ''
        text_adder_alt = '{:.2f}'.format(adder_altitude) if base > 0 else ''
        clearance_min = base + adder_elect + adder_altitude
        clearance_min = get_rounded(clearance_min, n_decimal=2, val_nearest=round_upto)
        clearance_design = clearance_min + margins[i]
        clearance_design = get_rounded(clearance_design, n_decimal=2, val_nearest=round_upto)
        text_min = '{:.2f}'.format(clearance_min) if base > 0 else ''
        text_design = '{:.2f}'.format(clearance_design) if base > 0 else ''
        data.append([item, text_base, text_adder, text_adder_alt, text_min, text_design])

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }
    return table_content


def get_table_233_alter(table_header, params, is_hori=False, alternate_method=1):
    '''
    :param table_header:
    :param params:
    :param is_hori:
    :param alternate_method: indicating if alternate calc is eligible
    :return:
    '''

    space4 = ''
    spacemany = ''
    for i in range(4):
        space4 += '&nbsp;'

    for i in range(85):
        spacemany += '&nbsp;'

    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    terrain_ele = get_float(params['terrain_ele'], 0.0)
    round_upto = get_float(params['round_upto'], 0.5)
    margin_v = get_float(params['margin_vert'], 0.0)
    margin_h = get_float(params['margin_hori'], 0.0)
    factor_switching = get_float(params['factor_switching'], value_default=1.0)

    title = ['T233V-Alternate. Vertical clearance between wires/conductors carried on different supporting structures '
             'by using the alternate method (C2-2017 NESC RULES 233C3b & assC3c)']

    items = [
        "<html>Communication lines</html>",
        "<html>Open supply lines that are effectively grounded or have ground fault relaying:<BR>"
        + space4 + "22 kV and below</html>",
        "<html>" + space4 + "46 kV</html>",
        "<html>" + space4 + "69 kV</html>",
        "<html>" + space4 + "115 kV</html>",
        "<html>" + space4 + "138 kV</html>",
        "<html>" + space4 + "161 kV</html>",
        "<html>" + space4 + "230 kV (<span style='color:red;'>Verify SOV for foreign line if practical</span>)</html>",
        "<html>" + space4 + "345 kV (<span style='color:red;'>Verify SOV for foreign line if practical</span>)</html>",
        "<html>" + space4 + "500 kV (<span style='color:red;'>Verify SOV for foreign line if practical</span>)</html>"
    ]

    clearance_base = [2.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    vol_cross = [0, 22 * math.sqrt(3), 46, 69*1.05, 115 * 1.05, 138 * 1.05, 161 * 1.05, 230 * 1.05, 345 * 1.05,
                   500 * 1.10]

    margin_ids = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    margins = []
    for margin_id in margin_ids:
        # margin_id: 0-Vertical, 1-Horizontal, 2-max
        if margin_id == 0:
            margins.append(margin_v)
        elif margin_id == 1:
            margins.append(margin_h)
        else:
            margins.append(max(margin_v, margin_h))

    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')


    # differentiate horizontal clearances
    if is_hori:
        items.pop(0)
        clearance_base.pop(0)
        vol_cross.pop(0)
        title = ['T233H-Alternate. Horizontal clearance between wires/conductors carried on different supporting structures '
            'by using the alternate method (C2-2017 NESC RULES 233-B2 & 235-B.3.a, 235-B.3.b)']

    table_footer = ['<html><B>Notes:</B></html>',
                    'Refer to NESC C2-2017 for more details including footnotes, exceptions, and variations associated '
                    'with the specific table.',
                    'Maximum voltage is used for 69kV to conform to industrial conservative practice although its line-to-ground voltage is less than 50kV'
                    ]
    if alternate_method == 0:
        table_footer.append("<font color='red'>Per Rule 232-D, the alternate method cannot be used with the specified voltage.</font>")

    elif alternate_method > 0:
        if is_hori:
            table_footer.append('Per NESC Rule 233B2, the alternate clearance shall be derived from the '
                                'computations required in Rules 235B3a and 235B3b. The clearance derived from '
                                'Rule 235B3a shall not less than the basic clearance given in Table 235-1 '
                                'computed for 169 kV ac. '
                                'The clearance values shown in this table may have been adjusted accordingly')
        else:
            table_footer.append("Per NESC Rule 233C3c, the alternate clearance shall be not less than the clearance "
                            "required by Rules 233C1 and 233C2 with the lower-voltage circuit at ground potential. "
                            "The clearance values shown in this table may have been adjusted accordingly.")


    # add footer index
    _table_footer_index(table_footer)

    data = []
    for i, item in enumerate(items):
        base = clearance_base[i]

        # calculate electrical adder
        if alternate_method == 0:
            adder_elect = -100
        else:
            v_cross = vol_cross[i]
            if is_hori:
                # equivalent voltage between two lines
                vol_line_equiv = (voltage_max + v_cross) / math.sqrt(3)
                if params['system_type'] == '1':
                    # primary circuit is DC
                    vol_line_equiv = voltage_max/math.sqrt(2) + v_cross/math.sqrt(3)
                adder_elect = get_adder_alternate_235(vol_line_equiv, factor_switching, a=1.15, b=1.03, k=1.4,
                                                      is_dc=False)
                # per Rule
                norm_235_hori = (29 + (169-50) * 0.4) / 12
                if norm_235_hori > adder_elect + base:
                    adder_elect = norm_235_hori - base
            else:
                # vertical
                adder_elect = get_adder_alternate_ft(voltage_max, factor_switching, a=1.15, b=1.03, c=1.2, k=1.4,
                                                   vol_l=v_cross, is_dc1=(params['system_type'] == '1'))

                # adder per normal calc with lower voltage at ground potential (Rule 233C3c)
                voltage_primary = max(voltage_max, vol_cross[i])
                adder_233_vert = get_adder_vol_ft(voltage_primary / math.sqrt(3), 22)
                norm_233_vert = adder_233_vert + 2 if vol_cross[i] > 0 else adder_233_vert + 5
                if norm_233_vert > adder_elect + clearance_base[i]:
                    adder_elect = norm_233_vert - clearance_base[i]

        adder_altitude = get_adder_altitude(terrain_ele, adder_elect, voltage_max, base_ele=1500)

        text_base = '{:.2f}'.format(base)
        text_adder = '{:.2f}'.format(adder_elect)
        text_adder_alt = '{:.2f}'.format(adder_altitude)
        clearance_min = base + adder_elect + adder_altitude
        clearance_min = get_rounded(clearance_min, n_decimal=2, val_nearest=round_upto)
        clearance_design = clearance_min + margins[i]
        clearance_design = get_rounded(clearance_design, n_decimal=2, val_nearest=round_upto)
        text_min = '{:.2f}'.format(clearance_min)
        text_design = '{:.2f}'.format(clearance_design)
        data.append([item, text_base, text_adder, text_adder_alt, text_min, text_design])

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }
    return table_content


def get_table_234_alter(table_header, params, alternate_method=1):
    '''
            :param alternate_method:
            :param table_header:
            :param params:
            :return:

            ENSURE length of first column not less than xx chars to get correct column width when displaying
        '''
    space4 = ''
    spacemany = ''
    for i in range(4):
        space4 += '&nbsp;'

    for i in range(85):
        spacemany += '&nbsp;'

    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    terrain_ele = get_float(params['terrain_ele'], 0.0)
    round_upto = get_float(params['round_upto'], 0.5)
    margin_v = get_float(params['margin_vert'], 0.0)
    margin_h = get_float(params['margin_hori'], 0.0)
    factor_switching = get_float(params['factor_switching'], value_default=1.0)

    title = ['T234-Alternate. Clearance from objects (e.g. buildings, bridges) by using the alternate method '
             '(C2-2017 NESC RULES 234H)']

    items = [
        "<html><b>1. Vertical</b></html>",
        "<html>a. Buildings</html>",
        "<html>b. Signs, chimneys, billboards, antennas, tanks etc.</html>",
        "<html>c. Superstructure of bridges</html>",
        "<html>d. Supporting structures of another line</html>",
        "<html>e. Dimension A of Figure 234-3</html>",
        "<html>f. Dimension B of Figure 234-3</html>",
        "<html><b>2. Horizontal</b></html>",
        "<html>a. Buildings</html>",
        "<html>b. Signs, chimneys, billboards, antennas,tanks etc.</html>",
        "<html>c. Superstructure of bridges</html>",
        "<html>d. Supporting structures of another line</html>",
        "<html>e. Dimension A of Figure 234-3</html>",
        "<html>f. Dimension B of Figure 234-3</html>"
    ]

    clearance_base = [0.0, 9.0, 9.0, 9.0, 6.0, 18.0, 14.0, 0.0, 3.0, 3.0, 3.0, 5.0, 0.0, 14.0]
    margin_ids = [0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1]
    margins = []
    for margin_id in margin_ids:
        # margin_id: 0-Vertical, 1-Horizontal, 2-max
        if margin_id == 0:
            margins.append(margin_v)
        elif margin_id == 1:
            margins.append(margin_h)
        else:
            margins.append(max(margin_v, margin_h))

    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')

    table_footer = ['<html><B>Notes:</B></html>',
                    "Refer to the NESC C2-2017 for clarifying notes, exceptions, and variations to the above mentioned areas and clearances.",
                    ]

    if alternate_method == 0:
        table_footer.append("<font color='red'>Per Rule 234-G, the alternate method cannot be used with the specified voltage.</font>")

    elif alternate_method == 1:
        table_footer.append(
            'Per Rule 234-G.1 Exception, alternate method may be used with the specified voltage. See Rule 234H and T234-Alternate for details.')
        table_footer.append(
            'Check the clearances in this table against Table 234-1, 234-2 & 3 computed for 169kV rms phase voltage. '
            'The clearance shall use the greater value (see Rule 234-H-4).')

    elif alternate_method == 2:
        table_footer.append(
            'Per Rule 234-G.1, the clearance SHALL BE DETERMINED BY THE <B>ALTERNATE</B> METHOD. See Rule 234H and T234-Alternate for details.')
        table_footer.append(
            'Check the clearances in this table against Table 234-1, 234-2 & 3 computed for 169kV rms phase voltage. '
            'The clearance shall use the greater value (see Rule 234-H-4).')
    if alternate_method > 0 and params['system_type'] == '0':
        table_footer.append('Per Rule 234-G.3, the 5mA rms rule shall be checked.')

    # add footer index
    _table_footer_index(table_footer)

    data = []
    for i, item in enumerate(items):
        base = clearance_base[i]
        adder_elect = get_adder_alternate_ft(voltage_max, factor_switching,
                                           a=1.15, b=1.03, c=1.2 if margin_ids[i] == 0 else 1.0, k=1.15,
                                           vol_l=0, is_dc1=(params['system_type'] == '1'))
        adder_altitude = get_adder_altitude(terrain_ele, adder_elect, voltage_max, base_ele=1500)

        text_base = '{:.2f}'.format(base) if base > 0 else ''
        text_adder = '{:.2f}'.format(adder_elect) if base > 0 else ''
        text_adder_alt = '{:.2f}'.format(adder_altitude) if base > 0 else ''
        clearance_min = base + adder_elect + adder_altitude
        clearance_min = get_rounded(clearance_min, n_decimal=2, val_nearest=round_upto)
        clearance_design = clearance_min + margins[i]
        clearance_design = get_rounded(clearance_design, n_decimal=2, val_nearest=round_upto)
        text_min = '{:.2f}'.format(clearance_min) if base > 0 else ''
        text_design = '{:.2f}'.format(clearance_design) if base > 0 else ''
        data.append([item, text_base, text_adder, text_adder_alt, text_min, text_design])

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }
    return table_content


def get_table_GO95_1(params):
    """Deprecated. See separate module for GO-95"""
    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    round_upto = get_float(params['round_upto'], 0.5)
    margin_v = get_float(params['margin_vert'], 0.0)
    margin_h = get_float(params['margin_hori'], 0.0)
    margin = max(margin_v, margin_v)

    title = ['Table 1. Vertical clearance above railroads, thoroughfares, ground or water surfaces (GO-95 Rule 37)']
    table_header = ['Nature of Surface',
              '<html>GO-95<BR>Base (ft)</html>',
              '<html>Voltage<BR>Adder (ft)</html>',
              '<html>Clearance<BR>Min(ft)</html>',
              '<html>Clearance<BR>Design (ft)</html>',
              '<html>Reduced<BR>Design (ft)</html>']
    # vol < 0.75
    base750 = [25.0, 20.0, 20.0, 19.0, 12.0,  8.0, 8.0, 3.0, 1.3, 1.0, 3.0, 15.0, 18.0, 26.0, 32.0, 38.0, 0, 0]

    # vol < 22.5
    base22 =  [28.0, 25.0, 25.0, 25.0, 17.0, 12.0, 8.0, 6.0, 1.5, 1.0, 6.0, 17.0, 20.0, 28.0, 34.0, 40.0, 1.5, 4.0]

    # vol < 300
    base300 = [34.0, 30.0, 30.0, 30.0, 25.0, 12.0, 8.0, 6.0, 1.5, 1.0, 10.0, 25.0, 27.0, 35.0, 41.0, 47.0, 1.5, 4.0]

    # vol >= 300
    base550 = [34.0, 34.0, 30.0, 30.0, 25.0, 20.0, 20.0, 15.0, 0.0, 9.58, 10.0, 25.0, 27.0, 35.0, 41.0, 47.0, 9.58, 10.0]

    # adder_id: 0 - 0.025ft/kv; 1 - 0.04ft/kV; 2 - 0 ft/kV
    adder_id = [0, 0, 0, 0, 0, 1, 2, 2, 2, 2, 1, 0, 0, 0, 0, 0, 2, 2]

    factor_reduce = [0.95, 0.9, 0.9, 0.9, 0.9, 0.9, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0]

    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')
    if voltage_max <= 0.75:
        base95 = base750
    elif voltage_max <= 22.5:
        base95 = base22
    elif voltage_max <= 300:
        base95 = base300
    else:
        base95 = base550

    items = [
        '<html><B>1</B>. Crossing above tracks of railroads which transport or propose to transport freight cars (maximum height 15 feet, 6 inches) where not operated by overhead contact wires.</html>',
        '<html><B>2</B>. Crossing or paralleling above tracks of railroads operated by overhead trolleys. (Note: Assumes Trolley Pole Throw of 26 feet. If Trolley Pole Throw is greater than 26 feet, clearance should be increased be the difference.  See Rule 54.4-B2 for more details)</html>',
        '<html><B>3</B>. Crossing or along thoroughfares in urban districts or crossing thoroughfares in rural districts.',
        '<html><B>4</B>. Above ground along thoroughfares in rural districts or across other areas capable of being traversed by vehicles or agricultural equipment.</html>',
        '<html><B>5</B>. Above ground in areas accessible to pedestrians only.',
        '<html><B>6</B>. Vertical clearance above walkable surfaces on buildings, (except generating plants or substations) bridges or other structures which do not ordinarily support conductors, whether attached or unattached.</html>',
        '<html><B>6a</B>. Vertical clearance above non–walkable surfaces on buildings, (except generating plants or substations) bridges or other structures, which do not ordinarily support conductors, whether attached or unattached.</html>',
        '<html><B>7</B>. Horizontal clearance of conductor at rest from buildings (except generating plants and substations), bridges or other structures (upon which men may work) where such conductor is not attached thereto.</html>',
        '<html><B>8</B>. Distance of conductor from center line of pole, whether attached or unattached. (Note: Conductors passing and unattached: The centerline clearance between poles and conductors which pass unattached shall be not less than 1-1/2 times the clearance specifed in table 1, case 8 except where the provisions of table 1, case 10, columns for suppy conductors and cables can be applied. See Rule 54.4-D3</html>',
        '<html><B>9</B>. Distance of conductor from surface of pole, crossarm or other overhead line structure upon which it is supported, providing it complies with case 8 above. (Note: For clearance of conductors deadended in vertical and horizontal configurations on poles see Rule 54.4-C4 and 54.4-D8, respectively.)</html>',
        '<html><B>10</B>. Radial centerline clearance of conductor or cable (unattached) from non–climbable street lighting or traffic signal poles or standards, including mastarms, brackets and lighting fixtures, and from antennas that are not part of the overhead line system.</html>',
        '<html><B>11</B>. Water areas not suitable for sailboating.</html>',
        '<html><B>12</B>. Water areas suitable for sailboating, surface area of:<BR>&nbsp;&nbsp;&nbsp;&nbsp; (A) Less than 20 acres</html>',
        '<html>&nbsp;&nbsp;&nbsp;&nbsp; (B) 20 to 200 acres</html>',
        '<html>&nbsp;&nbsp;&nbsp;&nbsp; (C) Over 200 to 2,000 acres</html>',
        '<html>&nbsp;&nbsp;&nbsp;&nbsp; (D) Over 2,000 acres</html>',
        '<html><B>13</B>. Radial clearance of bare line conductors from tree branches or foliage.</html>',
        '<html><B>14</B>. Radial clearance of bare line conductors from vegetation in Extreme and Very High Fire Threat Zones in Southern California.</html>'
    ]

    table_footer = ['<html><B>Notes:</B></html>',
                    '1. Check GO-95 clearance requirements per Rule 37 <a href="https://www.cpuc.ca.gov/gos/GO95/go_95_table_1.html">HERE (Table 1)</a>',
                    '2. Proposed clearances for voltages above 550kV shall be submitted to CPUC for approval prior to construction.',
                    '3. ',
                    '4. ',
                    '5. ']
    data = []
    adder_factor = [0.025, 0.04, 0]
    for i, item in enumerate(items):
        base = base95[i]
        adder95 = max(voltage_max - 300, 0) * adder_factor[adder_id[i]]
        min_clearance = base + adder95
        design = min_clearance + margin
        design_reduced = min_clearance * factor_reduce[i] + margin

        min_clearance = get_rounded(min_clearance, n_decimal=2, val_nearest=round_upto)
        design = get_rounded(design, n_decimal=2, val_nearest=round_upto)
        design_reduced = get_rounded(design_reduced, n_decimal=2, val_nearest=round_upto)

        text_base = '{:.2f}'.format(base) if base > 0 else 'N/A'
        text_adder = '{:.2f}'.format(adder95) if base > 0 else 'N/A'
        text_min_clearance = '{:.2f}'.format(min_clearance) if base > 0 else 'N/A'
        text_design = '{:.2f}'.format(design) if base > 0 else 'N/A'
        text_design_reduced = '{:.2f}'.format(design_reduced) if base > 0 else 'N/A'
        data.append([item, text_base, text_adder, text_min_clearance, text_design, text_design_reduced])

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }

    return table_content


def get_table_GO95_2(params):
    voltage_sys = get_float(params['system_voltage'], value_default=0.0)
    factor_vol = get_float(params['factor_voltage'], value_default=1.0)
    round_upto = get_float(params['round_upto'], 0.5)
    margin_v = get_float(params['margin_vert'], 0.0)
    margin_h = get_float(params['margin_hori'], 0.0)
    margin = max(margin_v, margin_h)

    title = ['Table 1. Vertical clearance above railroads, thoroughfares, ground or water surfaces (GO-95 Rule 37)']
    table_header = ['Nature of Surface',
                    '<html>GO-95<BR>Base (ft)</html>',
                    '<html>Voltage<BR>Adder (ft)</html>',
                    '<html>Clearance<BR>Min(ft)</html>',
                    '<html>Clearance<BR>Design (ft)</html>',
                    '<html>Reduced<BR>Design (ft)</html>']
    # vol < 0.75
    base_75 = [24, 48, 48, 24, 48, 48, 96, 0, 48, 24, 48, 48, 72, 72, 0, 12, 0, 11.5, 0, 11.5, 15, 3, 0, 11.5, 3, 0, 0, 0, 48]
    # vol < 7.5K
    base7_5 = [36, 48, 48, 48, 48, 72, 96, 0, 48, 48, 48, 48, 48, 60, 0, 18, 0, 11.5, 0, 11.5, 18, 6, 0, 11.5, 6, 0, 24, 0, 72]
    # vol < 20K
    base20 = [36, 72, 72, 48, 72, 72, 96, 0, 72, 48, 48, 48, 48, 60, 0, 18, 0, 17.5, 0, 17.5, 18, 6, 0, 17.5, 9, 0, 24, 0, 72]
    # vol < 35K
    base35 = [72, 96, 96, 96, 96, 96, 96, 0, 72, 72, 48, 48, 48, 60, 0, 24, 0, 24, 0, 24, 18, 12, 0, 24, 12, 0, 24, 0, 72]
    # vol , 75K
    base75 = [72, 96, 96, 96, 96, 96, 96, 0, 72, 72, 48, 48, 48, 60, 0, 48, 0, 48, 0, 48, 18, 24, 0, 36, 18, 0, 36 or 48, 0, 120]
    # vol < 150K
    base150 = [78, 96, 96, 96, 96, 96, 96, 0, 78, 78, 60, 60, 60, 60, 0, 60, 0, 60, 0, 60, 24, 60, 0, 36, 24, 0, 48, 0, 0]
    # vol < 300
    base300 = [78, 96, 96, 96, 96, 96, 96, 0, 87, 87, 90, 90, 90, 90, 0, 90, 0, 90, 0, 90, 36, 90, 0, 78, 48, 0, 48, 0, 0]
    # vol >= 300
    base550 = [138, 156, 156, 156, 156, 156, 156, 0, 147, 147, 150, 150, 150, 150, 0, 150, 0, 150, 0, 150, 120, 150, 0, 138, 86, 0, 48, 0, 0]

    # adder_id: 0 - 0.025ft/kv; 1 - 0.04ft/kV; 2 - 0 ft/kV
    # adder_id = [0, 0, 0, 0, 0, 1, 2, 2, 2, 2, 1, 0, 0, 0, 0, 0, 2, 2]

    # factor_reduce = [0.95, 0.9, 0.9, 0.9, 0.9, 0.9, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0]

    voltage_max = _get_voltage_max(voltage_sys, factor_vol, params['system_type'] == '1')
    if voltage_max <= 0.75:
        base95 = base_75
    if voltage_max <= 7.5:
        base95 = base7_5
    elif voltage_max <= 20:
        base95 = base20
    elif voltage_max <= 35:
        base95 = base35
    elif voltage_max <= 75:
        base95 = base75
    elif voltage_max <= 150:
        base95 = base150
    elif voltage_max <= 300:
        base95 = base300
    else:
        base95 = base550

    items = [
        '<html><B>Clearance between wires, cables and conductors not supported on the same poles, vertically at crossings in spans and radially where colinear or approaching crossings</B></html>',
        '<html>1. Span wires, guys and messengers</html>',
        '<html>2. Trolley contact conductors, 0 - 750 volts</html>',
        '<html>3. Communication conductors</html>',
        '<html>4. Supply conductors, 500 kV</html>',
        '<html>5. Supply conductors, 345 kV</html>',
        '<html>6. Supply conductors, 300 kV</html>',
        '<html>7. Supply conductors, 230 kV</html>',
        '<html>7a. Supply conductors, 138 kV</html>',
        '<html>7b. Supply conductors, 115 kV</html>',
        '<html>7c. Supply conductors, 69 kV</html>',
        '<html><B>Vertical separation between conductors and/or cables, on separate crossarms or other supports at different levels (excepting on related line and buck arms) on the same pole and in adjoining midspans</B></html>',
        '<html>8. Communication Conductors and Service Drops</html>',
        '<html>9. Supply Conductors Service Drops and Trolley Feeders, 0 - 750 Volts</html>',
        '<html>10. Supply conductors, 500 kV</html>',
        '<html>11. Supply conductors, 345 kV</html>',
        '<html>12. Supply conductors, 300 kV</html>',
        '<html>13. Supply conductors, 230 kV</html>',
        '<html>13a. Supply conductors, 138 kV</html>',
        '<html>13b. Supply conductors, 115 kV</html>',
        '<html>13c. Supply conductors, 69 kV</html>',
        '<html><B>Vertical clearance between conductors on related line arms and buck arms</B></html>',
        '<html>14. Line arms above or below related buck arms</html>',
        '<html><B>Horizontal separation of conductors on same crossarm</B></html>',
        '<html>15. Pin spacing of longitudinal conductors vertical conductors and service drops</html>',
        '<html><B>Radial separation of conductors on same crossarm, pole or structure—incidental pole wiring</html>',
        '<html>16. Conductors, taps or lead wires of different circuits</html>',
        '<html>16a. Uncovered, grounded, non-dielectric fiber optic cables on metallic structures, in transition</html>',
        '<html>17. Conductors, taps or lead wires of the same circuit</html>',
        '<html><B>Radial separation between guys and conductors</B></html>',
        '<html>18. Guys passing conductors supported on other poles, or guys approximately parallel to conductors supported on the same poles</html>',
        '<html>19. Guys and span wires passing conductors supported on the same poles</html>',
        '<html><B>Vertical and horizontal insulators clearances between conductors</B></html>',
        '<html>20. Vertical clearance between conductors of the same circuit on horizontal insulators</html>',
        '<html>Vertical clearance above supply and/or communication lines</html>',
        '<html>21. Antennas and associated elements on the same support structure.</html>'
    ]

    table_footer = ['<html><B>Notes:</B></html>',
                    '1. Check GO-95 clearance requirements per Rule 38 <a href="https://www.cpuc.ca.gov/gos/GO95/go_95_table_2.html">HERE (Table 2)</a>',
                    '2. Unit of values in Talbe 2 is inch. Minimum clearance requirement is based on a temperature of 60 deg F, and no wind.',
                    '3. ',
                    '4. ',
                    '5. ']
    data = []
    adder_factor = [0.025, 0.04, 0]
    for i, item in enumerate(items):
        base = base95[i]
        adder95 = max(voltage_max - 300, 0) * adder_factor[adder_id[i]]
        min_clearance = base + adder95
        design = min_clearance + margin
        design_reduced = design * factor_reduce[i]

        min_clearance = get_rounded(min_clearance, n_decimal=2, val_nearest=round_upto)
        design = get_rounded(design, n_decimal=2, val_nearest=round_upto)
        design_reduced = get_rounded(design_reduced, n_decimal=2, val_nearest=round_upto)

        text_base = '{:.2f}'.format(base) if base > 0 else 'N/A'
        text_adder = '{:.2f}'.format(adder95) if base > 0 else 'N/A'
        text_min_clearance = '{:.2f}'.format(min_clearance) if base > 0 else 'N/A'
        text_design = '{:.2f}'.format(design) if base > 0 else 'N/A'
        text_design_reduced = '{:.2f}'.format(design_reduced) if base > 0 else 'N/A'
        data.append([item, text_base, text_adder, text_min_clearance, text_design, text_design_reduced])

    table_content = {
        'title': title,
        'header': table_header,
        'footer': table_footer,
        'data': data
    }

    return table_content


def get_table_GO95_2A(params):
    return ''

def get_table_GO95_swim(params):
    return ''


''' ---------------------------------------
HELPER FUNCTIONS [COMMON for PDF]
--------------------------------------- '''
def get_rounded(value, n_decimal, val_nearest):
    # round up to the nearest value
    # n_decimal: number of decimals
    # ceil to round up, floor to round down
    n_count = math.ceil(value / val_nearest)
    return round(n_count * val_nearest, n_decimal)


def get_float(text, value_default):
    try:
        float_value = float(text)
    except:
        float_value = value_default
    return float_value


def get_adder_vol_ft(vol_ground, vol_base, kv_increment=0.4):
    # adder based on ground voltage; return ft
    voltage = max(vol_ground - vol_base, 0)
    adder = voltage * kv_increment
    return adder / 12


def get_adder_vol_in(vol_gnd, vol_base, kv_increment=0.4):
    # adder based on 8.7 kV, 0.4"/kV, return inch
    voltage = max(vol_gnd - vol_base, 0)
    adder = voltage * kv_increment
    return adder


def get_adder_altitude(terrain_ele, adder_vol, voltage_max, base_ele=3300, is_dc=False):
    if is_dc:
        voltage_gnd = voltage_max
    else: voltage_gnd = voltage_max / math.sqrt(3)

    if voltage_gnd <= 50/math.sqrt(3): # conform to conservative industrial practice
        return 0.0
    else:
        ele_additional = max(terrain_ele - base_ele, 0)
        return (ele_additional / 1000) * 0.03 * adder_vol


def get_adder_alternate_ft(vol_h, pu, a, b, c, k, is_dc1=False, is_dc2=False, vol_l=0):
    '''
    Associated Rule 233 for line crossing
    :param vol_h: maximum primary voltage, kV
    :param vol_l: maximum secondary voltage if applicable, e.g. crossing; assumed to be AC line-to-line
    :param pu: switch factor
    :param a, b, c, k: constants
    is_dc: primary voltage type, either 3-phase AC (line-to-line), or DC (pole-to-ground)
    :return: electrical adder D
    '''

    # get crest voltage
    if not is_dc1:
        vol_h = vol_h / math.sqrt(3) * math.sqrt(2)

    if not is_dc2:
        vol_l = vol_l / math.sqrt(3) * math.sqrt(2)

    root = (vol_h * pu + vol_l) * a / (500 * k)
    d = 3.28 * pow(root, 1.667) * b * c
    #print(vol_h, vol_l, pu, a, b, c, k)

    return d


def get_adder_alternate_235(vol_line_line, pu, a, b, k, c=1.0, is_dc=False):
    '''
    :param vol_line_line: RMS between two lines for AC; Crest to ground for DC
    :param pu:
    :param a:
    :param b:
    :param k:
    :param c:
    :param is_dc:
    :return:
    '''
    if is_dc:
        vol_line_line *= 2
    vol_line_line *= math.sqrt(2)
    root = vol_line_line * pu * a / (500 * k)
    d = 3.28 * pow(root, 1.667) * b * c

    return d

''' ---------------------------------------
HELPER FUNCTIONS
--------------------------------------- '''


def _verify_json_spreadsheet(params_dict):
    list_err = []
    file_type = params_dict.get('file_type')
    file_cont = params_dict.get('file_content')
    if file_type is None:
        list_err.append('File does not have the correct data type. Stopped processing for security reason')
    elif file_type != 'json_ecs_created':
        list_err.append('File does not have the correct data type. Stopped processing for security reason')

    if file_cont is None:
        list_err.append('File data content cannot be verified. Stopped processing for security reason')
    elif file_cont != 'clearance_table_list':
        list_err.append('File data content cannot be verified. Stopped processing for security reason')

    try:
        if not params_dict.get('system_voltage').replace('.', '', 1).isdigit():
            list_err.append('Invalid format identified for System Voltage.')
        if not params_dict.get('system_type').isdigit():
            list_err.append('Invalid System Type')
        if not params_dict.get('factor_voltage').replace('.', '', 1).isdigit():
            list_err.append('Invalid format identified for voltage multiplier.')
        if not params_dict.get('terrain_ele').replace('.', '', 1).isdigit():
            list_err.append('Invalid format identified for Terrain Elevation.')
        if not params_dict.get('margin_vert').replace('.', '', 1).isdigit():
            list_err.append('Invalid format identified for Vertical Clearance Margin.')
        if not params_dict.get('margin_hori').replace('.', '', 1).isdigit():
            list_err.append('Invalid format identified for Horizontal Clearance Margin.')
        if not params_dict.get('round_upto').replace('.', '', 1).isdigit():
            list_err.append('Invalid format identified for Round-up-to definition.')
        if params_dict.get('round_upto').replace('.', '', 1).isdigit():
            round_upto = float(params_dict.get('round_upto'))
            if abs(round_upto - 0) < 0.00001:
                list_err.append('Invalid value identified for Round-up-to definition (has to be non-zero).')
        if not params_dict.get('mot_temperature').replace('.', '', 1).isdigit():
            list_err.append('Invalid format identified for MOT.')
        if not params_dict.get('factor_switching').replace('.', '', 1).isdigit():
            list_err.append('Invalid format identified for Switching Factor.')
    except:
        list_err.append('Missing one or more data entries for input.')

    return list_err


def _convert_table_content_to_ws(table_content, table_desc):
    '''
    INPUT: table_content - returned from API for APP output
    RETURN: worksheet format ready for spreadsheet output

    '''
    color_bkg_header = GlobalVar.color_bkg_header
    color_bkg_data1 = GlobalVar.color_bkg_data_1
    color_bkg_data2 = GlobalVar.color_bkg_data_2
    color_bkg_special = GlobalVar.color_bkg_special

    list_data = []

    title = '\n'.join(table_content.get('title')).replace('<html>', '').replace('</html>', '')
    title = title.replace('<BR>', '\n')
    
    footers = table_content.get('footer')
    headers = table_content.get('header')
    headers = [h.replace('<BR>', '\n') for h in headers]
    footers = [f.replace('<BR>', '\n') for f in footers]
    list_idx_red_footer = [footers.index(footer) for footer in footers if 'color:red' in footer]
    regex = re.compile('<[/a-z\s=:;"\']+>')
    for i, h in enumerate(headers):
        brackets = regex.findall(h)
        if brackets:
            for b in brackets:
                h = h.replace(b, '')
        headers[i] = h

    footers[0] = 'Notes:'  # remove <B></B>
    # change href link to fit spreadsheet cell
    list_url_display = []
    for i, footer in enumerate(footers):
        # find all urls for each footer
        urls = common_utils.find_all_urls_in_string(footer)
        if urls:
            for url in urls:
                # isolate display text
                footer = footer.replace(' target="_blank"', '')
                idx_display = footer.find(url) + len(url)
                idx_anchor_end = footer.find('</a>')
                href_display = footer[idx_display + 2: idx_anchor_end]
                list_url_display.append({'url': url, 'href_display': href_display})

                # remove HREF related text from string
                #footer = footer.replace(url, '')
            #footer = footer.replace('<a href="">', '').replace('</a>', '')

        if 'Table' in table_desc or 'Swimming' in table_desc:
            # not applicable for GO-95 tables. Skip
            continue

        brackets = regex.findall(footer)
        if brackets:
            for bra in brackets:
                footer = footer.replace(bra, '')
            footers[i] = footer.replace('Red font', 'Red background')
            list_idx_red_footer.append(i)

    # add references for URLS
    if list_url_display:
        footers.append('')
        footers.append('Reference Links:')
        for url in list_url_display:
            footers.append('=HYPERLINK("{}", "{}")'.format(url.get('url'), url.get('href_display')) + '\n')

    data = table_content.get('data')

    list_data.append([title])
    list_data.append(headers)
    list_idx_red_row = []
    for i, row in enumerate(data):
        # reformat first column by removing html tags
        row[0] = row[0].replace('<BR>', '\n').replace('<b>', '').replace('</b>', '').replace('&nbsp;', ' ')
        row[0] = row[0].replace('<B>', '').replace('</B>', '').replace('<i>', '').replace('</i>', '')

        brackets = regex.findall(row[0])
        if brackets:
            for bra in brackets:
                row[0] = row[0].replace(bra, '')

        # heavy background color if data is shown for reference only and alternate method is required
        if not ('Table' in table_desc or 'Swimming' in table_desc):
            # except for GO-95 tables
            for j, cell in enumerate(row):
                # remove anything in <> and mark row number
                bracket_sharp = regex.findall(cell)
                if bracket_sharp:
                    for bracket in bracket_sharp:
                        cell = cell.replace(bracket, '')

                    row[j] = cell
                    list_idx_red_row.append(i)

        list_data.append(row)

    # cell styles
    n_row = len(list_data)
    n_col = len(data[0])
    range_merge = [(1, 1, 1, n_col, 'center')]
    range_font_bold = [(1, 1, 1, n_col)]
    range_font_size = [(1, 1, 1, n_col, 16)]
    row_height = [(1, 60)]
    col_width = [(1, 80)]
    range_border = [(1, 1, n_row, n_col)]
    for i in range(1, n_col):
        col_width.append((i + 1, 20))
    range_color = [(2, 1, 2, n_col, color_bkg_header)]
    for i in range(2, n_row):
        if i % 2 == 0:
            range_color.append((i + 1, 1, i + 1, n_col, color_bkg_data1))
        else:
            range_color.append((i + 1, 1, i + 1, n_col, color_bkg_data2))
    
    for idx in list_idx_red_row:
        # skip two rows, add 1 due to idx base
        range_color.append((idx+3, 1, idx+3, n_col, color_bkg_special))
    for idx in list_idx_red_footer:
        idx_row = n_row + 1 + idx + 1
        range_color.append((idx_row, 1, idx_row, n_col, color_bkg_special))

    cell_range_style = {
        'range_merge': range_merge,
        'range_color': range_color,
        'range_font_bold': range_font_bold,
        'range_font_size': range_font_size,
        'range_border': range_border,
        'column_width': col_width,
        'row_height': row_height
    }

    ws = {
        'ws_name': table_desc if len(table_desc) < 30 else table_desc[:29],
        'ws_content': list_data,
        'footer': footers,
        'cell_range_style': cell_range_style
    }
    return ws



def _get_voltage_max(system_voltage, factor_vol, is_dc=False):
    voltage_gnd = system_voltage / math.sqrt(2) if is_dc else system_voltage / math.sqrt(3)

    if voltage_gnd >= 50:
        return system_voltage * factor_vol

    elif system_voltage>68 and (not is_dc):
        # vol_line = 69kC: will use maximum operating to conform to utilities' conservative practice
        return system_voltage * factor_vol

    else:
        return system_voltage


def _table_footer_index(table_footer):
    # adder ordered index for footnotes
    for i, footer in enumerate(table_footer):
        if i > 0:
            table_footer[i] = '{:2d}. '.format(i) + footer


def _retrieve_form_data_input(form_data):
    try:
        params = {'system_type': form_data['system_type'],
                  'system_voltage': form_data['system_voltage'],
                  'factor_voltage': form_data['factor_voltage'],
                  'terrain_ele': form_data['terrain_ele'],
                  'margin_vert': form_data['margin_vert'],
                  'margin_hori': form_data['margin_hori'],
                  'round_upto': form_data['round_upto'],
                  'mot_temperature': form_data['mot_temperature'],
                  'factor_switching': form_data['factor_switching'],
                  'item_name': form_data['item_name'],
                  'form_id': form_data['form_id']
                  }

        #self.params = params
        print('::: params updated successfully with menue_item={}.'.format(form_data['item_name']))

    except:
        # params = self.params_init
        params = None
        print('::: params cannot be updated due to lack of form data.')

    return params


def _retrieve_form_data_proj(form_data):
    try:
        params = {'proj_name': form_data['proj_name'],
                  'proj_numb': form_data['proj_numb'],
                  'proj_subj': form_data['proj_subj'],
                  'proj_engineer': form_data['proj_engineer'],
                  'proj_checker': form_data['proj_checker'],
                  'date_design': form_data['date_design'],
                  'date_check': form_data['date_check'],
                  'proj_note': form_data['proj_note'],
                  'submit': form_data['submit'],
                  'form_id': form_data['form_id']}
        return params

    except:
        return None


def _retrieve_form_data_nesc235(form_data):

    try:
        params = {
            'same_circuit': form_data['same_circuit'],
            'same_utility': form_data['same_utility'],
            'circuit1': form_data['circuit1'],
            'vol_multipler1': form_data['vol_multipler1'],
            'type_system1': form_data['type_system1'],
            'circuit2': form_data['circuit2'],
            'vol_multipler2': form_data['vol_multipler2'],
            'type_system2': form_data['type_system2'],
            'factor_experience': form_data['factor_experience'],
            'angle_swing': form_data['angle_swing'],
            'value_sag': form_data['value_sag'],
            'len_insulator': form_data['len_insulator'],
            'elevation_ter': form_data['elevation_ter'],
            'factor_pu': form_data['factor_pu'],
            'rst_hori_per_volt': form_data['rst_hori_per_volt'],
            'rst_vert_at_support': form_data['rst_vert_at_support'],
            'rst_hori_per_sag': form_data['rst_hori_per_sag'],
            'rst_vert_in_span': form_data['rst_vert_in_span'],
            'rst_hori_alternate': form_data['rst_hori_alternate'],
            'rst_vert_alternate': form_data['rst_vert_alternate'],
            'notes_area': form_data['notes_area']
        }
        return params

    except:
        return None


def _retrieve_form_data_csa59(params, form_data):
    try:
        prams = {}
        return params

    except:
        return None


def _get_clearance_per235(params_rule235):
    if params_rule235 is None:
        print('::: unexpected error when processing Rule 235...')
        return
    if not _validate_input_rule235(params=params_rule235):
        print('::: Error in form field - character encountered when digit number expected.')
        return

    factor_exps = [0.67, 1.15, 1.20, 1.25]
    is_same_circuit = (params_rule235['same_circuit'] == '0')
    is_same_utility = (params_rule235['same_utility'] == '0')
    is_dc1 = params_rule235['type_system1'] == '1'
    is_dc2 = params_rule235['type_system2'] == '1'
    is_com = params_rule235['type_system2'] == '2'

    vol_primary = float(params_rule235['circuit1'])
    vol_secondy = float(params_rule235['circuit2'])
    vol_factor1 = float(params_rule235['vol_multipler1'])
    vol_factor2 = float(params_rule235['vol_multipler2'])
    factor_experience = factor_exps[int(params_rule235['factor_experience'])]
    angle_swing = float(params_rule235['angle_swing'])
    value_sag = float(params_rule235['value_sag'])
    len_insulator = float(params_rule235['len_insulator'])
    elevation_ter = float(params_rule235['elevation_ter'])
    factor_pu = float(params_rule235['factor_pu'])

    # get maximum voltage value; use max for 69kV as well (add in notes)
    vol_primary = vol_primary * vol_factor1 if vol_primary >= 50 else vol_primary
    vol_secondy = vol_secondy * vol_factor2 if vol_secondy >= 50 else vol_secondy

    # get voltage between lines, rms
    if is_same_circuit:
        vol_equiv = vol_primary * 2 / math.sqrt(2) if is_dc1 else vol_primary

    else:
        vol_primary_gnd = vol_primary / math.sqrt(2) if is_dc1 else vol_primary / math.sqrt(3)
        vol_secondy_gnd = vol_secondy / math.sqrt(2) if is_dc2 else vol_secondy / math.sqrt(3)
        vol_equiv = vol_primary_gnd + vol_secondy_gnd


    # **************** Clearance Horizontal per Voltage ***********************************
    # Horizontal clearance (a) 235-B1a (Table 235-1) per voltage
    if vol_equiv < 50:
        # no altitude adjustment is needed
        clear_hori_vol = 12 + get_adder_vol_in(vol_gnd=vol_equiv, vol_base=8.7, kv_increment=0.4)
    else:
        clear_hori_base = 29
        clear_hori_adder = get_adder_vol_in(vol_gnd=vol_equiv, vol_base=50, kv_increment=0.4)

        # altitude adder per Rule 235-B1b(1) and 235-B1b(2) for horizontal clearance per sag
        # altitude adder for horizontal clearance per voltage is not specified but applied in this calc
        clear_hori_adder_ele = get_adder_altitude(terrain_ele=elevation_ter, adder_vol=clear_hori_adder,
                                                   voltage_max=vol_equiv, base_ele=3300)
        clear_hori_vol = clear_hori_base + clear_hori_adder + clear_hori_adder_ele

    # round up to whole number for inch, then convert to feet
    clear_hori_vol = math.ceil(clear_hori_vol) / 12.0
    clear_hori_vol = math.ceil(clear_hori_vol * 100) / 100

    # **************** Clearance Horizontal per Sag ***********************************
    # (b) 235-B1b (clause) per sag and swinging, combined with Experience factor defined in RUS
    # calc'ed according to eq 6-1 in RUS 200
    pie = math.atan(1.0) * 4
    clear_hori_sag = 0.3/12 * vol_equiv + factor_experience * math.sqrt(value_sag) + \
                     len_insulator * math.sin(angle_swing * pie /180)
    adder_altitude = get_adder_altitude(terrain_ele=elevation_ter, adder_vol=clear_hori_sag,
                                         voltage_max=vol_equiv, base_ele=3300)

    clear_hori_sag += adder_altitude
    clear_hori_sag = math.ceil(clear_hori_sag*100) / 100

    # **************** Clearance Vertical at Support ***********************************
    # clause 235C; table 235-5
    clear_vert_base = 16 if is_same_utility else 40
    if vol_equiv <= 50:
        clear_vert_at_support = clear_vert_base + get_adder_vol_in(vol_gnd=vol_equiv, vol_base=8.7, kv_increment=0.4)
    else:
        adder_per_volt50 = get_adder_vol_in(vol_gnd=50, vol_base=8.7, kv_increment=0.4)
        adder_per_volt = get_adder_vol_in(vol_gnd=vol_equiv, vol_base=50, kv_increment=0.4)

        adder_altitude = get_adder_altitude(terrain_ele=elevation_ter, adder_vol=adder_per_volt,
                                             voltage_max=vol_equiv, base_ele=3300, is_dc=False)
        clear_vert_at_support = clear_vert_base + adder_per_volt50 + adder_per_volt + adder_altitude

    clear_vert_at_support = math.ceil(clear_vert_at_support) / 12
    clear_vert_at_support = math.ceil(clear_vert_at_support * 100) / 100

    # **************** Clearance Vertical in Span ***********************************
    if vol_equiv <= 50:
        clear_vert_at_support = clear_vert_base + get_adder_vol_in(vol_gnd=vol_equiv, vol_base=8.7, kv_increment=0.4)
        clear_vert_in_span = clear_vert_at_support * 0.75

    else:
        adder_per_volt50 = get_adder_vol_in(vol_gnd=50, vol_base=8.7, kv_increment=0.4)
        adder_per_volt = get_adder_vol_in(vol_gnd=vol_equiv, vol_base=50, kv_increment=0.4)
        adder_altitude = get_adder_altitude(terrain_ele=elevation_ter, adder_vol=adder_per_volt,
                                             voltage_max=vol_equiv, base_ele=3300, is_dc=False)
        clear_vert_in_span = (clear_vert_base + adder_per_volt50) * 0.75 + adder_per_volt + adder_altitude
        # round up to inches per example in 235C2a(1)
        clear_vert_in_span = math.ceil(clear_vert_in_span) / 12
        clear_vert_in_span = math.ceil(clear_vert_in_span * 100) / 100

    # **************** Clearance Horizontal Alternate Method (235-B3) ***********************************
    # ONLY applicable to different circuit (Rule 235-B3), but table 235-4 assumes same circuit;
    # THEREFORE calc for same cct is provided as reference

    # vol_equiv is voltage between the two lines in RMS; will be converted to crest in subroutine:
    clear_hori_alter = get_adder_alternate_235(vol_line_line=vol_equiv, pu=factor_pu, a=1.15, b=1.03, k=1.4)

    # lower limit per 235-B3b - not less than table 235-1 computed for 169 kV ac
    clear_hori_alter_limit = (29 + (169-50) * 0.4) / 12
    clear_hori_alter = max(clear_hori_alter, clear_hori_alter_limit)

    # altitude adjust
    clear_hori_altitude = get_adder_altitude(terrain_ele=elevation_ter, adder_vol=clear_hori_alter,
                                              voltage_max=vol_equiv, base_ele=1500)

    clear_hori_alter += clear_hori_altitude
    clear_hori_alter = math.ceil(clear_hori_alter * 100) / 100

    # **************** Clearance Vertical Alternate Method (235-C3) equation in 233C3 ********************************
    # ONLY applicable to different circuit (Rule 235-C3)
    # calc for same cct is provided as reference

    if is_same_circuit:
        # No need to convert to ground voltage; voltage between lines will beused
        clear_vert_alter = get_adder_alternate_235(vol_line_line=vol_primary, pu=factor_pu,
                                                   a=1.15, b=1.03, c=1.2, k=1.4, is_dc=is_dc1)

    else:
        clear_vert_alter = get_adder_alternate_ft(vol_h=vol_primary, pu=factor_pu,
                                                a=1.15, b=1.03, c=1.2, k=1.4,
                                                is_dc1=is_dc1, is_dc2=is_dc2, vol_l=vol_secondy)



    # check lower limit
    # per 233-C3c Limit: calc cannot be less than vertical clearance calculated for VH
    v_gnd1 = vol_primary / math.sqrt(2) if is_dc1 else vol_primary / math.sqrt(3)
    v_gnd2 = vol_secondy / math.sqrt(2) if is_dc1 else vol_secondy / math.sqrt(3)
    v_gnd = max(v_gnd1, v_gnd2)
    clear_vert_regular = get_adder_vol_in(vol_gnd=max(v_gnd1, v_gnd2), vol_base=22, kv_increment=0.4) / 12
    clear_vert_regular += (5.0 if is_com else 2.0)
    clear_vert_alter = max(clear_vert_alter, clear_vert_regular)

    # use vol_equiv to test applicability
    clear_vert_altitude = get_adder_altitude(terrain_ele=elevation_ter, adder_vol=clear_vert_alter,
                                              voltage_max=vol_equiv, base_ele=1500)

    clear_vert_alter += clear_vert_altitude

    # according to Table 233-3
    clear_vert_alter += 2.0 if is_com else 0.0
    clear_vert_alter = math.ceil(clear_vert_alter * 100) / 100

    # **************** ADD NOTES ********************************
    list_notes = _get_notes_rule235(params_rule235)
    for i, note in enumerate(list_notes):
        list_notes[i] = '{:2d}. '.format(i+1) + note

    notes = '\n'.join(list_notes)

    ###################### ASSIGN TO FORM FIELDS ########################################
    params_rule235['rst_hori_per_volt'] = clear_hori_vol
    params_rule235['rst_vert_at_support'] = clear_vert_at_support
    params_rule235['rst_hori_per_sag'] = clear_hori_sag
    params_rule235['rst_vert_in_span'] = clear_vert_in_span
    params_rule235['rst_hori_alternate'] = clear_hori_alter
    params_rule235['rst_vert_alternate'] = clear_vert_alter
    params_rule235['notes_area'] = notes


def _get_notes_rule235(params_rule235):
    list_notes = ['Refer to NESC Rule 235 and RUS 1724E-200 Chapter 6 for details.)',
                  'Experience factor, swing angle, sag and insulator length are used for calculation per RUS Bulletin 1724E-200.']

    # notes about maximum voltage
    v1 = float(params_rule235['circuit1'])
    v2 = float(params_rule235['circuit2'])
    is_same_circuit = (params_rule235['same_circuit'] == '0')
    is_dc1 = params_rule235['type_system1'] == '1'
    is_dc2 = params_rule235['type_system2'] == '1'
    v1_gnd = v1 if is_dc1 else v1 / math.sqrt(3)
    v2_gnd = v2 if is_dc2 else v2 / math.sqrt(3)
    elevation_ter = float(params_rule235['elevation_ter'])

    if (35 < v1_gnd < 50) or (35 < v2_gnd < 50):
        list_notes.append('Maximum voltage is used for 69kV to conform to industrial conservative practice although '
                          'its line-to-ground voltage is less than 50 kV.')
    if elevation_ter > 3300:
        list_notes.append('Clearance is adjusted per altitude {} ft'.format(elevation_ter))

    elif elevation_ter > 1500:
        list_notes.append('Clearance calculated using Alternate Method is adjusted per altitude {} ft'.format(elevation_ter))

    else:
        list_notes.append('Clearance is not adjusted per altitude {} ft that is below base elevation.'.format(elevation_ter))

    if is_dc1:
        list_notes.append('The first circuit assumes +/-{} kV pole-to-ground voltage (DC)'.format(v1))
    if is_dc2:
        list_notes.append('The second circuit assumes +/-{} kV pole-to-ground voltage (DC)'.format(v2))

    return list_notes


def _validate_input_rule235(params):
    try:
        vol_primary = float(params['circuit1'])
        vol_secondy = float(params['circuit2'])
        vol_factor1 = float(params['vol_multipler1'])
        vol_factor2 = float(params['vol_multipler2'])
        factor_experience = int(params['factor_experience'])
        angle_swing = float(params['angle_swing'])
        value_sag = float(params['value_sag'])
        len_insulator = float(params['len_insulator'])
        elevation_ter = float(params['elevation_ter'])
        pu = float(params['factor_pu'])
        return True

    except:
        return False
