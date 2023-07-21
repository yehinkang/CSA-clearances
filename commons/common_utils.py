import os, sys,  json, shutil, re, math
import pandas as pd
from datetime import datetime
import stateplane, utm, pyproj
import warnings
import zipfile
from flask import request
from urllib.parse import urlsplit, urlunsplit
from config import Config


def is_digit(data):
    if data.startswith('-'):
        data = data[1:]
    if data.replace('.', '', 1).isdigit():
        return True
    else:
        return False


def deg_to_dms(deg, t='lat'):
    if type(deg) is not float:
        deg = float(deg)
    if t == 'lat':
        deg_dir = 'N' if deg > 0 else 'S'
    else:
        deg_dir = 'E' if deg > 0 else 'W'

    deg = abs(deg)
    d = math.trunc(deg)
    m = math.trunc((deg - d) * 60)
    s = (deg - d - m / 60) * 3600
    s = round(s * 100) / 100
    return d, m, s, deg_dir

def get_coord_system_per_pls(pls_meta):
    fips = None
    epsg_zone = None
    utm_zone_number = None
    utm_zone_letter = None

    coord_system = pls_meta.get('coordinatesystem')
    if coord_system.lower() == 'nad83':
        fips = str(pls_meta.get('zone')).split('.')[0]
    
    elif coord_system.lower() == 'epsg':
        epsg_zone = str(pls_meta.get('zone')).split('.')[0]
    
    elif coord_system.lower() == 'utm':
        # utm zone per pls-cadd: 14N_WGS84
        zone = pls_meta.get('zone')
        regex = re.compile('[0-9]+')
        try:
            utm_zone_number = int(regex.findall(zone)[0])
            utm_zone_letter = zone.split('_')[0][-1]
        except ValueError as ve:
            print('>>> Failed to retrieve UTM zone number or zone letter')

    else:
        print('>>> Not NAD83 or UTM or EPSG coordinate system. Conversion is canceled.')
    
    coord_system = {'fips': fips, 'epsg_zone': epsg_zone, 'utm_zone_number': utm_zone_number, 'utm_zone_letter': utm_zone_letter}
    return coord_system


def convert_ft_to_meter(x, unit='ft_survey'):
    """
    INPUT: x - a value with unit=unit
    meter is always used when converting xy to lat/lon
    """
    if unit.lower() == 'meter':
        factor_to_m = 1
        factor_to_ft = 3.2808333333
    elif unit.lower() == 'ft_international':
        factor_to_m = 3.2808398950131
        factor_to_ft = 1
    elif unit.lower() == 'ft_survey':
        factor_to_m = 3.2808333333
        factor_to_ft = 1
    else:
        factor_to_m = None
        factor_to_ft = None
    x = x / factor_to_m

    return x

def spcs_to_latlon(easting, northing, fips):
    """
    State Plane Coordinate System to Lat/Lon
    easting, northing: meter
    fips: str, e.g. '4201'
    Return: tuple (lat, lon)
    """
    try:
        # padding with leading zeros to make fips 4 char long
        fips = '{:0>4}'.format(fips)
        latlon = stateplane.to_latlon(easting=easting, northing=northing, fips=fips)
        #latlon = (lonlat[1], lonlat[0])
    except Exception as e:
        print('>>>Error: Failed to convert State Plane Coordinates to Lat/Lon: ', e.__doc__)
        latlon = None
        
    return latlon


def utm_to_latlon(easting, northing, zone_number, zone_letter):
    """
    easting, northing: meter
    zone_number: int
    zone_letter: str
    northern: not to specify if zone_letter is used

    Return: tuple (lat, lon)
    """
    latlon = utm.to_latlon(easting=easting, northing=northing, zone_number=zone_number, zone_letter=zone_letter)
    return latlon

def wgs84_to_latlon(epsg_zone, easting, northing):
    """
    INPUT: 
        epsg_zone (a 4-digit number, e.g. '3857')
        easting, northing: meter

    OUTPUT: epsg:4326 that represents lat/lon
    
    REF: https://jingwen-z.github.io/how-to-convert-projected-coordinates-to-latitude-longitude-by-python/
    """
    
    """
    # TESTING
    fips = '3502'
    latlon = spcs_to_latlon(easting=2741020.842/3.28, northing=584577.503/3.28, fips=fips)
    """
    
    epsg_src = f'epsg:{epsg_zone}'
    epsg_dst = 'epsg:4326'
    transformer = pyproj.Transformer.from_crs(epsg_src, epsg_dst)
    latlon = transformer.transform(xx=easting, yy=northing)

    return latlon
    
def utm_from_latlon(lat, lon):
    """
    RETURN: (utm_x in meter, utm_y in meter, zone_number, zone_letter)
    """
    xy = utm.from_latlon(latitude=lat, longitude=lon)
    return xy


def get_float_validation(val, rst=None):
    '''
    INPUT: val can be either str or float
    RETURN: float for valid number, None for invalid number
        if rst is defined (not None), will return rst as default val in case validation fails
    '''
    is_valid = True

    if type(val) is str:
        if val.replace('.', '', 1).replace('-', '', 1).isdigit():
            val = float(val)
        else:
            is_valid = False

    elif type(val) is float or type(val) is int:
        pass
    else:
        is_valid = False

    return val if is_valid else rst


def compose_filename_utc(prefix=None, suffix=None, ext='json', folder=None):
    filename = datetime.utcnow().strftime('utc_%Y%m%d_%H%M%S%f')
    if prefix:
        filename = prefix + '_' + filename
    if suffix:
        filename += '_{:s}'.format(suffix)
    filename += '.{:s}'.format(ext)
    if folder is None:
        folder = 'tmp'
    
    filename = folder + '/' + filename
    
    return filename


def create_dir_utc(suffix=None, folder='tmp'):
    dir_utc = os.path.join(folder, datetime.utcnow().strftime('utc_%Y%m%d_%H%M%S%f'))
    if not os.path.isdir(folder):
        os.mkdir(folder)

    if suffix:
        dir_utc += '_{:s}'.format(suffix)
    if not os.path.isdir(dir_utc):
        try:
            os.mkdir(dir_utc)
            return dir_utc
        except OSError as error:
            print(error)
            return None


def folder_cleanup_or_create(path_dir):
    """
    if the path_dir exists, cleanup; else create it (empty)
    """
    if os.path.isdir(path_dir):
        # cleanup aug-image folder
        for root, dirs, files in os.walk(path_dir):
            for file in files:
                os.remove(os.path.join(root, file))
            
            # delete empty folders
            for dir in dirs:
                dir_path = os.path.join(root, dir)
                if len(os.listdir(dir_path)) == 0:
                    try:
                        shutil.rmtree(dir_path)
                    except Exception as e:
                        print(f'>>> Failed to delete folder {dir}: {e.__doc__}')

    else:
        # create a folder to save aug-images
        os.mkdir(path_dir)


def delete_folder_or_file(path_to_delete):
    if type(path_to_delete) is str:
        if os.path.isfile(path_to_delete):
            try:
                os.remove(path_to_delete)
            except Exception as e:
                print(f'>>> Err when removing file {path_to_delete}: {e.__doc__}')
        elif os.path.isdir(path_to_delete):
            try:
                shutil.rmtree(path_to_delete)
            except Exception as e:
                print(f'>>> Err when removing directory {path_to_delete}: {e.__doc__}')
        else:
            pass
    else:
        print(f'>>> Err when removing a path: not a string')


def cleanup_hour_old(n_hours=1):
    # check the tmp folder and cleanup all files that are older than one hour
    '''
    number of seconds since epoch: os.path.getmtime(filepath)
    epoch time is defined as (1970,1,1,0,0,0)

    Does NOT delete empty folders
    '''
    warnings.warn('Deprecated. Use "cleanup_per_hour_age" instead to clean up empty folders as well.')

    list_files = os.listdir('tmp')
    for file in list_files:
        filepath = 'tmp/' + file

        if os.path.isdir(filepath):
            continue

        time_file_modification = os.path.getmtime(filepath)
        time_epoch = datetime(1970, 1, 1, 0, 0, 0)
        time_now_epoch = (datetime.utcnow() - time_epoch).total_seconds()

        # hour age of the file
        time_delt_in_hour = (time_now_epoch - time_file_modification) / 3600
        if time_delt_in_hour > n_hours:
            os.remove(filepath)


def cleanup_per_hour_age(n_hours=1, folder='tmp'):
    """
    Check the folder and cleanup all files/subfolders that are older than one hour
        number of seconds since epoch: os.path.getmtime(filepath)
        epoch time is defined as (1970,1,1,0,0,0)

        DOES delete empty folders
    """
    
    if os.path.isdir(folder):
        for root, dirs, files in os.walk(folder):
            for file in files:
                if 'do_not_delete' in file:
                    continue
                
                file_path = os.path.join(root, file)
                hour_age = get_file_hour_age(file_path=file_path)

                try:
                    if hour_age > n_hours:
                        os.remove(path=file_path)
                except Exception as e:
                    print(f'>>> File {file_path} exists but cannot be deleted: {e.__doc__}')

            # delete empty folders
            for dir in dirs:
                dir_path = os.path.join(root, dir)
                if len(os.listdir(dir_path)) == 0:
                    try:
                        shutil.rmtree(dir_path)
                    except Exception as e:
                        print(f'>>> Failed to delete folder {dir}: {e.__doc__}')



def get_download_link(filename=None, spec_folder=None):
    """
    INPUTS:
        - filename: full path of the file to be downloaded; if None, link to spec_folder will be returned
        - spec_folder: if None, nothing will be passed to download process, and 'tmp' folder is used as default
                       if sepcified, link will be create with "static={sepc_folder}" as a parameter passed to the download process;
                       by doing this, the download will point to the correct folder.

    RETURN: a downloadable link
    """
    
    if not os.path.isfile(filename):
        print(f'>>> File {filename} does NOT exist.')
        return None

    try:
        url_full = request.base_url
        url_split = urlsplit(url_full)
        url_home_api = f'{url_split.scheme}://{url_split.netloc}'
    except:
        url_home_api = Config.URL_HOME

    filename_base = os.path.basename(filename)
    link_download = f'{url_home_api}/download/{filename_base}'

    if spec_folder:
        link_download += f'?static={spec_folder}'

    return link_download


def get_download_linkdir(dirname=None, src='tmp'):
    """
    INPUTS: dirname - the folder name (not path) that holds the files to be requested for downloading
            src - 'tmp' by default: will refer to the server tmp/ folder
                  'share': will refer to the tlbushare/tlbutmp folder on storage account to avoid overflow 
    """
    if dir is None or (not os.path.isdir(dirname)):
        print(f'>>> dir {dirname} does not exist...')
        return None

    try:
        url_full = request.base_url
        url_split = urlsplit(url_full)
        url_home_api = f'{url_split.scheme}://{url_split.netloc}'
        print('url_full=',url_full)
        print('::: url_home_api=', url_home_api)
    except Exception as e:
        print(f'::: Cannot get base_url from request; using static url_home_api: {e.__doc__}')
        url_home_api = Config.URL_HOME

    dirname_base = os.path.basename(dirname)
    link_download = f'{url_home_api}/downloaddir/{dirname_base}?src={src}'

    return link_download


def get_file_hour_age(file_path):
    time_file_modification = os.path.getmtime(file_path) # modification time since epoch in seconds
    time_epoch = datetime(1970, 1, 1, 0, 0, 0) # epoch time = 0
    time_now_epoch = (datetime.utcnow() - time_epoch).total_seconds()

    # hour age of the file
    time_delt_in_hour = (time_now_epoch - time_file_modification) / 3600
    return time_delt_in_hour


def get_modification_time(filepath):
    """
    UTC Time of file modification
    """
    if os.path.isfile(filepath):
        time_stamp = os.path.getmtime(filename=filepath)
        time_utc = datetime.utcfromtimestamp(time_stamp).strftime('%Y-%m-%d %H:%M:%S') + ' UTC'
    else:
        time_utc = None
    return time_utc


def find_all_urls_in_string(text):
    '''
    Example:
    'Cases 5 and 6 are not applicable to supply conductors of same phase and polarity (see rule <a href="https://ia.cpuc.ca.gov/gos/GO95/go_95_rule_54_4.html#c3c" target="_blank">54.4–C3c</a>), Insulated supply conductors in multi–conductor cables (see rule <a href="https://ia.cpuc.ca.gov/gos/GO95/go_95_rule_57_4.html#c" target="_blank">57.4–C</a>) or communication insulated conductors or multiple–conductor cables (see rule <a href="https://ia.cpuc.ca.gov/gos/GO95/go_95_rule_87_4.html#c" target="_blank">87.4–C1</a>).'
    '''
    regex = r"(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'\".,<>?«»“”‘’]))"
    urls = re.findall(regex, text)
    if urls:
        for i, url in enumerate(urls):
            if type(url) is tuple:
                urls[i] = url[0]
    return urls


def process_response(response):
    '''
    RESPONSE STATUS CODES: https://developer.mozilla.org/en-US/docs/Web/HTTP/Status
    Informational responses (100–199)
    Successful responses (200–299)
    Redirection messages (300–399)
    Client error responses (400–499)
    Server error responses (500–599)
    '''

    try:
        status_code = response.status_code
    except:
        status_code = -500

    if response.ok:
        # status_code is 200-299
        if type(response.text) is str:
            if type(response) is int:
                response_json = {
                    'error_msg': [f'Status Code: {response}',
                                  'This may be caused by restricted access; or content being under development.']
                }
            else:
                response_json = json.loads(response.text)

        else:
            response_json = {
                'error_msg': ['Internal issue - response cannot be recognized',
                              'Most likely caused by conflict with other process.',
                              'Contact developer if this continues to occur.']
            }
    elif status_code == -500:
        response_json = {'error_msg': ['The Server does not return a valid response. '
                                  'Contact the Developer if this continues to occur.']}
    else:
        response_code = f'Internal error with status code: {status_code}'
        response_json = {
            'error_msg': [response_code, 'This may be due to huge file size; conflict with other process; '
                                         'or incorrect input content.',
                          'Contact the Developer if this continues to occur.']
        }

    return response_json


def get_ge_txt_header(line_id=None):
    if line_id is None:
        line_id = 'Structure List'
    file_id = line_id.replace(' ', '_')
    txt = '<?xml version="1.0" encoding="UTF-8"?>\n'
    txt += '<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">\n'
    txt += '<Document>\n'
    txt += f'\t<name>{file_id}</name>\n'
    txt += '\t<Style id="s_ylw-pushpin_hl00">\n'
    txt += '\t\t<IconStyle>\n'
    txt += '\t\t\t<scale>1.3</scale>\n'
    txt += '\t\t\t<Icon>\n'
    txt += '\t\t\t\t<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>\n'
    txt += '\t\t\t</Icon>\n'
    txt += '\t\t\t<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
    txt += '\t\t</IconStyle>\n'
    txt += '\t</Style>\n'
    txt += '\t<Style id="s_ylw-pushpin2">\n'
    txt += '\t\t<IconStyle>\n'
    txt += '\t\t\t<scale>1.1</scale>\n'
    txt += '\t\t\t<Icon>\n'
    txt += '\t\t\t\t<href>http://maps.google.com/mapfiles/kml/pushpin/pink-pushpin.png</href>\n'
    txt += '\t\t\t</Icon>\n'
    txt += '\t\t\t<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
    txt += '\t\t</IconStyle>\n'
    txt += '\t\t<ListStyle>\n'
    txt += '\t\t</ListStyle>\n'
    txt += '\t</Style>\n'

    txt += '\t<StyleMap id="msn_R">\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>normal</key>\n'
    txt += '\t\t\t<styleUrl>#sn_R</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>highlight</key>\n'
    txt += '\t\t\t<styleUrl>#sh_R</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t</StyleMap>\n'

    txt += '\t<StyleMap id="msn_grn-pushpin">\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>normal</key>\n'
    txt += '\t\t\t<styleUrl>#sn_grn-pushpin</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>highlight</key>\n'
    txt += '\t\t\t<styleUrl>#sh_grn-pushpin</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t</StyleMap>\n'
    txt += '\t<StyleMap id="m_ylw-pushpin3">\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>normal</key>\n'
    txt += '\t\t\t<styleUrl>#s_ylw-pushpin00</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>highlight</key>\n'
    txt += '\t\t\t<styleUrl>#s_ylw-pushpin_hl</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t</StyleMap>\n'
    txt += '\t<StyleMap id="msn_wht-pushpin">\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>normal</key>\n'
    txt += '\t\t\t<styleUrl>#sn_wht-pushpin</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>highlight</key>\n'
    txt += '\t\t\t<styleUrl>#sh_wht-pushpin</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t</StyleMap>\n'
    txt += '\t<StyleMap id="m_ylw-pushpin00">\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>normal</key>\n'
    txt += '\t\t\t<styleUrl>#s_ylw-pushpin20</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>highlight</key>\n'
    txt += '\t\t\t<styleUrl>#s_ylw-pushpin_hl00</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t</StyleMap>\n'
    txt += '\t<Style id="sn_grn-pushpin">\n'
    txt += '\t\t<IconStyle>\n'
    txt += '\t\t\t<scale>1.1</scale>\n'
    txt += '\t\t\t<Icon>\n'
    txt += '\t\t\t\t<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>\n'
    txt += '\t\t\t</Icon>\n'
    txt += '\t\t\t<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
    txt += '\t\t</IconStyle>\n'
    txt += '\t\t<BalloonStyle>\n'
    txt += '\t\t</BalloonStyle>\n'
    txt += '\t\t<ListStyle>\n'
    txt += '\t\t</ListStyle>\n'
    txt += '\t</Style>\n'
    txt += '\t<Style id="s_ylw-pushpin_hl000">\n'
    txt += '\t\t<IconStyle>\n'
    txt += '\t\t\t<scale>1.3</scale>\n'
    txt += '\t\t\t<Icon>\n'
    txt += '\t\t\t\t<href>http://maps.google.com/mapfiles/kml/pushpin/pink-pushpin.png</href>\n'
    txt += '\t\t\t</Icon>\n'
    txt += '\t\t\t<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
    txt += '\t\t</IconStyle>\n'
    txt += '\t\t<ListStyle>\n'
    txt += '\t\t</ListStyle>\n'
    txt += '\t</Style>\n'
    txt += '\t<Style id="sh_grn-pushpin">\n'
    txt += '\t\t<IconStyle>\n'
    txt += '\t\t\t<scale>1.3</scale>\n'
    txt += '\t\t\t<Icon>\n'
    txt += '\t\t\t\t<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>\n'
    txt += '\t\t\t</Icon>\n'
    txt += '\t\t\t<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
    txt += '\t\t</IconStyle>\n'
    txt += '\t\t<BalloonStyle>\n'
    txt += '\t\t</BalloonStyle>\n'
    txt += '\t\t<ListStyle>\n'
    txt += '\t\t</ListStyle>\n'
    txt += '\t</Style>\n'
    txt += '\t<Style id="s_ylw-pushpin00">\n'
    txt += '\t\t<IconStyle>\n'
    txt += '\t\t\t<scale>1.1</scale>\n'
    txt += '\t\t\t<Icon>\n'
    txt += '\t\t\t\t<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>\n'
    txt += '\t\t\t</Icon>\n'
    txt += '\t\t\t<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
    txt += '\t\t</IconStyle>\n'
    txt += '\t</Style>\n'
    txt += '\t<Style id="sh_wht-pushpin">\n'
    txt += '\t\t<IconStyle>\n'
    txt += '\t\t\t<scale>1.3</scale>\n'
    txt += '\t\t\t<Icon>\n'
    txt += '\t\t\t\t<href>http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png</href>\n'
    txt += '\t\t\t</Icon>\n'
    txt += '\t\t\t<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
    txt += '\t\t</IconStyle>\n'
    txt += '\t\t<BalloonStyle>\n'
    txt += '\t\t</BalloonStyle>\n'
    txt += '\t\t<ListStyle>\n'
    txt += '\t\t</ListStyle>\n'
    txt += '\t</Style>\n'
    txt += '\t<Style id="sn_wht-pushpin">\n'
    txt += '\t\t<IconStyle>\n'
    txt += '\t\t\t<scale>1.1</scale>\n'
    txt += '\t\t\t<Icon>\n'
    txt += '\t\t\t\t<href>http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png</href>\n'
    txt += '\t\t\t</Icon>\n'
    txt += '\t\t\t<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
    txt += '\t\t</IconStyle>\n'
    txt += '\t\t<BalloonStyle>\n'
    txt += '\t\t</BalloonStyle>\n'
    txt += '\t\t<ListStyle>\n'
    txt += '\t\t</ListStyle>\n'
    txt += '\t</Style>\n'
    txt += '\t<StyleMap id="m_ylw-pushpin000">\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>normal</key>\n'
    txt += '\t\t\t<styleUrl>#s_ylw-pushpin2</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t\t<Pair>\n'
    txt += '\t\t\t<key>highlight</key>\n'
    txt += '\t\t\t<styleUrl>#s_ylw-pushpin_hl000</styleUrl>\n'
    txt += '\t\t</Pair>\n'
    txt += '\t</StyleMap>\n'
    txt += '\t<Style id="s_ylw-pushpin_hl">\n'
    txt += '\t\t<IconStyle>\n'
    txt += '\t\t\t<scale>1.3</scale>\n'
    txt += '\t\t\t<Icon>\n'
    txt += '\t\t\t\t<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>\n'
    txt += '\t\t\t</Icon>\n'
    txt += '\t\t\t<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
    txt += '\t\t</IconStyle>\n'
    txt += '\t</Style>\n'
    txt += '\t<Style id="s_ylw-pushpin20">\n'
    txt += '\t\t<IconStyle>\n'
    txt += '\t\t\t<scale>1.1</scale>\n'
    txt += '\t\t\t<Icon>\n'
    txt += '\t\t\t\t<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>\n'
    txt += '\t\t\t</Icon>\n'
    txt += '\t\t\t<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
    txt += '\t\t</IconStyle>\n'
    txt += '\t</Style>\n'

    txt += '\t<Folder>\n'
    txt += f'\t\t<name>{line_id}</name>\n'
    txt += '\t\t<open>1</open>\n'

    return txt


def get_ge_txt_closure():
    txt = '\t</Folder>\n'
    txt += '</Document>\n'
    txt += '</kml>\n'
    return txt


def get_zipped_file(filepath, ext_new=None, del_src=False):
    if ext_new is None:
        ext_new = 'zip'
    file_name, ext = os.path.splitext(filepath)
    filepath_new = file_name + '.' + ext_new

    # define a zfile and write to the same folder
    zfile = zipfile.ZipFile(filepath_new, 'w')
    zfile.write(filepath, os.path.basename(filepath))

    if del_src:
        os.remove(filepath)

    return filepath_new


def get_zipped_file_per_folder(folder):
    return ''


def read_file_text(filepath):
    """
    Read input from text file representing pls-pole
    Return: a list of line content after trimming line break symbol
    """

    """
    Read file with OS default encoding. If failed, try the other two encodings: latin-1 and ascii 
        (data will not be corrupted if it is simply read in, processed as ASCII and written back)
        ref: http://python-notes.curiousefficiency.org/en/latest/python3/text_file_processing.html
    """
    list_content = None
    try:
        with open(file=filepath, mode='r') as fh:
            list_content = fh.readlines()
    except UnicodeDecodeError as decodeErr:
        print(f'::: Failed to read text file with default encoding: {decodeErr.__doc__}')
        print(f'::: Trying next encoding method (latin-1)...')
    
    if not list_content:
        try:
            with open(file=filepath, mode='r', encoding='latin-1') as fh:
                list_content = fh.readlines()
        except UnicodeDecodeError as decodeErr:
            print(f'::: Failed to read text file with encoding=latin-1: {decodeErr.__doc__}')
            print(f'::: Trying next encoding method (ascii)...')

    if not list_content:
        try:
            with open(file=filepath, mode='r', encoding='ascii', errors='surrogateescape') as fh:
                list_content = fh.readlines()
        except UnicodeDecodeError as decodeErr:
            print(f'::: Failed to read text file with encoding=ascii and errors=surrogateescape: {decodeErr.__doc__}')
            print(f'>>> Error - Failed to read text file after using various encoding. Stopped trying...')
    
    if list_content:
        # remove line break symbol for each line
        list_content = [txt.replace('\n', '') for txt in list_content]
    else:
        print('>>> I/O Error: empty content returned from reading file {filepath}')

    return list_content


def read_spreadsheet(filepath, sheetname=None, to_str=False, dtype=None):
    """
    Read csv or xlxs into dataframe, then convert to 2D list
    to_str: convert values to string if True
    dtype: type name or dict of column -> e.g. {'a': np.float64, 'b: np.int32, 'c': str}
    """
    _, file_ext = os.path.splitext(filepath)

    error_msg = []

    if not (file_ext == '.csv' or file_ext == '.xlsx'):
        error_msg.append('Incorrect data file format.')
    elif file_ext == '.csv':
        # csv file found
        df = pd.read_csv(filepath)
    else:
        # xlxs file found
        try:
            if sheetname:
                df = pd.read_excel(filepath, sheet_name=sheetname, dtype=dtype)
            else:
                df = pd.read_excel(filepath, dtype=dtype)
            
        except Exception as e:
            error_msg.append(f'Failed to read data from xlxs file: {e.__doc__}.')
            print(f'>>>Error: Failed to read data from xlxs file: {e.__doc__}.')
    if not error_msg:
        list_data = df.values.tolist()


        if to_str:
            list_data = [[str(cell) for cell in row] for row in list_data]
    else:
        list_data = []
        
    return {'error_msg': error_msg, 'list_data': list_data}



"""
def html_to_pdf(html, pdf):
    app_ = QtWidgets.QApplication(sys.argv)

    page = QtWebEngineWidgets.QWebEnginePage()

    print(f'html={html}: exist={os.path.isfile(html)}')
    print(sys.argv)

    def handle_print_finished(filename, status):
        print(f'Finished converting {filename} with status={status}')
        QtWidgets.QApplication.quit()

    def handle_load_finished(status):
        if status:
            page.printToPdf(pdf)
        else:
            print(f'Failed to convert {html}')
            QtWidgets.QApplication.quit()

    page.pdfPrintingFinished.connect(handle_print_finished)
    page.loadFinished.connect(handle_load_finished)
    page.load(QtCore.QUrl.fromLocalFile(html))
    app_.exec_()
"""


def test_read_spreadsheet():
    file_path = 'C:/Users/jdong/OneDrive - POWER Engineers, Inc/Documents/ProjectsPyCharm/references/table_employee_line.xlsx'
    read_spreadsheet(filepath=file_path, dtype={'Account Location': str, 'Department Number': str})

def test_xy_to_latlon():
    # stateplane

    # WGS84
    epsg_zone = '3780'
    x = -8312.999
    y = 5646954.142
    latlon = wgs84_to_latlon(epsg_zone=epsg_zone, easting=x, northing=y)
    print(latlon)
    

if __name__ == '__main__':
    
    test_xy_to_latlon()