'''Russell W. Myers'''
'''https://github.com/rwmyers46'''
'''https://www.linkedin.com/in/myersrussell/'''

import pandas as pd
import numpy as np
import tabula
import glob
import sys
import os
import re
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
from PIL import Image 
import pytesseract
from pdf2image import convert_from_path
import shutil
import warnings
warnings.filterwarnings("ignore")

# update this path: copy / paste file location for local filepath after the r' below:
pytesseract.pytesseract.tesseract_cmd = r'location_of_local_tesseract_dir'

# define class object:

class Mode:
    def __init__(self, name, sort_key, scan_area, ship_scan, order_scan, 
                 index_val, shipping_svc, op_select, csvx, slips_path, labels_path):
        self.name = name
        self.sort_key = sort_key
        self.scan_area = scan_area
        self.ship_scan = ship_scan
        self.order_scan = order_scan
        self.index_val = index_val
        self.shipping_svc = shipping_svc
        self.op_select = op_select
        self.csvx = csvx
        self.slips_path = slips_path
        self.labels_path = labels_path
        
def Startup():
    '''Run I/O block'''
    csvx = 'y'
    labels_path = ''
    label_type = ''
    s1 = 'Enter file path for Packing Slips: '
    s2 = 'Would you like a CSV export? [y/n]: '
    s3 = 'Enter file path for Shipping Labels: '
    
    print('[1] Export packing slip data to CSV.')
    print('[2] Reorder packing slips by item ID.')
    print('[3] Match & reorder packing slips by shipping labels.')

    function = int(input('\nSelect Option by Number: '))
    if function < 1 or function > 3:
        print('Incorrect entry...\n')
        startup()
    else:
        if function == 1:
            slips_path = input(s1).strip()
        elif function == 2:
            slips_path = input(s1).strip()
            csvx = input(s2)
        elif function == 3:
            slips_path = input(s1).strip()
            labels_path = input(s3).strip()
            csvx = input(s2).strip()
        return(function, slips_path, labels_path, csvx)        
    
def pdf_splitter(path, page_order, fname, temp_dir, no_match_keys):
    '''Separate pdf into individual, uniquely-named pages'''
    pdf = PdfFileReader(path)
    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))
        
        if mode.op_select == 2:
            output_filename = '{}tmp_{}_page_{}.pdf'.format(temp_dir, fname, page_order[page + 1])
        else:
            if page in no_match_keys:
                output_filename = 'no_match_page_{}.pdf'.format(page)
            else:
                output_filename = '{}tmp_{}_page_{}.pdf'.format(temp_dir, fname, page_order[page])

        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)
            out.close()

def merger(output_path, input_paths):
    '''Merge individual pdf pages to common file'''
    pdf_merger = PdfFileMerger()
 
    for path in input_paths:
        pdf_merger.append(path)
 
    with open(output_path, 'wb') as fileobj:
        pdf_merger.write(fileobj)
        
def hasNumbers(string):
    '''Determine whether given string has numbers'''
    return(any(char.isdigit() for char in string))

def format_string(j):
    '''Standardize string formatting'''
    j = j.upper().replace(',','')
    return(j)

def city_state(data):
    '''Separate city, state, zip code'''
    comma = data.find(',')
    city = format_string(data[0:comma])
    t = data[comma:].strip(",").split()
    state = format_string(t[0])
    zip_code = t[1]
    return(city, state, zip_code)

def strip_Addr(q):
    '''Separate mailing address into individual components'''
    addr_keys = ['full_name', 'first_name', 'last_name', 'company', 'addr_line1', 
                 'addr_line2', 'addr_line3', 'city', 'state', 'zip_code', 'csz']
    addr_book = dict.fromkeys(addr_keys, '')

    for i, j in enumerate(q):
        if i == 0:
            try:
                t = j.split()
                addr_book['first_name'] = format_string(t[0])
                addr_book['last_name'] = format_string(t[1])
                addr_book['full_name'] = format_string(j)
            except:
                addr_book['full_name'] = format_string(j)
        if i == 1 and hasNumbers(j) == False:
            addr_book['addr_line1'] = format_string(j)
            addr_book['company'] = format_string(j)
        elif i == 1:
            addr_book['addr_line1'] = format_string(j)
        elif i == 2 and len(q) > 3:
            addr_book['addr_line2'] = format_string(j)
        elif i == 3 and len(q) == 5:
            addr_book['addr_line3'] = format_string(j)
        if i == len(q) - 1:
            CSZ = city_state(j)
            addr_book['city'] = CSZ[0]
            addr_book['state'] = CSZ[1]
            addr_book['zip_code'] = CSZ[2]
            addr_book['csz'] = ' '.join(CSZ)
    return(addr_book)

def labels_Ripper(text):
    '''Clean and standardize address strings'''
    label_keys = ['full_name', 'addr_line1', 'addr_line2', 'addr_line3', 'addr_line4']
    label_vals = dict.fromkeys(label_keys, '')
    lab_txt = text.splitlines()
    try:
        lab_txt = list(filter(lambda x: x != '', lab_txt))
        lab_txt = list(filter(None, [re.sub('\d{10}', '', i) for i in lab_txt]))
        if text.lower().find('ship to:') != -1:
            try:
                lab_idx = lab_txt.index('SHIP TO:') + 1
            except:
                lab_idx = 0
        elif text.lower().find('to') != -1:
            try:
                lab_idx = list(filter(lambda x: 'TO' in x, lab_txt))[0]
                lab_idx = lab_txt.index(lab_idx)
            except:
                lab_idx = list(filter(lambda x: 'BILL 3rd PARTY' in x, lab_txt))[0]
                lab_idx = lab_txt.index(lab_idx) + 1

        addr = lab_txt[lab_idx:]
        addr = list(filter(None, [re.sub('^(\+\d{1,2}\s)?\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}$','', i) for i in addr]))
        addr = list(filter(lambda x: 'REF:' not in x, addr))
        try:
            addr[0] = addr[0].replace('TO', '').strip(' ')
        except:
            pass

        label_vals['full_name'] = format_string(addr[0])
        for i in range(1, len(addr)):
            label_vals[label_keys[i]] = addr[i].rstrip()
    except:
        label_vals['full_name'] = 'Label_Error'
    return(label_vals)

def page_Mod(order_dict):
    '''Modify page number formatting'''
    for k, v in order_dict.items():
        if order_dict[k] < 10:
            order_dict[k] = '000' + str(order_dict[k])
        elif order_dict[k] >= 10 and order_dict[k] < 100:
            order_dict[k] = '00' + str(order_dict[k])
        else:
            order_dict[k] = '0' + str(order_dict[k])
    return(order_dict)

def Read_Labels():
    '''Extract address label text to dict and create error log file'''
    c_labels = dict()
    label_errors = []
    # remove residual temp dir or create new
    temp_dir = 'labels_temp/'
    if os.path.isdir(temp_dir):
        shutil.rmtree(temp_dir)
    os.mkdir(temp_dir)

    # convert label pdfs to jpgs:
    pages = convert_from_path(mode.labels_path, dpi=250)
    print('\nProcessing shipping labels...')
    for idx, page in enumerate(pages):   #idx doubles as the labels page order
        tmp_name = temp_dir + str(idx) + '.jpg'
        page.save(tmp_name, 'JPEG')
        Image.open(tmp_name).crop((5, 150, 850, 455)).convert('L').save(tmp_name)
        text = str(((pytesseract.image_to_string(Image.open(tmp_name))))) 
        c_labels[idx] = labels_Ripper(text)
        ref = idx + 1
        sys.stdout.write(f'\rReading Shipping Label: %d/{len(pages)}' % ref)
        sys.stdout.flush()
    shutil.rmtree(temp_dir)

    # process label error to new file:
    for i, j in c_labels.items():
        if j['full_name'] == 'Label_Error':
            label_errors.append(i)
            
    if len(label_errors) > 0:
        print(f'\n\n >>Label Read Error... Printing {len(label_errors)} labels to Label_Errors.pdf')
        temp_dir = 'lab_errors_temp/'
        if os.path.isdir(temp_dir):
            shutil.rmtree(temp_dir)
        os.mkdir(temp_dir)
        
        errors = convert_from_path(mode.labels_path, dpi=200)
        for idx, err in enumerate(errors):   #idx doubles as the labels page order
            if idx in label_errors:
                img_path = temp_dir + str(idx) + '.jpg'
                err.save(img_path, 'JPEG')
                img = Image.open(img_path)
                label_errors[label_errors.index(idx)] = img

        lab_err_filename = 'Label_Errors.pdf'
        label_errors[0].save(lab_err_filename, 'PDF' ,resolution=100.0, save_all=True, append_images=label_errors)
        shutil.rmtree(temp_dir)

    print('')
    return(c_labels)

def csv_XP(order_info, d, shipping, addr_dict, file_name):
    '''Export order & label information to CSV file'''
    xport_list = []
    print(f'\nExporting addresses to CSV...')

    for k, v in addr_dict.items():
        if mode.name == 'BedBath':
            mode.shipping_svc = shipping.iloc[k]
            
        xport_load = {
                'Order Number' : order_info.Order_Number.iloc[k],
                'Item Quantity' : int(d[k]['QTY'][0]),
                'Item Marketplace ID' : d[k][mode.sort_key][0],
                'Order Requested Shipping Service' : mode.shipping_svc,
                'Recipient Full Name' : v['full_name'],
                'Recipient First Name' : v['first_name'],
                'Recipient Last Name' : v['last_name'],
                'Recipient Company' : v['company'],
                'Address Line 1' : v['addr_line1'],
                'Address Line 2' : v['addr_line2'],
                'Address Line 3' : v['addr_line3'],
                'City' : v['city'],
                'State' : v['state'],
                'Postal Code' : v['zip_code']
            }
        xport_list.append(xport_load)

        if d[k]['Unique_Skews'][0] > 1:
            xplc = xport_load.copy()
            for i in range(1, d[k]['Unique_Skews'][0]):
                xplc['Item Marketplace ID'] = d[k][mode.sort_key][i]
                xplc['Item Quantity'] = int(d[k]['QTY'][i])
                xport_list.append(xplc)

    df_xport = pd.DataFrame(xport_list, columns = cols)
    df_xport.to_csv(file_name + '_Export.csv')
    print('Export Complete!')

def match_Labels(addr_dict, clean_labels):
    '''Match extracted contact data from shipping labels to packing slips'''
    order_dict = dict()
    no_match_keys = []
    dupes_idx = []

    val_list_names = [k['full_name'] for i, k in addr_dict.items()]
    val_list_addr1 = [k['addr_line1'] for i, k in addr_dict.items()]
    val_list_addr2 = [k['addr_line2'] for i, k in addr_dict.items()]
    val_list_csz = [k['csz'] for i, k in addr_dict.items()]
    
    # catch duplicate names
    for a, b in enumerate(val_list_names):
        if val_list_names.count(b) > 1:
            dupes_idx.append(a)

    for j, k in clean_labels.items():
        if k['full_name'] in val_list_names:
            order_dict[j] = val_list_names.index(k['full_name'])  
            dup = order_dict[j]
        elif k['addr_line1'] in val_list_addr1:
            order_dict[j] = val_list_addr1.index(k['addr_line1'])
        elif k['addr_line2'] in val_list_csz:
            order_dict[j] = val_list_csz.index(k['addr_line2'])
        elif k['addr_line3'] in val_list_csz:
            order_dict[j] = val_list_csz.index(k['addr_line3']) 
        else:
            print('No slip for: ', k)
        try:
            dup = order_dict[j]
            if dup in dupes_idx:   # remove duplicate vals
                val_list_names[dup] = 'xxx'
        except:
            pass
            
    order_keys = list(order_dict.values())
    for k, v in addr_dict.items():
        if k not in order_keys:
            no_match_keys.append(k)
                
    print(f' \nPages to be printed to No_Match_Slips: {len(no_match_keys)}')
    return(order_dict, no_match_keys)

def bb_ship():
    '''Strip order data from Bed, Bath, and Beyond invoices'''
    # collect shipping info: this is for BedBath
    shipping = tabula.read_pdf(mode.slips_path, area = mode.order_scan, pages = 'all', pandas_options={'header': None})
    order_info = pd.DataFrame({'Order_Number':shipping[1].iloc[::3].values}) #review this line
    
    shipping = shipping.iloc[:,-1:]
    shipping.rename(columns= {3:'Shipping'}, inplace = True)
    shipping.drop(shipping[shipping.Shipping.isna()].index, inplace=True)
    shipping.reset_index(drop = True, inplace = True)
    shipping = shipping.Shipping

    # process recipient info:
    ship_to = tabula.read_pdf(mode.slips_path, area = mode.ship_scan, stream = False, pages = 'all', header = None)
    ship_to.rename(columns= {'Shipped To:' : 'Shipping_Addr'}, inplace = True)
    return(order_info, shipping, ship_to)

def tar_ship():
    '''Strip order data from Target invoices'''
    order_info = tabula.read_pdf(mode.slips_path, area = mode.order_scan, pages = 'all', pandas_options={'header': None})
    order_info.rename(columns = {1:'Order_Number'}, inplace = True)
    
    ship_to = tabula.read_pdf(mode.slips_path, area = mode.ship_scan, pages = 'all', pandas_options={'header': None})
    ship_to.rename(columns= {1: 'Shipping_Addr'}, inplace = True)
    ship_to.drop(columns = 0, inplace = True)
    ship_to.drop(ship_to[ship_to.Shipping_Addr.isna()].index, inplace=True)
    ship_to = ship_to[1:].reset_index(drop = True)
    return(order_info, ship_to)

def slips_Reorder(d):
    '''Reorder packing slips alphanumerically by part numbers'''
    df = pd.concat(d.values(), ignore_index=True)
    
    df_mul = df[df.Unique_Skews > 1]
    df_mul.drop_duplicates('Page', inplace = True)
    df_mul.Page.astype(int)
    df_mul['Order'] = df_mul.Page / 1000
    
    df_sin = df[df.Unique_Skews == 1]
    df_sin.sort_values(by = mode.sort_key, inplace = True)
    df_sin.reset_index(drop = True, inplace = True)
    df_sin['Order'] = df_sin.index + 1

    df2 = pd.concat([df_mul, df_sin], ignore_index=True)
    df2.Order = df2.index
    
    df2['NewOrder'] = ''
    for idx, row in df2.iterrows():
        if idx < 10:
            temp = '000' + str(row.Order)
        elif (idx >= 10 and idx < 100):
            temp = '00' +  str(row.Order)
        else:
            temp = '0' + str(row.Order)
        df2.at[idx,'NewOrder'] = temp
    return(df2)

def slip_Sort(temp_dir):
    '''Sort packing slip file names'''
    slips = os.listdir(temp_dir)
    slips.sort()
    if '.ipynb_checkpoints' in slips:
        slips.remove('.ipynb_checkpoints')
    slips = [temp_dir + str(i) for i in slips]
    return(slips)

## launch main code block:
if __name__ == "__main__":
    
    ## column names for export
    cols = ['Order Number', 'Order Created Date', 'Order Date Paid', 'Order Total',
       'Order Amount Paid', 'Order Tax Paid', 'Order Shipping Paid',
       'Order Requested Shipping Service', 'Order Total Weight (oz)',
       'Order Custom Field 1', 'Order Custom Field 2', 'Order Custom Field 3',
       'Order Source', 'Order Notes from Buyer', 'Order Notes to Buyer',
       'Order Internal Notes', 'Order Gift Message', 'Order Gift - Flag',
       'Buyer Full Name', 'Buyer First Name', 'Buyer Last Name', 'Buyer Email',
       'Buyer Phone', 'Buyer Username', 'Recipient Full Name',
       'Recipient First Name', 'Recipient Last Name', 'Recipient Phone',
       'Recipient Company', 'Address Line 1', 'Address Line 2',
       'Address Line 3', 'City', 'State', 'Postal Code', 'Country Code',
       'Item SKU', 'Item Name', 'Item Quantity', 'Item Unit Price',
       'Item Weight (oz)', 'Item Options', 'Item Warehouse Location',
       'Item Marketplace ID']

    # call startup to define initial variables:
    op_select, slips_path, labels_path, csvx = Startup()

    addr_index = [0]
    d = dict()
    addr_dict = dict()
    file_name = os.path.splitext(os.path.basename(slips_path))[0]

    # automated mode selector
    if slips_path.lower().find('target') != -1:
        mode = Mode('Target', 'MFG ID', (210, 10, 400, 575), 
                    (100, 250, 225, 575), (5, 350, 15, 575),
                    'SEND TO:', '', op_select, csvx, slips_path, labels_path)
    else:
        mode = Mode('BedBath', 'Vendor Part #', 
                    (160.52, 20.12, 225.91, 572.59), 
                    (600, 305, 750, 600), (10, 200, 75, 600), 
                    'Shipped To:', '', op_select, csvx, slips_path, labels_path)

    print('\nMode:', mode.name)

    # get number of packing slips
    with open(mode.slips_path, "rb") as filehandle:
        pdf = PdfFileReader(filehandle)
        total_pages = pdf.getNumPages()
        filehandle.close()

    print('Processing packing slips...')
    for i in range(total_pages):
        ref = i + 1
        d[i] = tabula.read_pdf(mode.slips_path, area = mode.scan_area, stream = False, pages = ref)
        d[i].drop(d[i][d[i].QTY.isna()].index, inplace=True)
        d[i]['Page'] = ref
        d[i]['Unique_Skews'] = d[i].shape[0]
        d[i].reset_index(inplace = True, drop = True)
        sys.stdout.write(f'\rReading Packing Slip: %d/{total_pages}' % ref)
        sys.stdout.flush()

    if mode.op_select != 1:
    # create temp directory / remove residual verions
        temp_dir = 'pdf_rip_temp/'
        if os.path.isdir(temp_dir):
            shutil.rmtree(temp_dir)
        os.mkdir(temp_dir)

    # option 2 process:
    if mode.op_select == 2:
        # use dataframe to clean & reorder labels
        df2 = slips_Reorder(d)
        order_dict = dict(zip(df2.Page, df2.NewOrder))
        # divide packing slips and assign file names corresponding to new order
        pdf_splitter(mode.slips_path, order_dict, file_name, temp_dir, [])
        # sort packing slips by filename and combine with merger 
        slips = slip_Sort(temp_dir)
        merger(file_name + '_Reordered.pdf', slips)
        # remove directory containing temporary files 
        shutil.rmtree(temp_dir)

    if mode.name == 'BedBath':
    # collect shipping info: this is for BedBath
        order_info, shipping, ship_to = bb_ship()
        #mode.shipping_svc = shipping.iloc[k]

    if mode.name == 'Target':
    # collect shipping info: this is for Target
        order_info, ship_to = tar_ship()
        shipping = ''

    #process shipping data:
    for idx, row in ship_to.iterrows():
        if row.Shipping_Addr == mode.index_val:
            addr_index.append(idx)

    for i, j in enumerate(addr_index):
        if j == addr_index[-1]:
            temp = ship_to.iloc[addr_index[i]:, :]
            temp = temp.Shipping_Addr.astype(str)
        else:
            temp = ship_to.iloc[addr_index[i]:addr_index[i + 1], :]
            temp = temp.Shipping_Addr.astype(str)
        if i != 0:
            temp = temp[1:]
        addr_dict[i] = temp

    for k, v in addr_dict.items():
        q = addr_dict[k]
        addr_dict[k] = strip_Addr(q)

    if mode.op_select == 3:
        clean_labels = Read_Labels()

        # create lists from dict keys: 
        name_list = [k['full_name'] for i, k in clean_labels.items()]
        addr1_list = [k['addr_line1'] for i, k in clean_labels.items()]
        addr2_list = [k['addr_line2'] for i, k in clean_labels.items()]
        addr3_list = [k['addr_line3'] for i, k in clean_labels.items()]

        if len(clean_labels) != total_pages:
            print('\nWARNING: # of Shipping Labels does not equal # Packing Slips.')

        order_dict, no_match_keys = match_Labels(addr_dict, clean_labels)

        check1 = len(no_match_keys)
        check2 = len(order_dict)

        if check1 + check2 != total_pages:
            print(f'\nREAD ERROR: {check1} packing slips and {check2} labels do not total to {total_pages} input file.\n')

        order_list = list(order_dict.values())
        order_dict = {k: v for v,k in enumerate(order_list)}

        # modify numbering format:
        order_dict = page_Mod(order_dict)

        # separate packing slips into individual files
        pdf_splitter(mode.slips_path, order_dict, file_name, temp_dir, no_match_keys)

        # sort packing slips
        slips = slip_Sort(temp_dir)

        # merge packing slips to new PDS:
        merger(file_name + '_Reordered.pdf', slips)

        # merge matchless files into PDF
        matchless = glob.glob('no_match_page_*.pdf')
        if len(matchless) > 0:
            merger(file_name + '_No_Match_Slips.pdf', matchless)

        for file in matchless:
            os.remove(file)
        shutil.rmtree(temp_dir)

    print('')
    print(file_name + ' Reorder Complete!')

    if mode.csvx.lower() == 'y':
        csv_XP(order_info, d, shipping, addr_dict, file_name)
    else:
        print('Process complete!')