# Converting excel file to CAN configuration database file used in automotive for signal logging.
# Excel file format is specified below and can be referred in the code.
# The library used is canlib and Kvaser tools.


# -*- coding: cp1252 -*-
#import sys
#sys.path.append("C:\\Program Files\\Kvaser\\Canlib\\bin_x64\\")
from canlib import kvadblib
import os.path
import xlrd

# Create a new database
db = kvadblib.Dbc(name='Histogram' )

path = "path_of_your_file"

wb = xlrd.open_workbook(path)
sheet = wb.sheet_by_index(0) 

nr = sheet.nrows

nc = sheet.ncols
i = 4

k = 13
msg_row = 4
count = 0
count_st = 0
nr_chg = 1
total = 0
mg_ct = 1


print ("Creating DBC")
while(i<nr):
    
    total = 0
    count = 0
   # for msg_ct in range(i,nr - 1,1):
    for msg_ct in range(i,nr - 1,1):   
        
        if((sheet.cell_value(i,7) == sheet.cell_value(msg_ct + 1,7)) or (sheet.cell_value(msg_ct + 1,7) == "")):
           
           count = count + 1
        else:
           break
    count = count + 1
    
       
    for curr_col in range(7,10,1):
        if(curr_col == 7):
            name_mess = sheet.cell_value(i,curr_col)
            name_mess = name_mess.strip()
            
            
        elif(curr_col==8):

            name_id   = sheet.cell_value(i,curr_col)
            name_id = int(name_id[2:],16)
            name_id = name_id%1000
            print ("name id is ",name_id)

        else:
            dlc_val = sheet.cell_value(i,curr_col)
            
           
            
    print ("count of messga eis ",count);
    message = db.new_message(name = name_mess ,
			 id = int(name_id),
			 dlc=int(dlc_val))
    
    for msg_row in range(i,i+count,1):
        
    
        for curr_col_s in range(k,33,1):
            
            if(curr_col_s == 13):
                name_signal = sheet.cell_value(msg_row,curr_col_s)
                name_signal = name_signal.strip()
                print (name_signal)
            elif(curr_col_s==15):
                name_start_bit = sheet.cell_value(msg_row,curr_col_s)
                print (name_start_bit)
                
            elif(curr_col_s==14): 
                name_length =  sheet.cell_value(msg_row,curr_col_s)
                print (name_length)
                
            elif(curr_col_s==26): 
                name_valuetype = sheet.cell_value(msg_row,curr_col_s)
                name_valuetype = name_valuetype.strip()
                print (name_valuetype)
            elif(curr_col_s==27): 
                name_init_val = sheet.cell_value(msg_row,curr_col_s)
                print (name_init_val)
                if(name_init_val[1] != 'x'):
                        name_init_val = 0;
                
            elif(curr_col_s==28):
                name_factor = sheet.cell_value(msg_row,curr_col_s)
                print (name_factor)
    
            elif(curr_col_s==29):
                name_offset = sheet.cell_value(msg_row,curr_col_s)
                print (name_offset)
                
            elif(curr_col_s==30):
                name_minimum = sheet.cell_value(msg_row,curr_col_s)
                print (name_minimum)
                
            elif(curr_col_s==31):
                name_maximum = sheet.cell_value(msg_row,curr_col_s)
                print (name_maximum)

            elif(curr_col_s==32):
                name_unit = sheet.cell_value(msg_row,curr_col_s)
                print (name_unit)
                
                
        _type = kvadblib.SignalType.SIGNED
        _type1 = kvadblib.SignalType.UNSIGNED
       
       
       
        if(name_valuetype == "Signed"):
            
            message.new_signal(name=name_signal ,
                           type=kvadblib.SignalType.SIGNED,
                           byte_order=kvadblib.SignalByteOrder .INTEL, # default
                           mode=kvadblib.SignalMultiplexMode .MUX_INDEPENDENT, # default
                           size=kvadblib.ValueSize( startbit=int(name_start_bit), length=int(name_length)) ,
                           scaling=kvadblib.ValueScaling( factor=float(name_factor), offset=float(name_offset)) ,
                           limits=kvadblib . ValueLimits(min=float(name_minimum), max=float(name_maximum)) ,
                           )
            
           
        else:
           message.new_signal(name=name_signal ,
                           type=kvadblib.SignalType.UNSIGNED,
                           byte_order=kvadblib.SignalByteOrder .INTEL, # default
                           mode=kvadblib.SignalMultiplexMode .MUX_INDEPENDENT, # default
                           size=kvadblib.ValueSize( startbit=int(name_start_bit), length=int(name_length)) ,
                           scaling=kvadblib.ValueScaling( factor=float(name_factor), offset=float(name_offset)) ,
                           limits=kvadblib.ValueLimits(min=float(name_minimum), max=float(name_maximum)) ,
                           )
                           
        
            
        
        k = 13
    i = i + count
    
    
    
    
                                              

db.write_file("db_histogram.dbc")

db.close()
del i 

del k 
del msg_row
del count
del count_st
del nr_chg
del total
del mg_ct 
