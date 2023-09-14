import win32com.client
import gml   # email sending
import gph

# open Excel app
excel_app = win32com.client.Dispatch("Excel.Application")
wb = excel_app.Workbooks.Open("C:\\Users\\Sakashita\\sakawork\\Scarpe_Pushlog\\report\\sample4.xlsx")

data_sheet_name = "明和合成（株）SCS-C-CS-20_1分周期"  # data sheet
target_sheet_name = "R3"
target_cell  = "F5"
target_cell_avg = "F14"
target_cell_max = "H14"
target_cell_min = "J14"

################## #proc-1   alarm check 
# choose  data sheet
data_sheet = wb.Sheets(data_sheet_name)

# set data range
data_range = data_sheet.Range("S2:S1442")


# init alarm count
alarm_count = 0

# check alarm count
for cell in data_range:
    if cell.Value != "異常なし":  # 異常なしでない場合
        alarm_count += 1

print("alarm_count : ", alarm_count)
print("-----------------------")

# write to excel cell 
target_sheet = wb.Sheets(target_sheet_name)
target_sheet.Range(target_cell).Value = alarm_count


################ #proc-2 SouBai temperture avg, max, min
temperature_range = data_sheet.Range("E2:E1442")
valid_values = []

for cell in temperature_range:
    try:
        value = float(cell.Value)
        #print ("value is :", value)
        valid_values.append(value)
    except (ValueError, TypeError):
        pass # ignore invalid values


if valid_values:
    avg_temperature =  round(sum(valid_values) / len(valid_values) , 1)
    max_temperature = max(valid_values)
    min_temperature = min(valid_values)
else:
    avg_temperature = max_temperature = min_temperature = None

# writing each values to cell
#target_sheet = wb.Sheets(target_sheet_name)


target_sheet.Range(target_cell_avg).Value = avg_temperature
target_sheet.Range(target_cell_max).Value = max_temperature
target_sheet.Range(target_cell_min).Value = min_temperature
print("ave temp:", avg_temperature)
print("max temp:", max_temperature)
print("min temp:", min_temperature)
print("-----------------------")

################ #proc-3 SouBai Atsu-ryoku avg, max, min
p_range = data_sheet.Range("G2:G1442")
v_arry = []

for cell in p_range:
    try:
        pv = float(cell.Value)
        v_arry.append(pv)
    except :
        pass

if v_arry:
    avg_p = round( sum(v_arry) / len(v_arry), 3)
    max_p = round(max(v_arry) , 3 )
    min_p = round(min(v_arry) , 3 )
else:
    avg_p = max_p = min_p = None


print("ave pressr:", avg_p)
print("max pressr:", max_p)
print("min pressr:", min_p)
print("-----------------------")

target_sheet.Range("F30").Value = avg_p
target_sheet.Range("H30").Value = max_p
target_sheet.Range("K30").Value = min_p

################ #proc-4 外気温度と湿度 (outtemp and humidity)
#outside temp
o_range = data_sheet.Range("O2:O1442")
o_arry = []

for cell in o_range:
    try:
        ov = float(cell.Value)
        o_arry.append(ov)
    except :
        pass

if v_arry:
    avg_o = round( sum(o_arry) / len(o_arry), 1)
    max_o = round(max(o_arry) , 1)
    min_o = round(min(o_arry) , 1)

print("ave out temp:", avg_o)
print("max out temp:", max_o)
print("min out temp:", min_o)
print("-----------------------")

target_sheet.Range("F46").Value = avg_o
target_sheet.Range("H46").Value = max_o
target_sheet.Range("K46").Value = min_o

# humidity ---------------------------------------
h_range = data_sheet.Range("P2:P1442")
h_arry = []

for cell in h_range:
    try:
        hv = float(cell.Value)
        h_arry.append(hv)
    except :
        pass


if h_arry:
    avg_h = round( sum(h_arry) / len(h_arry), 1)
    max_h = round(max(h_arry) , 1)
    min_h = round(min(h_arry) , 1)

print("ave humidity:" , avg_h)
print("max humidity:" , max_h)
print("min humidity:" , min_h)
print("-----------------------")
target_sheet.Range("F48").Value = avg_h
target_sheet.Range("H48").Value = max_h
target_sheet.Range("K48").Value = min_h

################ #proc-4   average  稼働率
ur_range = data_sheet.Range("H2:H1442")

ur_arry = []

for cell in ur_range:
    try:
        urv = float(cell.Value)
        ur_arry.append(urv)
    except :
        pass


if ur_arry:
    avg_ur = round( sum(ur_arry) / len(ur_arry), 1)

print("ave utilize rate:" , avg_ur)
print("-----------------------")

##########################################################

# save -> close
wb.Save()

#wb.Close(False)
excel_app.Quit()

##################### mail sending process ##############

print ("excel file update done")
sender_email   = 'sakashita@kannetsu.co.jp'
#receiver_email = 'ueno@kannetsu.co.jp'
receiver_email = 'sakanujin@gmail.com'
username = 'sakashita@kannetsu.co.jp'
password = 'knnts1109'
subject   = 'ファイルの添付メールテスト'

body = 'いつもお世話になっております。メールを添付しました。ご査収ください'
file_path = './sample.xlsx'

gml.send_email(sender_email, receiver_email, username, password, subject, body, file_path)

