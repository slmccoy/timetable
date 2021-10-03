# import
import pandas as pd
import numpy as np

# imports for final excel output
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment

# set display options
pd.set_option('display.max_rows', 8)

# list/dictionaries
sl = []
pl = ['TP4','TP5','TP6','TP7']
#pl = ['HP4','HP5','HP6','HP7']
ld = {}
tsd = {}
std = {}
fcsd = {}
sfcd = {}

# read in data frames
tf = pd.read_csv('C:/Users/sarah/PycharmProjects/portfolio/timetable/teachers.csv',index_col='initials')
smtf = pd.read_csv('C:/Users/sarah/PycharmProjects/portfolio/timetable/smt.csv',index_col='initials')
ttf = pd.read_csv('C:/Users/sarah/PycharmProjects/portfolio/timetable/tea.csv',index_col='initials')
stf = pd.read_csv('C:/Users/sarah/PycharmProjects/portfolio/timetable/subj.csv',index_col='Subject Code')
rtf = pd.read_csv('C:/Users/sarah/PycharmProjects/portfolio/timetable/room.csv',index_col='room')

def lesson_dic(l):
    c = str(l).split("\n")[0]  # class info without room or teacher (LAC/Cs or 11M/Ma5)

    if len(c.split('/')) == 2:  # Check class info can be split

        # input and output yrs
        yli = ['5', '6', '7', '8', '9', '10', '11', 'L', 'U']
        ylo = ['5', '6', '7', '8', '9', '10', '11', '12', '13']

        for i, y in enumerate(yli):
            if c.find(y) == 0:
                yr = ylo[i]

                # Add lesson to dictionary
                s = c.split('/')[1]
                ld[l] = [yr, s[:2]]

                # Add set if applicable
                if len(s) == 3:
                    ld[l].append(f'Set {s[2]}')
                break


# Teacher timetable remove unnecessary periods and Games and create lesson dictionary
for p in list(ttf):
    if p not in pl:
        ttf.drop(p, axis=1, inplace=True)

    # Remove Games and make lesson dictionary for export
    else:
        for t in ttf.index.values:
            l = ttf.at[t,p]
            c = str(l).split("\n")[0] # class info without room (LAC/Cs or 11M/Ma5)

            if 'Games' in c: # Remove Games and Off Games
                ttf.at[t,p]= np.NaN
            elif 'Meeting' in c: # Remove Games and Off Games
                ttf.at[t,p]= np.NaN
            elif 'Part Time' in c: # Remove Games and Off Games
                ttf.at[t,p]= np.NaN

            lesson_dic(l)

            if l not in ld.keys(): # final check for annomalies in data
                ttf.at[t, p] = np.NaN

# create list of teachers for cover
for t in tf.index.values:
    if t not in smtf.index.values:
        sl.append(t)

# create dictionary teachers name to subjects taught
for t in ttf.index.values: # Add all teachers first
    tsd[t]=[]

for s in stf.index.values:
    if s not in ['Ga', 'Tk', 'Ps','Ls']:
        std[s] = []
        for t in list(stf)[3:]:

            itl = stf.at[s, t]

            if itl is not np.NaN and pd.notna(itl): # remove na
                if str(itl).find(', ') == -1:  # check for multiple entries (applicable in HoDs column)
                    if itl in tsd.keys() and s not in tsd[itl]:
                        tsd[itl].append(s)
                    if itl not in std[s]:
                        std[s].append(itl)
                else:
                    for i in itl.split(', '):
                        if i in tsd.keys() and s not in tsd[i]:
                            tsd[i].append(s)

# Reverse each list as HoDs need to be last to be picked
for itl in std.keys():
    std[itl].reverse()

# Room timetable remove periods and games
for p in list(rtf):
    if p not in pl:
        rtf.drop(p, axis=1, inplace=True)
    else:
        for t in rtf.index.values:
            l = rtf.at[t, p]
            c = str(l).split("\n")[0]  # class info without room (LAC/Cs or 11M/Ma5)

            if 'Games' in c:  # Remove Games and Off Games
                rtf.at[t, p] = 'NaN'

            lesson_dic(l)

            if l not in ld.keys(): # final check for annomalies in data
                rtf.at[t, p] = np.NaN


# Define faculties to subject dictionaries
fcsd["Languages"] = ["EnL", "En", "Sp", "Fr", "Gm", "Gr", "In", "La", "Ru", "Cl", "La", "Ja", "Cn"]
fcsd["Business/Economics"] = ["Ec", "Bs", "Bm", "BsB"]
fcsd["Philosophy/Psychology"] = ["Py", "Pp", "Po", "So", "Ps", "Tk", "Dp"]
fcsd["Geography/History"] = ["Hi", "Gg", "Gp"]
fcsd["Science/Maths"] = ["Ma", "Ph", "Ch", "Bi", "Dt", "As", "Cs", "IT", "Sc"]
fcsd["Arts"]=["Ar", "Dr", "Mu"]

for f in fcsd.keys():
    for s in fcsd[f]:
        if s not in sfcd.keys():
            sfcd[s]=f

# free teachers/room for a period of interest, period: 0 = 4/6, 1 = 5/7
def free(tf,p):
    fl = []
    cl = []
    for r in tf.index.values:
        fst = tf.at[r,pl[p]]
        snd = tf.at[r,pl[p+2]]
        if fst is np.NaN and snd is np.NaN:
            fl.append(r)
        elif fst is not np.NaN and snd is not np.NaN:

            cl.append([r,fst,snd])
    return(fl,cl)


# return both free and clashing teachers and staff across both pairs of periods
ft0l, ct0l = free(ttf,0)
fr0l, cr0l = free(rtf,0)
ft1l, ct1l = free(ttf,1)
fr1l, cr1l = free(rtf,1)


def teacher_collapse(fl,cl):
    res = {}
    for r in cl:
        t = r[0]
        fst = r[1] # stay and cover
        snd = r[2]  # move with teacher

        s = ld[fst][1]

        for ct in std[s]: # for pot cover teacher within subject teachers
            if ct in fl: # if the teacher is free
                res[t]=[ct,fst,snd] # assign cover
                fl.remove(ct)
                break
    return res, fl


def room_collapse(fl,cl):
    res = {}
    for l in cl:
        r = l[0]
        fst = l[1] # stay and cover
        snd = l[2]  # move with teacher

        res[r]=[fl[0],fst,snd]
        fl.remove(fl[0])

    return res, fl


rest0, ft0 = teacher_collapse(ft0l,ct0l)
rest1, ft1 = teacher_collapse(ft1l,ct1l)
resr0, fr0 = room_collapse(fr0l,cr0l)
resr1, fr1 = room_collapse(fr1l,cr1l)

# Export
wb = Workbook()
file = 'C:/Users/sarah/PycharmProjects/portfolio/timetable/tt_collapse.xlsx'
wb.save(file)

def exp_t(file, title, period,res,f):
    wb = load_workbook(file)
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False

    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    ws.column_dimensions["A"].width = 5
    ws.cell(2, 2).value = "Teacher"
    ws.column_dimensions["B"].width = 10
    ws.cell(2, 3).value = "P" + str(period) + " Lesson"
    ws.column_dimensions["C"].width = 10
    ws.cell(2, 4).value = "P" + str(period + 2) + " Lesson"
    ws.column_dimensions["D"].width = 10
    ws.cell(2, 5).value = "Resolution"
    ws.column_dimensions["E"].width = 30
    ws.cell(2, 6).value = "Cover Teacher"
    ws.column_dimensions["F"].width = 10
    ws.cell(2, 8).value = "Free Teachers"
    ws.column_dimensions["H"].width = 10

    for r,t in enumerate(res.keys()):
        fstl = ld[res[t][1]]
        fstrm = res[t][1].split('\n')
        if len(fstrm) == 2:
            fstrm = fstrm[1]
        else:
            fstrm = 'TBC'
        sndl = ld[res[t][2]]

        ws.cell(3+r,2).value = t
        ws.cell(3+r,3).value = f'Yr{fstl[0]} - {fstl[1]}'
        ws.cell(3 + r,4).value = f'Yr{sndl[0]} - {sndl[1]}'
        ws.cell(3 + r, 5).value = f'Yr{sndl[0]} - {sndl[1]} to be covered in {fstrm}'
        ws.cell(3 + r, 6).value = res[t][0]

    for r,t in enumerate(f):
        ws.cell(3+r, 8).value = t

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')

    for row in range(2,ws.max_row+1):
        for column in [2,3,4,5,6,8]:
            ws.cell(row=row, column=column).border = thin_border

    wb.save(file)


exp_t(file, 'Teacher 4-6', 4,rest0,ft0)
exp_t(file, 'Teacher 5-7', 5,rest1,ft1)

def exp_r(file, title, period,res,f):
    wb = load_workbook(file)
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False

    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    ws.column_dimensions["A"].width = 5
    ws.cell(2, 2).value = "Room"
    ws.column_dimensions["B"].width = 10
    ws.cell(2, 3).value = "P" + str(period) + " Lesson"
    ws.column_dimensions["C"].width = 10
    ws.cell(2, 4).value = "P" + str(period + 2) + " Lesson"
    ws.column_dimensions["D"].width = 10
    ws.cell(2, 5).value = "Resolution"
    ws.column_dimensions["E"].width = 30
    ws.cell(2, 6).value = "New Room"
    ws.column_dimensions["F"].width = 10
    ws.cell(2, 8).value = "Free Room"
    ws.column_dimensions["H"].width = 10

    for r,t in enumerate(res.keys()):
        fstl = ld[res[t][1]]
        fstt = res[t][1].split('\n')[1]
        sndl = ld[res[t][2]]

        ws.cell(3+r,2).value = t
        ws.cell(3+r,3).value = f'Yr{fstl[0]} - {fstl[1]}'
        ws.cell(3 + r,4).value = f'Yr{sndl[0]} - {sndl[1]}'
        ws.cell(3 + r, 5).value = f'Yr{fstl[0]} - {fstl[1]} with {fstt} to move rooms'
        ws.cell(3 + r, 6).value = res[t][0]

    for r,t in enumerate(f):
        ws.cell(3+r, 8).value = t

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')

    for row in range(2,ws.max_row+1):
        for column in [2,3,4,5,6,8]:
            ws.cell(row=row, column=column).border = thin_border

    wb.save(file)


exp_r(file, 'Room 4-6', 4,resr0,fr0)
exp_r(file, 'Room 5-7', 5,resr1,fr1)

wb = load_workbook(file)
ws = wb["Sheet"]
wb.remove(ws)
wb.save(file)