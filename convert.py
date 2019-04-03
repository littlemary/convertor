from tkinter import *
from tkinter import filedialog as fd
import os
import xml.dom.minidom
import math
import xml.etree.ElementTree as ET
import codecs

mypath = "C:/AdateMSP/"
maxY = 28500
minY = 5000
maxX = 30900
minX = 6000
L = 300.0
is_obrabot = 0

error_vars = []
errors_xml = {}
writefilename = ''
writefilenamexml = ''
filename=''
filenamexml=''
kols_row = 0

def writetofile(arrfile):
    global writefilename
    global error_vars
    global kols_row
    global is_obrabot
    kols_row = 0
    if writefilename!='':
        try:
            fw = open(writefilename, "w")
        except:
            lbl_message["text"] = u"Ошибка открытия файла на запись"
            lbl_message["bg"] = "red"
            lbl_message["width"] = "50"
            lbl_message["height"] = "3"
            return 0
        for rows, vals in arrfile.items():
            err = 0
            if int(vals['Y']) < minX:
                error_vars.append("Ошибка для ячейки "+vals['P']+": Слишком низкая")
                err = 1
            if int(vals['Y']) > maxX:
                error_vars.append("Ошибка для ячейки " + vals['P'] + ": Слишком высокая")
                err = 1
            if int(vals['X']) < minY:
                error_vars.append("Ошибка для ячейки " + vals['P'] + ": Слишком узкая")
                err = 1
            if int(vals['X']) > maxY:
                error_vars.append("Ошибка для ячейки " + vals['P'] + ": Слишком широкая")
                err = 1
            if is_obrabot == 1:
                barcode = vals['K']+vals['P']
                if barcode in errors_xml.keys():
                    error_vars.append("Ошибка для ячейки " + vals['P'] + ": "+errors_xml[barcode])
                    err = 1
            if (err==0):
                line = 'N' + vals['N'] + 'K' + vals['K'] + 'P' + vals['P'] + 'F' + vals['F'] + 'X' + vals['Y'] + 'Y' + vals['X'] + 'Z' + vals['Z'] + 'C' + vals['C'] + 'S' + vals['S'] + '\n'
                kols_row = kols_row + 1
                fw.write(line)
        fw.close()
    else:
        return 0


def convertxml():
    global filenamexml
    writefilenamexml = mypath + os.path.basename(filenamexml)

    try:
        f = open(filenamexml, 'r')
    except:
        print('error')
    str_f = ''
    for line in f:
        str_f = str_f + line.strip()

    tree = ET.ElementTree(ET.fromstring(str_f))
    # tree = ET.parse(filenamexml)
    root = tree.getroot()
    for c in root:
        for child in c:  # <Fensterdaten
            child.text = ""
            hoehe = child.attrib['Hoehe']
            for cur in child:
                cur.attrib['Laenge'] = hoehe
                for cur1 in cur:
                    cur1.text = ""
                if cur.attrib['TeileNr'] == "1":
                    cur.text = ""
                    for cur_ in cur:
                        new_tag = ET.SubElement(cur_, 'FensterBearb')
                        new_tag.attrib['Wkz'] = '0'
                        new_tag.attrib['BNr'] = '1'  # must be str; cannot be an int
                        new_tag.attrib['xPos'] = '0'
                        new_tag.attrib['yPos'] = '0'
                        new_tag.attrib['zPos'] = '0'
                        new_tag.attrib['WkzPos'] = '0'
                        new_tag = ET.SubElement(cur_, 'FensterBearb')
                        new_tag.attrib['Wkz'] = '0'
                        new_tag.attrib['BNr'] = '1'  # must be str; cannot be an int
                        new_tag.attrib['xPos'] = '640'
                        new_tag.attrib['yPos'] = '0'
                        new_tag.attrib['zPos'] = '0'
                        new_tag.attrib['WkzPos'] = '0'
                if cur.attrib['TeileNr'] == "4":
                    cur.text = ""
                    new_tag = ET.SubElement(cur, 'FensterBearb')
                    new_tag.attrib['Wkz'] = '0'
                    new_tag.attrib['BNr'] = '1'  # must be str; cannot be an int
                    new_tag.attrib['xPos'] = '0'
                    new_tag.attrib['yPos'] = '0'
                    new_tag.attrib['zPos'] = '0'
                    new_tag.attrib['WkzPos'] = '0'
                    new_tag = ET.SubElement(cur, 'FensterBearb')
                    new_tag.attrib['Wkz'] = '0'
                    new_tag.attrib['BNr'] = '1'  # must be str; cannot be an int
                    new_tag.attrib['xPos'] = '640'
                    new_tag.attrib['yPos'] = '0'
                    new_tag.attrib['zPos'] = '0'
                    new_tag.attrib['WkzPos'] = '0'
    strfile = xml.dom.minidom.parseString(
        ET.tostring(
            tree.getroot(),
            'UTF-8')).toprettyxml('  ')
    strfile = strfile.strip()
    strfile = strfile.replace('\n', '\r\n')
    strfile = strfile.replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-8"?>')
    str1 = '<FensterBearb BNr="1" Wkz="0" WkzPos="0" xPos="0" yPos="0" zPos="0"/>'
    str2 = '<FensterBearb Wkz="0" BNr="1" XPos="0" YPos="0" ZPos="0" WkzPos="0"/>'
    strfile = strfile.replace(str1, str2)
    str3 = '<FensterBearb BNr="1" Wkz="0" WkzPos="0" xPos="640" yPos="0" zPos="0"/>'
    str4 = '<FensterBearb Wkz="0" BNr="1" XPos="640" YPos="0" ZfPos="0" WkzPos="0"/>'
    strfile = strfile.replace(str3, str4)
    try:
        fichierTemp = codecs.open(writefilenamexml, "w", encoding="UTF-8", errors="ignore")
        fichierTemp.write(strfile)
        fichierTemp.close()
    except:
        print("Ошибка открытия файла")


def myimportxml():
    global filenamexml
    global errors_xml
    filenamexml = fd.askopenfilename(filetypes=(("XML files", "*.xml"), ("all files", " *.*")
                                         ))
    lbl_filexml["text"] = filenamexml
    if filenamexml:
        try:
            doc = xml.dom.minidom.parse(filenamexml)
        except:
            lbl_message["text"] = u"Ошибка парсера xml"
            lbl_message["bg"] = "red"
            lbl_message["width"] = "100"
            lbl_message["height"] = "3"
            return 0
    staffs = doc.getElementsByTagName("Fensterdaten")
    for cur in staffs:
        barcode = cur.getAttribute('Barcode')
        items = cur.getElementsByTagName("FensterTeiledaten")
        for cur1 in items:
            teileNr = int(cur1.getAttribute('TeileNr'))
            Laenge = float(cur1.getAttribute('Laenge'))
            items1 = cur1.getElementsByTagName("FensterBearb")
            bnr1 = 0
            bnr2 = 0
            for cur2 in items1: #проходимся по FensterBearb и смотрим BNr = 43
                bnr = cur2.getAttribute('BNr')
                xpos= float(cur2.getAttribute('XPos'))
                if bnr=="43":
                    if bnr1==0:
                        bnr1 = xpos
                    else:
                        bnr2 = xpos
                    if (teileNr==1):
                      update = {barcode: "Не может быть Bnr=43 в Teile 1"}
                      errors_xml.update(update)
                    elif (teileNr==4):
                        update = {barcode: "Не может быть Bnr=43 в Teile 4"}
                        errors_xml.update(update)
                    else:
                        if xpos < L:
                            update = {barcode: "xpos "+str(xpos)+" < Laenge"}
                            errors_xml.update(update)
                        res1 = math.fabs(Laenge - xpos)
                        if res1 < L:
                            update = {barcode: "Laenge("+str(Laenge)+") - xpos("+str(xpos)+") < L"}
                            errors_xml.update(update)
                        if bnr2!=0:
                            res2 = math.fabs(bnr1 - bnr2)
                            if res2 < L:
                                update = {barcode: "xpos1 (" +str(bnr1)+ ") - xpos2 (" +str(bnr2)+") < L"}
                                errors_xml.update(update)
    convertxml()
    return 1

def myimporttxt():
    global filename
    filename = fd.askopenfilename(filetypes=(("TXT files", "*.txt"), ("all files", " *.*")
                                         ))
    lbl_filetxt["text"] = filename

def myconvert():
    global writefilename
    global filename
    global filenamexml
    global is_obrabot
    var1 = c_obrvar.get() # галочка "варить с обработкой"
    if var1==True:
        is_obrabot = 1
    else:
        is_obrabot = 0
    if is_obrabot==1 and filenamexml=='':
        lbl_message["text"] = u"Выберите XML файл для конвертации"
        lbl_message["bg"] = "red"
        lbl_message["width"] = "50"
        lbl_message["height"] = "3"
        return 0
    if filename=='':
        lbl_message["text"] = u"Выберите файл для конвертации"
        lbl_message["bg"] = "red"
        lbl_message["width"] = "50"
        lbl_message["height"] = "3"
        return 0

    if filename:
        try:
            f = open(filename, 'r')
        except:
            lbl_message["text"] = u"Ошибка открытия текстового файла"
            lbl_message["bg"] = "red"
            lbl_message["width"] = "50"
            lbl_message["height"] = "3"
            return 0
    else:
        return 0
    writefilename = mypath + os.path.basename(filename)
    kol=0
    arr_rows = {}
    rows = {}
    for line in f:
        str_f = line
        update = {kol:{'N': str_f[1: 4],
                       'K': str_f[5: 15],
                       'P': str_f[16: 21],
                       'F': str_f[22: 25],
                       'X': str_f[26: 31],
                       'Y': str_f[32: 37],
                       'Z': str_f[38: 43],
                       'C': str_f[44: 64],
                       'S': str_f[65: 68]}}
        arr_rows.update(update)
        kol = kol+1
    f.close()
    error_vars.clear()
    writetofile(arr_rows)
    kol1 = str(kols_row)
    resultmessage = "Файл сконвертирован. Обработано "+kol1+" строк"
    lbl_message["text"] = resultmessage
    lbl_message["bg"] = "lightgreen"
    lbl_message["width"] = "50"
    lbl_message["height"] = "3"
    frame2_c = Frame(root, width="70")
    for widget in frame2_c.winfo_children():
        widget.destroy()

    frame2_c.grid(row=10, columnspan=2, sticky=NW,  padx=3, pady=2)

    # Add a canvas in that frame.
    canvas = Canvas(frame2_c, bg="Yellow")
    for widget in canvas.winfo_children():
        widget.destroy()

    canvas.grid(row=0, column=0 )

    # Create a vertical scrollbar linked to the canvas.
    vsbar = Scrollbar(frame2_c, orient=VERTICAL, command=canvas.yview)
    vsbar.grid(row=0, column=1, sticky=NS)
    canvas.configure(yscrollcommand=vsbar.set)

    # Create a horizontal scrollbar linked to the canvas.
    hsbar = Scrollbar(frame2_c, orient=HORIZONTAL, command=canvas.xview)
    hsbar.grid(row=1, column=0, sticky=EW)
    canvas.configure(xscrollcommand=hsbar.set)

    # Create a frame on the canvas to contain the buttons.
    frame2 = Frame(canvas, bg="grey", bd=2)
    for widget in frame2.winfo_children():
        widget.destroy()

    i = 0
    for i1 in error_vars:
        i = i+1
        lbl = Label(frame2, width="70", text=i1, font=("Tahoma", 10), bg="lightgrey", padx=10, pady=5)
        lbl.grid(row=i, columnspan=2, padx=1, pady=1)
    canvas.create_window((0, 0), window=frame2, anchor=NW)

    frame2.update_idletasks()  # Needed to make bbox info available.
    bbox = canvas.bbox(ALL)  # Get bounding box of canvas with Buttons.
    # print('canvas.bbox(tk.ALL): {}'.format(bbox))
    LABEL_BG = "#ccc"  # Light gray.
    ROWS, COLS = 13, 9  # Size of grid.
    ROWS_DISP = 10  # Number of rows to display.
    COLS_DISP = 9  # Number of columns to display.

    # Define the scrollable region as entire canvas with only the desired
    # number of rows and columns displayed.
    w, h = bbox[2] - bbox[1], bbox[3] - bbox[1]
    dw, dh = int((w / COLS) * COLS_DISP), int((h / ROWS) * ROWS_DISP)
    canvas.configure(scrollregion=bbox, width=dw, height=dh)

    return 1

root = Tk()
mycolor2 = '#e0e0e0'

lbl_head = Label(root, text=u"Конвертор", width="55", font=("Tahoma", 15), bg="yellow")
lbl_message = Label(root, text="", font=("Tahoma", 12), width="55", height="0", bg="lightgrey")
lbl_message.grid(row=1, columnspan=4)

but_import = Button(root,
           text= u"Загрузить",
           width=20, height=1,
           font=("Tahoma", 12),
           bg="orange", command=myimporttxt
                  )
lbl_filetxt = Label(root, text="Текстовый файл для конвертации", font=("Tahoma", 12), width="40", height="2", bg="lightgrey")
lbl_filexml = Label(root, text="xml файл для обработки", font=("Tahoma", 12), width="40", height="2", bg="lightgrey")

c_obrvar = BooleanVar()
c = Checkbutton(root, text="Варить с обработкой", variable=c_obrvar, width="50", bg="lightgrey")
but_importx = Button(root,
           text= u"Загрузить",
           width=20, height=1,
           font=("Tahoma", 12),
           bg="orange", command=myimportxml
                  )

but_convert = Button(root,
           text= u"Конвертировать",
           width=20, height=1,
           font=("Tahoma", 12),
           bg="orange", command=myconvert
                  )

lbl_head.grid(row=0, columnspan=2)
lbl_message.grid(row=1, columnspan=2)
but_import.grid (row=2, column=1)
lbl_filetxt.grid(row=2, column=0)

but_importx.grid (row=4, column=1)
lbl_filexml.grid(row=4, column=0)
c.grid(row = 5, column=0)
but_convert.grid (row=5, column=1)



root.title(u"Конвертор")
root.geometry('640x700')
root.configure(bg="lightgrey")
root.mainloop()
