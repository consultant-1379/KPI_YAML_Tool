import datetime
import json
import os
import logging
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from tkinter import font
from tkinter.font import BOLD
import openpyxl
import pandas as pd
import ruamel.yaml
from xlsxwriter import *


'''
<----------------------Note:- Before running this script------------------>
Make sure you have Python 3.7.2 installed.
Step 1: Copy paste the below code in an empty file and save it as lib_install.py under H:/downloads. 
Step 2: Open cmd and type cd H:/downloads and hit enter.
Step 3: Now type python lib_install.py and hit enter.

Once the installation is finished, you are good to go.
Now run the Child Tool.py script.

-Or you can manually install the below libraries placed inside req_lib list,
Example:- Open Command Prompt and type pip install tk==0.1.0 and hit enter.

<---------------------- Code to paste in lib_install.py file --------------->
from subprocess import call
req_lib = ['tk==0.1.0', 'ruamel.yaml==0.17.21','pandas==1.3.5','xlsxWriter==3.0.3','openpyxl==3.0.9']
for lib in libraries:
    call(f"python -m pip install {lib} --user", shell=True)

'''

#creating the main/first GUI window
root = Tk()
root.resizable(False, False)
mywidth = 250
myheight = 130
scrwdth = root.winfo_screenwidth()
scrhgt = root.winfo_screenheight()
xLeft = int((scrwdth / 2) - (mywidth / 2)) - 70
yTop = int((scrhgt / 2) - (myheight / 2)) - 30
root.geometry(str(mywidth) + "x" + str(myheight) + "+" + str(xLeft) + "+" + str(yTop))
root.configure(background="#b1cee9")
root.title("YAML Tool")
log_dir=os.curdir
system_time = datetime.datetime.now()
logger_filename_main = log_dir+r"/KPI&Error_log_{}.txt".format(system_time.strftime("%d%h%Y%H%M") ) # "%Y%m%d%H%M")
#print("log_dir :" + log_dir)
#function to create the GUI window for YAMl to JSON conversion
'''def Parse_CHILD_File():
    roots = Toplevel(root)
    roots.resizable(False, False)
    mywidth = 480
    myheight = 150
    scrwdth = roots.winfo_screenwidth()
    scrhgt = roots.winfo_screenheight()
    xLeft = int((scrwdth / 2) - (mywidth / 2)) - 70
    yTop = int((scrhgt / 2) - (myheight / 2)) - 30
    roots.geometry(
        str(mywidth) + "x" + str(myheight) + "+" + str(xLeft) + "+" + str(yTop)
    )
    roots.configure(background="#b1cee9")

    selectfile = Label(
        roots,
        text=" Select YAML Folder ",
        bg="#b1cee9",
        fg="black",
        font=["Helvetica", 10],
    )
    selectfile.grid(row=3, column=1)  
    selectfileentry = Entry(roots, width=40)
    selectfileentry.grid(row=3, column=2)  ''' #EQEV-112782 - For future reference, if needed (Do not delete this code).

    #function to take directory input from user and store all the YAML file names with path into filelist3
    #and Convert YAML files into JSON files
def YAML_to_Json(FilePath):
    #filepath3, filelist3
    filepath3 = FilePath
    #l=[]
        #selectfileentry.insert(0, filepath3)
    filelist3 = []
    for root, dirs, files in os.walk(filepath3):
        for file in files:
            if file.endswith(".yaml"):
                filelist3.append(os.path.join(root, file))

    
    json_file_counts = 0
    #if len(selectfileentry.get()) == 0:
    #    messagebox.showerror("Error", "Please Select CHILD Directory")
    #else:
    logging.basicConfig(filename=logger_filename_main, level=logging.DEBUG, 
                format='\n%(asctime)s-> %(levelname)s: %(message)s\n')
    logger=logging.getLogger(__name__)
    logger.info( "This Logger File for Convert YAML files to JSON files function and it is generated on: "
                + str(system_time.strftime("%d-%h-%Y %H:%M"))
                + " by "
                + os.getlogin())
    try:
        input_files_total=len(filelist3)
        passed_files_count=0
        failed_files_count=0
        #q = len(filelist3)
        #print("Q: "+str(q))
        if input_files_total==0:
            raise Exception('No YAML Files in selected directory')
        s='Files below are having Errors\n'
        r='File below is having Error\n'
        filename=''
        old_f=[]
        yaml = ruamel.yaml.YAML(typ="safe")
        for files in filelist3:
            try:
                with open(files, encoding="utf-8") as i:
                    data = yaml.load(i)
                    f = files.replace(".yaml", "")
                    with open(f + ".json", "w") as o:
                        json.dump(data, o, indent=2)
                        passed_files_count+=1
            except Exception as e:
                filename+='\t'+files+'\n'
                failed_files_count+=1
                continue
        if filename:
            if failed_files_count>1:
                #print(a,s,filename)
                logger.info(str(failed_files_count) + " " + s + filename)
            else:
                #print(a,r,filename)
                logger.info(str(failed_files_count) + r + filename)
        '''for root, dirs, files in os.walk(filepath3):
            for file in files:
                if file.endswith(".json"):
                    json_file_counts += 1'''
        '''messagebox.showinfo("YAML->JSON Success", "    Conversion Completed    "+"\n\n"
                                                 "    Total yaml files : "+ str(input_files_total)+"\n"
                                                 "    Yaml files passed :"+str(passed_files_count)+"\n"
                                                 "    Yaml files failed :"+str(failed_files_count))'''
        return [input_files_total,passed_files_count,failed_files_count]

    except Exception as e:
        logger.error(e)
        messagebox.showerror('Error',f"{e} \n Error Log file has been created under path: {filepath3} ")

    '''def clear2():
        selectfileentry.delete(0, END)

    ttk.Button(roots, text="Browse", command=open_YAML_dir).grid(
        row=3, column=3, pady=10, padx=5
    )
    ttk.Button(roots, text="Convert", command=convert_yaml_to_json).grid(
        row=5, column=2, padx=10, pady=5
    )
    ttk.Button(roots, text="Clear", command=clear2).grid(
        row=10, column=1, padx=10, pady=5
    )
    ttk.Button(
        roots, text="Close", command=lambda: [roots.destroy(), root.deiconify()]
    ).grid(row=10, column=3, padx=20, pady=10)
    roots.title("YAML Tool")
    # roots.iconbitmap(r"H:\Desktop\eric.ico")

    roots.mainloop()'''#EQEV-112782 - For future reference, if needed (Do not delete this code).

#function to create GUI window for New node files to excel
def New_Node_Files_To_Excel():
    root1 = Toplevel(root)
    root1.resizable(False, False)
    mywidth = 550
    myheight = 250
    scrwdth = root1.winfo_screenwidth()
    scrhgt = root1.winfo_screenheight()
    xLeft = int((scrwdth / 2) - (mywidth / 2)) - 70
    yTop = int((scrhgt / 2) - (myheight / 2)) - 30
    root1.geometry(
        str(mywidth) + "x" + str(myheight) + "+" + str(xLeft) + "+" + str(yTop)
    )
    root1.configure(background="#b1cee9")

    selectfile5 = Label(
        root1,
        text="PARSE INPUT TO EXCEL:--",
        bg="#b1cee9",
        fg="black",
        font=["Helvetica", 10, UNDERLINE, BOLD],
    )
    selectfile4 = Label(
        root1,
        text="""Select CHILD Folder""",
        bg="#b1cee9",
        fg="black",
        font=["Helvetica", 10],
    )
    selectfile8 = Label(
        root1,
        text="Save File",
        bg="#b1cee9",
        fg="black",
        font=["Helvetica", 10],
    )
    selectfile5.grid(row=6, column=1)  
    selectfile4.grid(row=7, column=1)  
    selectfile8.grid(row=8, column=1)  
    selectfileentry4 = Entry(root1, width=40)
    selectfileentry8 = Entry(root1, width=40)
    selectfileentry4.grid(row=7, column=2)  
    selectfileentry8.grid(row=8, column=2)  

    #function to take directory input from user and store all json file names into filelist5
    def new_input_only():
        global filepath4
        filepath4 = filedialog.askdirectory(title="Select a Directory")
        selectfileentry4.insert(0, filepath4)

    #function to take directory input from user, where you want to save new node excel file
    def save_new_version_into_excel():
        global fvc
        fvc = filedialog.askdirectory(title="Save Excel File to ")
        selectfileentry8.insert(0, fvc)
        system_time = datetime.datetime.now()
        fvc = fvc + "/KPI_Parsed_Document_{}.xlsx".format(
            system_time.strftime("%d%h%Y%H%M")
        )

    #function to print json files into Excel sheet
    def only_print_into_Excel():
        global fvc
        if not selectfileentry4.get():
            messagebox.showerror("Error", "Please Select New version files")
        elif not selectfileentry8.get():
            messagebox.showerror("Error", "Please Select Excel Save path")
        else:
            l=YAML_to_Json(filepath4)
            filelist5 = []
            for root, dirs, files in os.walk(filepath4):
                for file in files:
                    if file.endswith(".json"):
                        filelist5.append(os.path.join(root, file))
            ss=os.path.dirname(fvc)
            logging.basicConfig(filename=logger_filename_main, level=logging.DEBUG, 
                        format='\n%(asctime)s-> %(levelname)s: %(message)s\n')
            logger=logging.getLogger(__name__)
            try:
                dfs3 = []
                pd.io.formats.excel.ExcelFormatter.header_style = None
                writer0 = pd.ExcelWriter(fvc, engine="xlsxwriter")
                err_files = ""
                failed_files_count=0
                passed_files_count=0
                #count_x=0
                s_x='Input files below are having Errors\n'
                r_x='Input File below is having Error\n'
                total_Json_files=len(filelist5)
                for file in filelist5:
                    with open(file) as f:
                        try:
                            abc = json.load(f)
                            gc = abc["Indicator"]
                            json_data2 = pd.json_normalize(gc, max_level=0)
                            dfs3.append(json_data2)
                            df = pd.concat(dfs3)
                            df.to_excel(writer0, sheet_name="KPI File", index=False)
                            passed_files_count+=1
                        except Exception as e:
                            err_files+='\t'+file+'\n'
                            failed_files_count+=1
                            continue
                if err_files:
                    if failed_files_count>1:
                        logger.info(str(failed_files_count) + " " + s_x + err_files)
                    else:
                        logger.info(str(failed_files_count) + " " + r_x + err_files)
                workbook1 = writer0.book
                worksheet1 = writer0.sheets["KPI File"]
                worksheet1.set_tab_color("#948A54")
                worksheet1.freeze_panes(1, 0)
                worksheet1.set_default_row(15)
                worksheet1.set_column(0, 26, 40)
                header_fmt1 = workbook1.add_format(
                    {"bg_color": "#00B0F0", "bold": True}
                )
                worksheet1.set_row(0, 20, header_fmt1)
                writer0.save()
                '''messagebox.showinfo("JSON(s)->Excel Success", "     File Generated Successfully    "+"\n\n"
                                                 "    Total Json files : "+ str(total_Json_files)+"\n"
                                                 "    Json files passed :"+str(passed_files_count)+"\n"
                                                 "    Json files failed :"+str(failed_files_count))'''
                messagebox.showinfo("JSON(s)->Excel Success", "    Yaml to Json Conversion    "+"\n"
                                                 "    Total Yaml files : "+ str(l[0])+"\n"
                                                 "    Yaml files passed :"+str(l[1])+"\n"
                                                 "    Yaml files failed :"+str(l[2])+"\n\n"
                                                 "    Excel File Generated Successfully    "+"\n"
                                                 "    Total Json files : "+ str(total_Json_files)+"\n"
                                                 "    Json files passed :"+str(passed_files_count)+"\n"
                                                 "    Json files failed :"+str(failed_files_count))
            except Exception as e:
                logger.error(e)
                messagebox.showerror('Error',f"{e} \n Error Log file has been created under path: {ss} ")

    def clear1():
        selectfileentry4.delete(0, END)
        selectfileentry8.delete(0, END)

    ttk.Button(root1, text="Browse", command=new_input_only).grid(
        row=7, column=3, pady=10, padx=5
    )

    ttk.Button(root1, text="Convert", command=only_print_into_Excel).grid(
        row=9, column=2, pady=10, padx=5
    )

    ttk.Button(root1, text="Browse", command=save_new_version_into_excel).grid(
        row=8, column=3, padx=10, pady=5
    )
    ttk.Button(root1, text="Clear", command=clear1).grid(
        row=10, column=1, padx=10, pady=5
    )
    ttk.Button(
        root1, text="Close", command=lambda: [root1.destroy(), root.deiconify()]
    ).grid(row=10, column=3, padx=10, pady=5)
    # root1.iconbitmap(r"H:\Desktop\eric.ico")

    root1.mainloop()

#creating GUI window for Get Delta option
def Get_Delta():
    root2 = Toplevel(root)
    root2.resizable(False, False)
    mywidth = 550
    myheight = 260
    scrwdth = root2.winfo_screenwidth()
    scrhgt = root2.winfo_screenheight()
    xLeft = int((scrwdth / 2) - (mywidth / 2)) - 70
    yTop = int((scrhgt / 2) - (myheight / 2)) - 30
    root2.geometry(
        str(mywidth) + "x" + str(myheight) + "+" + str(xLeft) + "+" + str(yTop)
    )
    # root2.iconbitmap(r"H:\Desktop\eric.ico")

    root2.configure(background="#b1cee9")

    Label(
        root2,
        text="GENERATE DELTA:--",
        bg="#b1cee9",
        fg="black",
        font=["Helvetica", 10, UNDERLINE, BOLD],
    ).grid(row=10, column=1)

    Label(
        root2,
        text="Select Old CHILD Folder",
        bg="#b1cee9",
        fg="black",
        font=["Helvetica", 10],
    ).grid(row=11, column=1)

    Label(
        root2,
        text="Select New CHILD Folder",
        bg="#b1cee9",
        fg="black",
        font=["Helvetica", 10],
    ).grid(row=12, column=1)

    Label(
        root2,
        text="Save Delta File",
        bg="#b1cee9",
        fg="black",
        font=["Helvetica", 10],
    ).grid(row=13, column=1)

    selectfileentry1 = Entry(root2, width=40)
    selectfileentry2 = Entry(root2, width=40)
    selectfileentry3 = Entry(root2, width=40)

    # packing entry fields
    selectfileentry1.grid(row=11, column=2)
    selectfileentry2.grid(row=12, column=2)
    selectfileentry3.grid(row=13, column=2)

    def open_file1():
        '''function to take directory input from user and convert all yaml files into json, selected in old input'''
        global filepath, old_json_files
        filepath = filedialog.askdirectory(title="Select a Directory")
        selectfileentry1.insert(0, filepath)
        
    def open_file2():
        '''function to take directory input from user and convert all yaml files into json, selected in new input'''
        global filepath2, new_json_files
        filepath2 = filedialog.askdirectory(title="Select a Directory")
        selectfileentry2.insert(0, filepath2)
        

    def buttontoexcel():
        '''function to take directory input from user, where user wants to save Delta Excel file'''
        global fvc
        fvc = filedialog.askdirectory(title="Save Delta Excel File to ")
        system_time = datetime.datetime.now()
        fvc = fvc + "/KPI_Delta_Document_{}.xlsx".format(
            system_time.strftime("%d%h%Y%H%M")
        )
        selectfileentry3.insert(0, fvc)
    
    def add2excel():
        '''function to print old and new json files into excel sheet and storing it in buffer for later use'''
        global totold,totnew
        ss=os.path.dirname(fvc)
        logging.basicConfig(filename=logger_filename_main, level=logging.DEBUG, 
                    format='\n%(asctime)s-> %(levelname)s: %(message)s\n')
        logger=logging.getLogger(__name__)
        filename_123=""
        global count_n, count_o
        try:
            if len(selectfileentry1.get()) == 0:
                messagebox.showerror("Error", "Please Select Old Version Files")
            elif len(selectfileentry2.get()) == 0:
                messagebox.showerror("Error", "Please Select New Version Files")
            elif len(selectfileentry3.get()) == 0:
                messagebox.showerror("Error", "Please Select the Output Path for Delta")
            else:
                totold=YAML_to_Json(filepath)
                old_json_files = []
                for root, dirs, files in os.walk(filepath):
                    for file in files:
                        if file.endswith(".json"):
                            old_json_files.append(os.path.join(root, file))
            
                totnew=YAML_to_Json(filepath2)
                new_json_files = []
                for root, dirs, files in os.walk(filepath2):
                    for file in files:
                        if file.endswith(".json"):
                            new_json_files.append(os.path.join(root, file))
                dfs1 = []
                pd.io.formats.excel.ExcelFormatter.header_style = None
                writer = pd.ExcelWriter(fvc, engine="xlsxwriter")
                '''             ********** EQEV-111381 *************        '''
                # Added new try- except blocks to handling the exception.
                error_files_old = ""
                count_o=0
                s_o='Old input files below are having "Indicator" tag Errors\n'
                r_o='Old input File below is having "Indicator" tag Error\n'
                for filename_123 in old_json_files:
                    with open(filename_123) as f:
                        try:
                            ab = json.load(f)
                            gg = ab["Indicator"]
                            json_data = pd.json_normalize(gg, max_level=0)
                            dfs1.append(json_data)
                            df = pd.concat(dfs1)
                            df.to_excel(writer, sheet_name="Old Input", index=False)
                        except Exception as e:
                            error_files_old+='\t'+filename_123+'\n'
                            count_o+=1
                            continue
                if error_files_old:
                    if count_o>1:
                        logger.info(str(count_o) + " " + s_o + error_files_old)
                    else:
                        logger.info(str(count_o) + " " + r_o + error_files_old)


                dfs2 = []
                error_files_new = ""
                count_n=0
                s='New input files below are having "Indicator" tag Errors\n'
                r='New input File below is having "Indicator" tag Error\n'
                for filename_123 in new_json_files:
                    with open(filename_123) as f:
                        try:
                            abc = json.load(f)
                            gc = abc["Indicator"]
                            json_data = pd.json_normalize(gc, max_level=0)
                            dfs2.append(json_data)
                            df = pd.concat(dfs2)
                            df.to_excel(writer, sheet_name="New Input", index=False)
                        except Exception as e:
                            error_files_new+='\t'+filename_123+'\n'
                            count_n+=1
                            continue   
                writer.save()
                if error_files_new:
                    if count_n>1:
                        logger.info(str(count_n) + " " + s + error_files_new)
                    else:
                        logger.info(str(count_n) + " " + r + error_files_new)
            '''             ********** EQEV-111381 end *************        '''
        except Exception as e:
            logger.error(e)
            logger.info("\n\t"+filename_123)
            messagebox.showerror('Error',f"{e} \n Error Log file has been created under path Indicator: {logger_filename_main} "
                                 "\n filename_123"+ filename_123)
                                 #+"\n filename_1234: " +filename_1234)


    def delta_gen():
        '''Getting the data from buffered Excel file and comparing and printing into Delta excel file'''
        ss=os.path.dirname(fvc)
        logging.basicConfig(filename=logger_filename_main, level=logging.DEBUG, 
                    format='\n%(asctime)s-> %(levelname)s: %(message)s\n')
        logger=logging.getLogger(__name__)
        try:
            df_OLD = pd.read_excel(fvc, sheet_name=0, keep_default_na=False)
            df_NEW = pd.read_excel(fvc, sheet_name=1, keep_default_na=False)
            key = ["Name", "NeType", "Category"]

            df_OLD = df_OLD.set_index(key)
            df_NEW = df_NEW.set_index(key)

            writer2 = pd.ExcelWriter(fvc, engine="xlsxwriter")

            added_col = list(set(df_NEW.columns) - set(df_OLD.columns))
            removed_col = list(set(df_OLD.columns) - set(df_NEW.columns))
            dfDiff = df_NEW.copy()
            df = df_OLD.copy()
            droppedRows = []
            newRows = []
            deprecate_rowso = []
            deprecate_rowsn = []
            O_rows = []
            cols_OLD = df_OLD.columns
            cols_NEW = df_NEW.columns
            sharedCols = list(set(cols_OLD).intersection(cols_NEW))
            dfDiff.insert(0, "Comment", "")
            df.insert(0, "Comment", "")

            for row2 in df_OLD.index:
                if df_OLD.loc[row2, "FormulaStatus"].lower()== "deprecated":
                    deprecate_rowso.append(row2)

            for row3 in df_OLD.index:
                if df_OLD.loc[row3, "FormulaStatus"].lower() == "obsolete":
                    O_rows.append(row3)

            for row4 in df_NEW.index:
                if df_NEW.loc[row4, "FormulaStatus"].lower() == "deprecated":
                    deprecate_rowsn.append(row4)
            for row5 in df_NEW.index:
                if df_NEW.loc[row5, "FormulaStatus"].lower() == "obsolete":
                    O_rows.append(row5)
            sx = []

            #actual comparison starts from here
            for row in dfDiff.index:
                s = ""
                if (row in df_OLD.index) and (row in df_NEW.index):
                    for col in sharedCols:
                        value_OLD = df_OLD.loc[row, col]
                        value_NEW = df_NEW.loc[row, col]

                        if value_OLD == value_NEW:
                            dfDiff.loc[row, col] = df_NEW.loc[row, col]
                        else:
                            sx.append(col)
                    if len(sx) != 0:
                        s += "Modified Columns:- {} ".format([c for c in sx])
                    sx.clear()

                    for x in added_col:
                        if len(dfDiff.loc[row, x]) != 0:
                            sx.append(x)
                    if len(sx) != 0:
                        s += "Added Columns:- {} ".format([c for c in sx])
                    sx.clear()
                    for x in removed_col:
                        if len(df_OLD.loc[row, x]) != 0:
                            sx.append(x)
                    if len(sx) != 0:
                        s += "Removed Columns:- {} ".format([c for c in sx])
                    sx.clear()
                    dfDiff.loc[row, "Comment"] = s
                    if dfDiff.loc[row, "Comment"] == "":
                        dfDiff.drop(row, inplace=True)
                else:
                    dfDiff.loc[row, "Comment"] = "New"
                    newRows.append(row)

            for row in df_OLD.index:
                if row not in df_NEW.index:
                    df.loc[row, "Comment"] = "Removed"
                    droppedRows.append(row)

                if df.loc[row, "Comment"] != "Removed":
                    df.drop(row, inplace=True)

            nan_value = float("NaN")
            dataframes = [df, dfDiff, df_OLD, df_NEW]
            for v in dataframes:
                v.replace(0, nan_value, inplace=True)
                v.replace("", nan_value, inplace=True)
                v.dropna(how="all", axis=1, inplace=True)

            if len(df.index) != 0:
                df.reset_index()
                df.to_excel(writer2, sheet_name="Delta_Rm", index=True)
                workbook1 = writer2.book
                worksheet1 = writer2.sheets["Delta_Rm"]
                worksheet1.set_tab_color("#948A54")
                worksheet1.freeze_panes(1, 0)
                worksheet1.set_default_row(15)
                worksheet1.set_column(0, 26, 40)
                header_fmt1 = workbook1.add_format(
                    {"bg_color": "#00B0F0", "bold": True}
                )
                worksheet1.set_row(0, 20, header_fmt1)
                removed_format = workbook1.add_format({"bg_color": "#FFBF00"})
                for rownum1 in range(df.shape[0]):
                    worksheet1.set_row(rownum1 + 1, 15, removed_format)
                    
            else:
                my_df = pd.DataFrame()
                my_df.to_excel(writer2, sheet_name="Delta_Rm")

            if len(dfDiff.index) != 0:
                dfDiff.to_excel(writer2, sheet_name="Delta", index=True)
                workbook = writer2.book
                worksheet = writer2.sheets["Delta"]
                worksheet.set_tab_color("#948A54")
                worksheet.freeze_panes(1, 0)
                worksheet.set_default_row(15)
                header_fmt = workbook.add_format({"bg_color": "#00B0F0", "bold": True})
                modified_format = workbook.add_format({"bg_color": "#ffe873"})
                added_row_format = workbook.add_format({"bg_color": "#8db600"})
                deprecated_fmt = workbook.add_format({"bg_color": "#ff6600"})
                obsolete_fmt = workbook.add_format({"bg_color": "#ff00ff"})
                worksheet.set_row(0, 20, header_fmt)
                worksheet.set_column(0, 26, 40)
                for rownum in range(dfDiff.shape[0]):
                    worksheet.set_row(rownum + 1, 15, modified_format)
                    row = dfDiff.index[rownum]
                    if dfDiff.loc[row, "FormulaStatus"].lower() == "deprecated":
                        worksheet.set_row(rownum + 1, 15, deprecated_fmt)
                    if dfDiff.loc[row, "FormulaStatus"].lower() == "obsolete":
                        worksheet.set_row(rownum + 1, 15, obsolete_fmt)
                    for x in newRows:
                        if x == row:
                            worksheet.set_row(rownum + 1, 15, added_row_format)
            else:
                my_df2 = pd.DataFrame()
                my_df2.to_excel(writer2, sheet_name="Delta")

            df_OLD.to_excel(writer2, sheet_name="Old_Input", index=True)
            df_NEW.to_excel(writer2, sheet_name="New_Input", index=True)

            workbook2 = writer2.book
            worksheet2 = writer2.sheets["Old_Input"]
            worksheet2.set_tab_color("#948A54")
            worksheet2.freeze_panes(1, 0)

            workbook3 = writer2.book
            worksheet3 = writer2.sheets["New_Input"]
            worksheet3.set_tab_color("#948A54")
            worksheet3.freeze_panes(1, 0)

            header_fmt2 = workbook2.add_format({"bg_color": "#00B0F0", "bold": True})
            header_fmt3 = workbook3.add_format({"bg_color": "#00B0F0", "bold": True})
            deprecated_format1 = workbook2.add_format({"bg_color": "#ff6600"})
            obsolete_format1 = workbook2.add_format({"bg_color": "#ff00ff"})
            deprecated_format = workbook3.add_format({"bg_color": "#ff6600"})
            obsolete_format = workbook3.add_format({"bg_color": "#ff00ff"})
            worksheet2.set_column(0, 26, 40)
            worksheet3.set_column(0, 26, 40)
            worksheet2.set_row(0, 20, header_fmt2)
            worksheet3.set_row(0, 20, header_fmt3)

            for D_row in range(df_OLD.shape[0]):
                rowd = df_OLD.index[D_row]
                for a in deprecate_rowso:
                    if a == rowd:
                        worksheet2.set_row(D_row + 1, 15, deprecated_format1)
                for b in O_rows:
                    if b == rowd:
                        worksheet2.set_row(D_row + 1, 15, obsolete_format1)

            for a_row in range(df_NEW.shape[0]):
                rowo = df_NEW.index[a_row]
                for a in deprecate_rowsn:
                    if a == rowo:
                        worksheet3.set_row(a_row + 1, 15, deprecated_format)
                for b in O_rows:
                    if b == rowo:
                        worksheet3.set_row(a_row + 1, 15, obsolete_format)

            l=['New KPIs','Deprecated KPIs','Removed KPIs','Obsolete KPIs','Modified KPIs']
            df=pd.DataFrame()
            df.to_excel(writer2, sheet_name="Color Summary", index=False)
            workbook = writer2.book
            worksheetc = writer2.sheets["Color Summary"]
            removed_format = workbook.add_format({"bg_color": "#FFBF00"})
            modified_format = workbook.add_format({"bg_color": "#ffe873"})
            added_row_format = workbook.add_format({"bg_color": "#8db600"})
            deprecated_fmt = workbook.add_format({"bg_color": "#ff6600"})
            obsolete_fmt = workbook.add_format({"bg_color": "#ff00ff"})
            worksheetc.set_column(0, 0, 50)
            worksheetc.set_tab_color("#948A54")
            some_fmt=workbook.add_format({"bold":'True','align': 'center'})
            colors=[added_row_format,deprecated_fmt,removed_format,obsolete_fmt,modified_format]
            worksheetc.write(0,0,'Color Coding used in Delta Document',some_fmt)
            for i in range(5):
                worksheetc.write(i+1,0,l[i],colors[i])

            
            writer2.save()
            old_json = 0
            new_json = 0
            old_yaml = 0
            new_yaml = 0
            for root, dirs, files in os.walk(filepath):
                for file in files:
                    if file.endswith(".json"):
                        old_json += 1
            for root, dirs, files in os.walk(filepath2):
                for file in files:
                    if file.endswith(".json"):
                        new_json += 1
            for root, dirs, files in os.walk(filepath):
                for file in files:
                    if file.endswith(".yaml"):
                        old_yaml += 1
            for root, dirs, files in os.walk(filepath2):
                for file in files:
                    if file.endswith(".yaml"):
                        new_yaml += 1
            old_yaml = str(old_yaml)
            new_yaml = str(new_yaml)
            old_json = str(old_json)
            new_json = str(new_json)
            u = str(len(df_OLD.index))
            vw = str(len(df_NEW.index))
            system_time = datetime.datetime.now()
            
            pt = os.path.dirname(fvc)
            
            logging.basicConfig(filename=logger_filename_main, level=logging.DEBUG) 
                    #format='%(asctime)s-> %(levelname)s: %(message)s')
            logger=logging.getLogger(__name__)
            logger.info( "This Logger File for above Delta is generated on: "
                    + str(system_time.strftime("%d-%h-%Y %H:%M"))
                    + " by "
                    + os.getlogin())
            logger.info("---------------------------------------------------------------------------")
            logger.info(" Old Input Path: " + filepath)
            logger.info(" New Input Path: " + filepath2)
            logger.info(" Number of YAML Files in Old Input is:- " + old_yaml)
            logger.info(" Number of YAML Files in New Input is:- " + new_yaml)
            logger.info(" Number of Generated JSON Files in Old Input is:- "
                    + old_json)
            logger.info(" Number of Generated JSON Files in New Input is:- "
                    + new_json)
            logger.info(" Number of Files in OLD Input Excel Sheet is:- " + u)
            logger.info(" Number of Files in New Input Excel Sheet is:- " + vw)
            
            '''             ********** EQEV-111381 *************        ''' 

            if count_n <1 and count_o <1:
                messagebox.showinfo("Success", "    Delta Generated Successfully    "+"\n\n"
                                                 "    Total old yaml files : "+ str(totold[0])+"\n"
                                                 "    Old yaml files passed :"+str(totold[1])+"\n"
                                                 "    Old yaml files failed :"+str(totold[2])+"\n"
                                                 "    Total new yaml files : "+str(totnew[0])+"\n"
                                                 "    New yaml files passed :"+str(totnew[1])+"\n"
                                                 "    New yaml files failed :"+str(totnew[2]))
            else:
                    
                new_failed=totnew[2]+count_n
                new_passed=totnew[1]-count_n                    
                old_failed=totold[2]+count_o
                old_passed=totold[1]-count_o
                    
                messagebox.showinfo("Success", "    Delta Generated Successfully    "+"\n\n"
                                                 "    Total old yaml files : "+ str(totold[0])+"\n"
                                                 "    Old yaml files passed :"+str(old_passed)+"\n"
                                                 "    Old yaml files failed :"+str(old_failed)+"\n"
                                                 "    Total new yaml files : "+str(totnew[0])+"\n"
                                                 "    New yaml files passed :"+str(new_passed)+"\n"
                                                 "    New yaml files failed :"+str(new_failed))
                
                '''             ********** EQEV-111381 end *************        ''' 
                            
        except Exception as e:
            if writer2:
                writer2.close()
            os.remove(fvc)
            logger.error(e)
            messagebox.showerror('Error',f"{e} \n Error Log file has been created under path: {ss} ")


    def clear():
        selectfileentry1.delete(0, END)
        selectfileentry2.delete(0, END)
        selectfileentry3.delete(0, END)

    ttk.Button(root2, text="Browse", command=open_file1).grid(
        row=11, column=3, padx=10, pady=5
    )

    ttk.Button(root2, text="Browse", command=open_file2).grid(
        row=12, column=3, padx=10, pady=5
    )

    ttk.Button(root2, text="Browse", command=buttontoexcel).grid(
        row=13, column=3, padx=10, pady=5
    )

    ttk.Button(
        root2, text="Generate Delta", command=lambda: [add2excel(), delta_gen(), clear()]
    ).grid(row=14, column=2, padx=10, pady=5)

    ttk.Button(
        root2, text="Close", command=lambda: [root2.destroy(), root.deiconify()]
    ).grid(row=15, column=3, padx=20, pady=10)

    ttk.Button(root2, text="Clear", command=clear).grid(
        row=15, column=1, padx=20, pady=10
    )

    root2.title("YAML Tool")
    root2.mainloop()

label = Label(
    root, text="Select the Option from Below", font=["Helvetica", 10, UNDERLINE, BOLD]
)
label.pack()  
label.configure(background="#b1cee9")

#creating drop-down list to choose the options.
options = [
    "Select Option",
    "Parse Input To Excel",
    "Get Delta",
]
clicked = StringVar()

def selected():
    #if clicked.get() == "Convert YAML to JSON":
    #    Parse_CHILD_File()
    if clicked.get() == "Parse Input To Excel":
        New_Node_Files_To_Excel()
    elif clicked.get() == "Select Option":
        messagebox.showerror("Error", "Please Select an option")
    else:
        Get_Delta()


drop = ttk.OptionMenu(root, clicked, *options)
drop.configure(width=30)
drop.pack(
    padx=10,
    pady=10,
)  
ttk.Button(root, text="OK", command=lambda: [root.withdraw(), selected()]).pack(
    side=LEFT, padx=20, pady=10
) 
ttk.Button(root, text="Close", command=root.destroy).pack(
    side=RIGHT, padx=20, pady=5
) 
# root.iconbitmap(r"H:\Desktop\eric.ico")
root.mainloop()
