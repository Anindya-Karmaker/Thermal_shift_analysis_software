import pandas as pd
from scipy.signal import find_peaks
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import messagebox
from matplotlib.figure import Figure
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import *
from tkinter import ttk
import time
import threading
global filename    
global data_main
global data_names
global wells_filled
global high, low
global final_data, done
done=0
options=["Polynomial Method", "Differential Fluorescence Method"]
def interface():
    
    def threaded_load_func():
        done=0
        t1=threading.Thread(target=select_file)
        t1.start()
        if done==1:
            t1.join()
    def threaded_save_func():
        done=0
        t2=threading.Thread(target=save_file)
        t2.start()
        if done==1:
            t2.join()
    def solve_graph(method,x,y,z,output_ok):
        
        if method=="Polynomial Method":
            poly=np.polyfit(x, y, deg=3)
            correlation_matrix = np.corrcoef(x, np.polyval(poly, x))
            correlation_xy = correlation_matrix[0,1]
            r_squared = correlation_xy**2
            first_der=np.polyder(poly)
            second_der=np.polyder(first_der)
            Tm=-second_der[1]/second_der[0]   
            diff_poly = np.polyfit(x, z, deg=3)
            diff_data=np.polyval(diff_poly, x)
        elif method=="Differential Fluorescence Method":
            poly=np.polyfit(x, y, deg=3)
            correlation_matrix = np.corrcoef(x, np.polyval(poly, x))
            correlation_xy = correlation_matrix[0,1]
            r_squared = correlation_xy**2
            diff_poly = np.polyfit(x, z, deg=3)
            diff_data=np.polyval(diff_poly, x)
            diff_data_ABS=abs(np.polyval(diff_poly, x))
            peaks, _ = find_peaks(diff_data_ABS)
            a=peaks.tolist()
            if a==[]:
                messagebox.showinfo("STATUS","Unable to evaluate data. Please narrow the temperature range!")
                Tm=0
            else:
                axx=x.tolist()
                Tm=axx[a[0]]            
        if output_ok==1:  
            out="R-squared value(Fluorescence): "+ str(round(r_squared,2))
            optlabelResultsLabel_Rsq.configure(text=out)                
            out="Melting Point, Tm: "+ str(round(Tm,2))   
            optlabelResultsOut.configure(text=out)
            ax.plot(final_data['T'], final_data['F'],  label='Raw Fluorescence Data')
            ax.plot(x, np.polyval(poly, x), label='Fitted Fluorescence Data', color='red')
            ax.plot(final_data['T'], final_data['dF'],  label='Differential Fluorescence Data')        
            ax.plot(x, diff_data, label='Fitted Differential Data', color='m')
            ax.axvline(Tm, color='k', label=f'Inflection Point')
            ax.set_xlabel('Temperature')
            ax.set_ylabel('Fluorescence', labelpad=0)
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.16), ncol=3)
            chart_type.draw()
            plt.pause(0.0001)
            time.sleep(0.1)            
        return Tm, r_squared, poly, diff_poly
    def select_file():
        global filename, done
        filetypes = (
            ('Excel files', ['*.xlsx','*.xlsb','*.xls']),
            ('All files', '*.*')
        )
    
        filename = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)
        print(filename)
        global data_main
        global data_names
        global wells_filled
        wells_filled=0
        listPlates.delete(0,END)
        # try:
        root.title("Thermal Shift Assay Analyzer: LOADING DATA")
        data_names=pd.read_excel(filename, sheet_name="Sample Setup",header=39)
        data_main=pd.read_excel(filename, sheet_name="Melt Curve Raw Data",header=39)
        # data_main=pd.read_excel(filename, sheet_name=2,header=39)
        print(data_main)
        wells_filled=data_main['Well Position'].tolist()
        wells_filled=list(dict.fromkeys(wells_filled))
        print("LOADING COMPLETED")   
        
        for items in range(len(wells_filled)):
            listPlates.insert(items, wells_filled[items])
        root.title("Thermal Shift Assay Analyzer V3.00")
        save_button.config(state=NORMAL)
        messagebox.showinfo("STATUS","Data Loading Successful")
        
        scrollbar.config(command=listPlates.yview)
        listPlates.config(yscrollcommand=scrollbar.set)
        # except:
        #     root.title("Thermal Shift Assay Analyzer V3.00")
        #     messagebox.showinfo("STATUS","Data Format Mismatch. Please select the right file!")
        done=1
   
    
    def save_file():
        # try:
            file_name= fd.asksaveasfilename(defaultextension=".xlsx", title="Save File", filetypes=[("Excel File", "*.xlsx")])
            global wells_filled, done
            tag_list=list()
            for i in range(len(wells_filled)):
                tag=data_names.loc[data_names['Well Position']==wells_filled[i]]['Sample Name'].tolist()[0]
                tag_list.append(tag)    
            names_dict=dict(zip(wells_filled, tag_list))
            data=pd.DataFrame()
            out_dat=pd.DataFrame()
            out2_dat=pd.DataFrame(columns=['Sample','Melting point','R^2 Value'])
            root.title("Thermal Shift Assay Analyzer: SAVING DATA")
            for i in range(len(wells_filled)):
                 sample_fluorescence=data_main.loc[data_main['Well Position']==wells_filled[i]]['Fluorescence'].tolist()
                 sample_temperature=data_main.loc[data_main['Well Position']==wells_filled[i]]['Temperature'].tolist()
                 sample_dt_fluorescence=data_main.loc[data_main['Well Position']==wells_filled[i]]['Derivative'].tolist()

                 final_data=pd.DataFrame(list(zip(sample_temperature,sample_fluorescence,sample_dt_fluorescence)), columns=['T','F','dF'])
                 high=int(optionHigh.get())
                 low=int(optionLow.get())
                 x=final_data.loc[(final_data['T']>=low) & (final_data['T']<=high)]
                 # poly=np.polyfit(x['T'], x['F'], deg=3)
                 # diff_poly = np.polyfit(x['T'], x['dF'], deg=3) 
                 # correlation_matrix = np.corrcoef(x['T'], np.polyval(poly, x['T']))
                 # correlation_xy = correlation_matrix[0,1]
                 # r_squared = correlation_xy**2
                 # first_der=np.polyder(poly)
                 # second_der=np.polyder(first_der)
                 # Tm=-second_der[1]/second_der[0]
                 Tm, r_squared, poly, diff_poly = solve_graph(clicked.get(), x['T'], x['F'], x['dF'],0)
                 s_name=names_dict.get(wells_filled[i])
                 column_names='Fluorescence '+str(s_name)+":"+str(wells_filled[i])
                 column_names_dF='Differential Fluorescence '+str(s_name)+":"+str(wells_filled[i])
                 print(column_names)
                 out2_dat.loc[len(out2_dat.index)]=[str(s_name)+":"+str(wells_filled[i]), Tm, r_squared]
                 out_dat['Temperature']=sample_temperature
                 out_dat[column_names]=sample_fluorescence
                 column_name_fit='Fluorescence FIT '+str(s_name)+":"+str(wells_filled[i])
                 column_name_fit2='Differential Fluorescence FIT '+str(s_name)+":"+str(wells_filled[i])
                 out_dat[column_name_fit]=np.polyval(poly, sample_temperature)
                 out_dat[column_names_dF]=sample_dt_fluorescence
                 out_dat[column_name_fit2]=np.polyval(diff_poly, sample_temperature)
                 data['Temperature']=sample_temperature
                 data[column_names]=sample_fluorescence
                 data[column_names_dF]=sample_dt_fluorescence
            with pd.ExcelWriter(file_name, engine="openpyxl") as writer:  
                data.to_excel(writer, sheet_name='RAW')
                out_dat.to_excel(writer, sheet_name='OUTPUT')
                out2_dat.to_excel(writer, sheet_name='OUTPUT FINAL')
            print(data)
            root.title("Thermal Shift Assay Analyzer V3.00")
            messagebox.showinfo("STATUS","Data Saved Successfully")
            done=1
        # except:
            # root.title("Thermal Shift Assay Analyzer: Nothing Saved")   
    def listPlatesChoose(evt):
        global data_main
        global data_names
        global wells_filled, final_data
        chosenplate=str((listPlates.get(ANCHOR)))
        print(chosenplate)
        ax.clear()    
        well_name=chosenplate
        well_selected=wells_filled.index(well_name)
        tag_list=list()
        for i in range(len(wells_filled)):
            tag=data_names.loc[data_names['Well Position']==wells_filled[i]]['Sample Name'].tolist()[0]
            tag_list.append(tag)
        
        names_dict=dict(zip(wells_filled, tag_list))
        s_name=names_dict.get(well_name)
        print(s_name)
        out_text="Sample Name: "+str(s_name)
        Sample_Name.configure(text=out_text)
        sample_fluorescence=data_main.loc[data_main['Well Position']==wells_filled[well_selected]]['Fluorescence'].tolist()
        sample_dt_fluorescence=data_main.loc[data_main['Well Position']==wells_filled[well_selected]]['Derivative'].tolist()
        sample_temperature=data_main.loc[data_main['Well Position']==wells_filled[well_selected]]['Temperature'].tolist()
        final_data=pd.DataFrame(list(zip(sample_temperature,sample_fluorescence, sample_dt_fluorescence)), columns=['T','F','dF'])
        
        # high=int(input("Input High Temperature Limit: "))
        # low=int(input("Input Low Temperature Limit: "))
        # high=65
        # low=45
        
        high=int(optionHigh.get())
        low=int(optionLow.get())
        x=final_data.loc[(final_data['T']>=low) & (final_data['T']<=high)]
        print(x)    
        poly=np.polyfit(x['T'], x['F'], deg=3)
        print(poly)
        solve_graph(clicked.get(), x['T'], x['F'], x['dF'],1)
        #testing new
        
        # poly4=np.polyfit(x['T'], x['F'], deg=5)
        # second_der_p4=np.polyder(poly4)
        # rootsp=np.roots(second_der_p4).tolist()
        # print("ROOTS: ",rootsp)
        # for i in range(len(rootsp)):
        #     x1=np.polyval(second_der_p4,rootsp[i])
        #     x_below=np.polyval(second_der_p4,rootsp[i]-1)
        #     x_above=np.polyval(second_der_p4,rootsp[i]+1)
        #     print("FOR ROOT: ", rootsp[i])
        #     print("VALUE 2nd der: ", x1)
        #     print("X_ABOVE: ", x_above)
        #     print("X_BELOW: ", x_below)
        #     if((x_below<0 and x_above>0) or (x_below>0 and x_above<0) and (isinstance(rootsp[i], complex)==False) and (low<rootsp[i]<high==True)):
        #         print("THE CONVENTIONAL MELTING POINT IS: ", rootsp[i])
        #         Tm2=rootsp[i]
        #     else:
        #         Tm2=0
        #spl = UnivariateSpline(x['T'], x['F'], k=3)
        #test_data=spl(x['T'])
        #print("TESTING DATA: ", test_data)
        #smooth_d2 = np.gradient(np.gradient(test_data))
        #infls = np.where(np.diff(np.sign(smooth_d2)))[0]
        #print("NEW TM: ", infls)
        
        #testing done
        
        # correlation_matrix = np.corrcoef(x['T'], np.polyval(poly, x['T']))
        # correlation_xy = correlation_matrix[0,1]
        # r_squared = correlation_xy**2
        # print("THE R-squared value is", r_squared)
        # out="R-squared value: "+ str(round(r_squared,2))
        # optlabelResultsLabel_Rsq.configure(text=out)
        # first_der=np.polyder(poly)
        # second_der=np.polyder(first_der)
        # print(second_der)
        # Tm=-second_der[1]/second_der[0]
        # out="Melting Point, Tm: "+ str(round(Tm,2))
        # print("THE MELTING POINT IS: ", Tm)
        # optlabelResultsOut.configure(text=out) 
        
        # diff_poly = np.polyfit(x['T'], x['dF'], deg=3)
        # print("DIFF POLY: ",diff_poly)
        # diff_data=np.polyval(diff_poly, x['T'])
        # print(diff_data)
        # ax.plot(final_data['T'], final_data['F'],  label='Raw Data')
        # ax.plot(x['T'], np.polyval(poly, x['T']), label='Fitted Data', color='c')
        # ax.plot(final_data['T'], final_data['dF'],  label='Differential Data')        
        # ax.plot(x['T'], diff_data, label='Fitted Differential Data', color='m')
        # ax.axvline(Tm, color='k', label=f'Inflection Point')
        # ax.set_xlabel('Temperature')
        # ax.set_ylabel('Fluorescence')
        # ax.legend()
        # chart_type.draw()
        # plt.pause(0.0001)
        # time.sleep(0.1)
        
    root=tk.Tk()
    root.geometry('550x700')
    root.resizable(False, False)
    root.title("Thermal Shift Assay Analyzer V3.00")
    lbox=tk.Frame(root)
    fig = Figure(figsize = (5, 5),
                     dpi = 85)
    ax = fig.add_subplot()
    a=[0,0,0,0,0]
    b=[0,0,0,0,0]
    ax.plot(a,b)
    plt.style.use('seaborn')
    chart_type = FigureCanvasTkAgg(fig, root)
    optionHigh=Entry(root, font=("Arial", 14))
    optionHigh.insert(0,'70')
    optionLow=Entry(root, font=("Arial", 14))
    optionLow.insert(0,'60')
    optlabelResultsOut=Label(root, text="Melting Point, Tm: -----------", font=("Arial", 12))
    # optlabelResults=Label(root, text="--------", font=("Arial", 12))
    Sample_Name=Label(root, text="Sample Name: -----------", font=("Arial", 12))
    # optlabelResults_RSQ=Label(root, text="--------", font=("Arial", 12))
    listPlates= tk.Listbox(lbox, width=35, height=5, font=("Helvetica", 12), selectmode=BROWSE,exportselection=False)
    scrollbar = Scrollbar(lbox, orient="vertical", width=30)
    scrollbar.pack(side=RIGHT,anchor=N,fill='y')
    open_button=tk.Button(root, text='Load Excel', font=("Arial", 12), command=threaded_load_func)
    open_button.pack(side=TOP, expand=True, fill='x')
    save_button=tk.Button(root, text='Save Excel', font=("Arial", 12), command=threaded_save_func)
    save_button.config(state=DISABLED)
    save_button.pack(side=TOP, expand=True, fill='x')
    # optlabelHigh=Label(root, text="Select Method", font=("Arial", 12))
    # optlabelHigh.pack(side=TOP, fill='both')
    clicked = StringVar()
    clicked.set(options[0])
    drop = OptionMenu(root, clicked, *options)#, font=("Arial", 12))
    drop.config(font=("Arial", 12))
    drop.pack(side=TOP, fill='x')
    optlabelHigh=Label(root, text="Enter High value", font=("Arial", 12))
    optlabelHigh.pack(side=TOP, fill='x')
    optionHigh.pack(side=TOP,  fill='x')
    optlabelLow=Label(root, text="Enter Low value", font=("Arial", 12))
    optlabelLow.pack(side=TOP,  fill='x')
    optionLow.pack(side=TOP,  fill='x')
    optlabelPlate=Label(root, text="Select Plate", font=("Arial", 12))
    optlabelPlate.pack(side=TOP, fill='x')
    listPlates.pack(side=TOP,  fill='x')
    lbox.pack(side=TOP, fill='both')
    Sample_Name.pack(side=TOP, anchor=NW)
    # optlabelResultsLabel=Label(root, text="Tm value: ", font=("Arial", 12))
    optlabelResultsOut.pack(side=TOP, anchor=NW)
    # optlabelResults.pack(side=TOP, anchor=NW)
    optlabelResultsLabel_Rsq=Label(root, text="R-squared value(Fluorescence): -----------", font=("Arial", 12))
    optlabelResultsLabel_Rsq.pack(side=TOP, anchor=NW)
    # optlabelResults_RSQ.pack(side=TOP, anchor=NW)
    chart_type.get_tk_widget().pack(side=BOTTOM, anchor=N, fill="x")
    listPlates.bind('<<ListboxSelect>>',listPlatesChoose)
    
    root.mainloop()

if __name__=='__main__':
    interface()