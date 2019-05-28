# -*- coding: utf-8 -*-
"""
Created on Sun May 12 19:32:07 2019

@author: 364970
"""

def IMPACT(folder):
    #user input to name final excel file
    print('What to name the output xlsx?')
    name = input()

    import os, glob
    import pandas as pd
    import matplotlib.pyplot as plt
    import numpy as np
    import scipy.integrate as integrate

    #sets working directory to the folder with raw data files
    os.chdir(folder)
    
    # .glob returns a list of filenames matching a pattern 
    # * represents any text. this line returns a list of all .csv files in the directory
    filenames = glob.glob('*Ch_1.csv')
    
    col_names = ['Point ID', 'Time (ms)', 'Force (N)', 'Energy (J)', 'Displacement (mm)', 'Velocity (m/s)', 'Voltage Ch1 (mV)']
    
    dfs = [pd.read_csv(filename, skiprows = 7, delimiter = '\t', names = col_names).dropna(axis = 0).astype(float) for filename in filenames]
    time =[]
    force = []
    energy = []
    
    
    
   #Creates the Overlay 
    plt.figure()
    for i in range(len(list(filenames))):
        T = dfs[i]['Time (ms)']
        time.append(T)
        time_df = pd.concat(time)
        F = dfs[i]['Force (N)']
        force.append(F)
        E = dfs[i]['Energy (J)']
        energy.append(E)
        plt.plot(time[i], force[i], label= 'Specimen ' + str(i+1)) 
        plt.xlim(xmin = 0, xmax = time_df.max()*1.05)
        plt.ylim(ymin = 0)
        plt.xlabel('Time (ms)')
        plt.ylabel('Force (N)')
        plt.legend()
        plt.savefig('Overlay')
        
 # Calculations for peak force       
    peak_force = []
    total_energy = []
    for i in range(len(list(filenames))):
        fmax = force[i].max()
        peak_force.append(fmax)
        e_tot= energy[i].max()
        total_energy.append(e_tot)
    
    
    
        
        
    for i in range(len(list(filenames))):
        plt.figure()
        plt.plot(time[i], force[i], label= 'Specimen ' + str(i+1)) 
        plt.xlabel('Time (ms)')
        plt.ylabel('Force (N)')
        plt.ylim(ymin=0)
        plt.xlim(xmin=0)
        plt.title('Specimen' + str(i+1))
        plt.savefig('Plot' + str(i+1))
        
    labels=[]
    for i in range(len(list(filenames))):
        labels.append('Specimen ' + str(i+1))
    force_energy = [peak_force, total_energy]
    force_energy = pd.DataFrame(force_energy, index = ['Peak Force (N)', 'Total Energy (J)'], columns = labels)
    dev = force_energy.std(axis=1).round(decimals=3)
    avg= force_energy.mean(axis=1).round(decimals=2)
    force_energy['AVG'] = avg
    force_energy['DEV'] = dev
    with pd.ExcelWriter(name) as writer: 
        for i in range(len(list(filenames))):
            dfs[i].to_excel(writer, sheet_name = 'Sample' + str(i+1) + 'RawData', index = False) #sends dfs to their own sheet in the generated excel file
            worksheet = writer.sheets['Sample' + str(i+1) + 'RawData']
            worksheet.insert_image('K1', 'Plot' + str(i+1) + '.png')
        force_energy.to_excel(writer, sheet_name='Report')
        reportsheet = writer.sheets['Report']
        reportsheet.insert_image('A6', 'Overlay.png')
        reportsheet.set_column(0,15,20)

        
    

IMPACT('C:\\Users\\364970\\Documents\\python\\18-0260')
    