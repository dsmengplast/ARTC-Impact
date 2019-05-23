# -*- coding: utf-8 -*-
"""
Created on Wed May 22 09:21:19 2019

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
    filenames = glob.glob('*RawData_1.csv')
    
    col_names = ['Point ID', 'Time (ms)', 'Force (N)', 'Velocity (m/s)', 'Energy (J)', 'Displacement (mm)']
    
    dfs = [pd.read_csv(filename, names = col_names, header = 2).dropna(axis = 0).reset_index().astype(float) for filename in filenames]
    
    time =[]
    force = []
    energy = []
    for i in range(len(list(filenames))):
        T = dfs[i]['Time (ms)']
        time.append(T)
        time_df = pd.concat(time)
        F = dfs[i]['Force (N)']
        force.append(F)
        E = dfs[i]['Energy (J)']
        energy.append(E)
     
    peak_end_times = []
    peak_force = []  
    etot = []
    plt.figure(figsize=(15,10))
    for i in range(len(list(filenames))):
        tf = force[i][:6000]
        PF = tf.max()
        peak_force.append(PF)
        Pindex= tf.idxmax()
        tf_low = tf[0:Pindex]
        tf_high = tf[Pindex:]
        SP = PF*.005
        PT1 = tf_low.where(tf_low<=SP).dropna()
        PT2 = tf_high.where(tf_high<=SP).dropna()
        point1 = list(PT1.index.values)
        point2 =list(PT2.index.values)
        peak_start = point1[-1]+1
        peak_end = point2[0]-1
        ETOT = energy[i][peak_end]
        etot.append(ETOT)
        
        adj_time = time[i] - time[i][peak_start]
        plt.plot(adj_time[peak_start:peak_end], tf[peak_start:peak_end], label = 'Specimen ' + str(i+1))
        peak_end_times.append(adj_time[peak_end])
        plt.ylim(ymin=0)
        plt.xlim(xmin = 0, xmax = max(peak_end_times))
        plt.xlabel('Time (ms)')
        plt.ylabel('Force (N)')
        plt.legend()
        plt.savefig('Overlay')
        

       
    #Indiviual plotting    
#    for i in range(len(list(filenames))):
#        plt.figure()
#        plt.plot(time[i][0:main_peak[i]], force[i][0:main_peak[i]], label= 'Specimen ' + str(i+1)) 
#        plt.xlabel('Time (ms)')
#        plt.ylabel('Force (N)')
#        plt.ylim(ymin=0)
#        plt.xlim(xmin=0)
#        plt.title('Specimen' + str(i+1))
#        plt.savefig('Plot' + str(i+1))
        
        
    #Info for report    
    labels=[]
    for i in range(len(list(filenames))):
        labels.append('Specimen ' + str(i+1))
    force_energy = [peak_force, etot]
    force_energy = pd.DataFrame(force_energy, index = ['Peak Force (N)', 'Total Energy (J)'], columns = labels)
    dev = force_energy.std(axis=1).round(decimals=3)
    avg= force_energy.mean(axis=1).round(decimals=2)
    force_energy['AVG'] = avg
    force_energy['DEV'] = dev
    
    
    with pd.ExcelWriter(name) as writer: 
        for i in range(len(list(filenames))):
            dfs[i].to_excel(writer, sheet_name = 'Sample' + str(i+1) + 'RawData', index = False) #sends dfs to their own sheet in the generated excel file
            worksheet = writer.sheets['Sample' + str(i+1) + 'RawData']
            #worksheet.insert_image('K1', 'Plot' + str(i+1) + '.png')
        force_energy.to_excel(writer, sheet_name='Report')
        reportsheet = writer.sheets['Report']
        reportsheet.insert_image('A6', 'Overlay.png', {'x_scale': 0.5, 'y_scale': 0.5})
        reportsheet.set_column(0,15,20)

        
    

IMPACT('C:\\Users\\364970\\Documents\\python\\19-0080')