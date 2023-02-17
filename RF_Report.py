###
# Command format : python(3) [scriptname].py [OutputfolderName] [ExcelFilePath]
###
import sys
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
# plotly
import plotly.graph_objects as go
import plotly.offline as pyo
import plotly
from plotly.subplots import make_subplots
from datetime import datetime

# importing the command line arguments in the format
args = sys.argv
# Error if more than 3 args are passed
if len(args) > 3:
    raise Exception("Arguments should not be more than 3")

# printing tha args for debugging
print("The output folder name is : " + args[1])
print("Image Folder path : " + args[2])

# Set the directory path for the images
excel_path = args[2]

# creating a dir for report
dirPath = os.getcwd()+'/ExcelReports/'
if not os.path.exists(dirPath):
    # Create the directory
    os.mkdir(dirPath)
    print(f"Directory '{dirPath}' created successfully.")
else:
    print(f"Directory '{dirPath}' already exists.")

fdf = pd.read_excel(excel_path, header=2)

df_main = fdf.iloc[:,:14]

loc = np.asarray(df_main["Position"].unique())

# Create a list to hold the tabs
tabs = []
start = df_main["Freq"].min()
end = df_main["Freq"].max()
step = 1000000

for l in loc:
    dfrf = df_main[df_main['Position']==l]
    start = dfrf["Freq"].min()
    end = dfrf["Freq"].max()
    step = 1000000
    traces = []
    maxdb = []
    maxfreq = []
    while start<end:
        # get the 10MHz freqs
        df = dfrf[(dfrf["Freq"]>=start) & (dfrf["Freq"]<start+step)]
        df = df.groupby(df["Freq"], as_index=False).aggregate({"Spectrum Values":'mean'})
        # plot the spectrum values
        #plot pylot
        x = np.asarray(df["Freq"])
        y = np.asarray(df["Spectrum Values"])
        trace = go.Scatter(x=x, y=y, mode='lines', name = "Frequency range from "+str(start)+" to "+str(start+step))
        traces.append(trace)
        maxfreq.append(df.loc[df["Spectrum Values"] == df["Spectrum Values"].max(),"Freq"].iloc[0])
        maxdb.append(df["Spectrum Values"].max())
        start = start + step

    # Create a subplot with the three traces
    fig = make_subplots(rows=len(traces), cols=1)
    for i, trace in enumerate(traces):
        fig.add_trace(trace, row=i+1, col=1)

    # Add label for each plot
    for i, trace in enumerate(traces):
        max_val = max(trace.y)
        fig.add_annotation(dict(font=dict(size=12), x=0.5, y=max_val,
                                xref='paper', yref='y'+str(i+1), showarrow=False,
                            text="Max Db is : " + str(maxdb[i])+' at frequency : '+str(maxfreq[i])))

    # Set layout for the plot
    fig.update_layout(height=5000, width=1500, title_text='Report for Spectrum Analysis')

    # Save the plot as an HTML file
    pyo.plot(fig, filename = dirPath +'RF_Plots_loc'+str(l)+'_'+datetime.now().strftime("%m-%d-%Y_%H-%M-%S")+'.html')





