import matplotlib as mpl
mpl.use('Agg')
import matplotlib.pyplot as plt
from random import randint


def piechart():


    fig = plt.figure(figsize=(10, 8))

    # defining labels 
    activities = ['In-Progrss', 'Defined', 'Completed', 'Accepted'] 

    # portion covered by each label 
    slices = [70, 12, 9, 21] 

    # color for each label 
    colors = ['r', 'y', 'g', 'b'] 

    # plotting the pie chart 
    plt.pie(slices, labels = activities, colors=colors, 
            startangle=90, shadow = True, explode = (0, 0, 0.1, 0), 
            radius = 1.2, autopct = '%1.1f%%') 

    # plotting legend 
    plt.legend() 

    fig.savefig('graph.png')
