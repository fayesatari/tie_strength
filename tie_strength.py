import pandas as pd
import numpy as np
from datetime import datetime

FinalValues= []
results=[]

now = datetime.now()
# read file
try: 
  print("Start reading file at " + now.strftime("%H:%M:%S"))
  df = pd.read_excel ("./files/input.xlsx" , sheet_name=0)
  dflen = len(df)
  print(" Number of rows in file: " + str(dflen) )
except ValueError:
  print("File or path could not find.The file name should be inpout.xlsx and it must be located in the folder is called files.")
  input("Press enter to exit ")
 
# remove duplicat rows 
try: 
  print("Start removing the  duplicates." )
  df.drop_duplicates(inplace=True)  
  dfremovedlen=  len(df)
  print(str(dflen - dfremovedlen ) + " duplicate rows is found and deleted.")
except ValueError:
  print("Error in removing duplicte records.")
  input("Press enter to exit ")
 
# create unique lists of movies and crews
try:
  movieLists = df['Movie'].unique()
  movieQty = len(movieLists)
  crewLists = df['newcode'].unique()
  crewDF={}
except ValueError:
  print("Error in Finding unique movies and crews.")
  input("Press enter to exit ")

# group data by movies and process for creating data and final array
try:
  y=0
  z=0
  dt = df.groupby(['Movie' ])
  print("Start Processing data of " + str(movieQty) + " movies")
  
  for x in movieLists:
    if( (z % 10)==0 ):
      print("Start processing movie number " + str(z+1) + " from " + str(movieQty) 
         + " (" + str(int(((z+1)/ movieQty) * 100))  + "%)" )
    movie=dt.get_group(x)
    crew = movie['newcode']
    datamovie=movie.iloc[0]
    for first in crew:
      for second in crew:
        if not (first == second ):
          if first not in crewDF:
            crewDF[first]={}
            crewDF[first][second] = 1            
          elif second not in crewDF[first]:
            crewDF[first][second] = 1
          else: 
            crewDF[first][second] += 1 
          if second not in crewDF:
            crewDF[second]={}
            crewDF[second][first] = -1
          
          FinalValues.append([])
          tmparr =[datamovie['Movie'] , datamovie['MovieCode'] , datamovie['Year'],
                           first , movie.loc[movie['newcode'].eq(first), 'Crew'].values[0],
                           second ,movie.loc[movie['newcode'].eq(second), 'Crew'].values[0], 1]
          FinalValues[y]= tmparr
          y=y+1
     
    z=z+1         
  
 
except ValueError:
  print("Unexpected error in processing data." )
  input("Press enter to exit ")

# remove 0 and 1 ties and creat final results
try:
  for row in FinalValues:
    tieval = crewDF[row[3]][row[5]]  
    if(tieval == -1): 
      tieval = int(crewDF[row[5]][row[3]])
    if(tieval>1):
      row[7]=tieval
      results.append(row)
except ValueError:
  print("Error in preaparing final results." )
  input("Press enter to exit ")

# save matrix and final result in file  
try:
  print("Start Saving the results of " + str(len(results)) + " data.")
  dfresult = pd.DataFrame(results ,columns=['Movie','MovieCode','Year','Crew 1_name','Code 1','Crew 2_name','Code2', 'Tie strength']) 
  dfresult.to_csv('./files/output-result.csv') 
  #crewDF.to_csv('./files/output-matrix.csv') 
  now = datetime.now() 
  input( "End of process in " +  now.strftime("%H:%M:%S") + ".  Press enter to exit ")
except ValueError:
  print("Error in saving the results." )
  input("Press enter to exit ")  

