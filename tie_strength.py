import pandas as pd
import numpy as np
from datetime import datetime


FinalValues= []
results=[]

now = datetime.now()
# read file
try: 
  print("Start reading the file at " + now.strftime("%H:%M:%S"))
  df = pd.read_excel ("./files/input.xlsx" , sheet_name=0)
  dflen = len(df)
  print("Number of rows in the file: " + str(dflen) )  
except ValueError:
  print("File or path could not find. The file name should be inpout.xlsx, and it must be located in the folder is called files.")
  input("Press enter to exit ")
 
# remove duplicat rows 
try:
  print("Start finding movies or crews without code.")
  is_NaN = df.isnull()
  row_has_NaN = is_NaN.any(axis=1)
  rows_with_NaN = df[row_has_NaN]
  if(len(rows_with_NaN) > 0) : 
    print("Please checklist bellow and update new code or MovieCode values with the correct number: ")
    print(rows_with_NaN)
    input("Press enter to exit ")    
except ValueError:
  print("Error in finding movies or crews without code.")
  input("Press enter to exit ")

try:
  crewGroup = df.groupby('newcode')
  filteredDf= df[~df['newcode'].isin(crewGroup.filter(lambda x: len(x) == 1)["newcode"])]
  dflen = len(filteredDf)
  print ("The number of rows for processing: " + str(dflen))
except ValueError:
  print("Error in cleansing crews with only one film.")
  input("Press enter to exit ")

# remove duplicat rows 
try: 
  print("Start removing the duplicates." )
  filteredDf.drop_duplicates(inplace=True)  
  dfremovedlen=  len(filteredDf)
  print(str(dflen - dfremovedlen ) + " duplicate rows are found and deleted.")
except ValueError:
  print("Error in removing duplicate records.")
  input("Press enter to exit ")
 
# create unique lists of movies and crews
try:
  filteredDf = filteredDf.sort_values(by=['MovieCode', 'newcode'])
  movieLists = filteredDf['MovieCode'].unique()
  movieQty = len(movieLists)
  
  crewDF={}
except ValueError:
  print("Error in Finding unique movies and crews.")
  input("Press enter to exit ")

# group data by movies and process for creating data and final array
try:
  y=0
  z=0
  dt = filteredDf.groupby(['MovieCode' ])
  print("Start processing data of " + str(movieQty) + " movies")
  
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
            if(crewDF[first][second]>0):
              crewDF[first][second] += 1 
          if second not in crewDF:
            crewDF[second]={}
            crewDF[second][first] = -1
          elif first not in crewDF[second]:
            crewDF[second][first] = -1
          
          if(crewDF[first][second] > 0) : 
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
    #if(tieval == -1): 
     # tieval = int(crewDF[row[5]][row[3]])
    if(tieval>1):
      row[7]=tieval
      results.append(row)

  #pd.DataFrame(results).to_csv('./files/results.csv')  
except ValueError:
  print("Error in preparing  final results." )
  input("Press enter to exit ")

# save matrix and final result in file  
try:
  print("Start Saving the results of " + str(len(results)) + " data.")
  dfresult = pd.DataFrame(results ,columns=['Movie','MovieCode','Year','Code 1','Crew 1_name','Code2','Crew 2_name', 'Tie strength']) 
  dfresult.to_csv('./files/output-result.csv') 
  
  dfresult1 = dfresult[['Code 1','Crew 1_name','Code2','Crew 2_name', 'Tie strength']]
  dfresult1.drop_duplicates(inplace=True)
  print("Start Saving the results 1 of " + str(len(dfresult1)) + " data.")
  dfresult1.to_csv('./files/output-result-1.csv') 

  now = datetime.now() 
  input( "End of process in " +  now.strftime("%H:%M:%S") + ".  Press enter to exit ")
except ValueError:
  print("Error in saving the results." )
  input("Press enter to exit ")  

