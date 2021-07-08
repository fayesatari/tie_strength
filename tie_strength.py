import pandas as pd
import numpy as np


FinalValues= []
results=[]

# read file
try: 
  print("Start reading file.")
  df = pd.read_excel ("./files/input.xlsx" , sheet_name=0)
except ValueError:
  print("File or path could not find.The file name should be inpout.xlsx and it must be located in the folder is called files.")
 
# remove duplicat rows 
try: 
  df.drop_duplicates(inplace=True)
  print("Remove the  duplicates.")
except ValueError:
  print("Error in removing duplicte records.")
 
# create unique lists of movies and crews
try:
  movieLists = df['Movie'].unique()
  crewLists = df['Crew'].unique()
  crewDF = pd.DataFrame( 0 , index=crewLists , columns =crewLists )
except ValueError:
  print("Error in Finding unique movies and crews.")

# group data by movies and process for creating data and final array
try:
  y=0
  dt = df.groupby(['Movie' ])
  print("Start Processing data.")
  for x in movieLists:
    movie=dt.get_group(x)
    crew = movie['Crew']
    datamovie=movie.iloc[0]
    for first in crew:
      for second in crew:
        if not (first == second ):
          crewDF[first][second] +=1
          FinalValues.append([])
          tmparr =[datamovie['Movie'] , datamovie['MovieCode'] , datamovie['Year'],
                           first , movie.loc[movie['Crew'].eq(first), 'newcode'].values[0],
                           second ,movie.loc[movie['Crew'].eq(second), 'newcode'].values[0], 1]
          FinalValues[y]= tmparr
          y=y+1
          if((y % 3000)==0):
            print("Continue processing - row " + str(y))
except ValueError:
  print("Unexpected error in processing data." )

# remove 0 and 1 ties and creat final results
try:
  for row in FinalValues:
    tieval = crewDF[row[3]][row[5]]
    if(tieval>1):
      row[7]=tieval
      results.append(row)
except ValueError:
  print("Error in preaparing final results." )

# save matrix and final result in file  
try:
  print("Start Saving the results")
  dfresult = pd.DataFrame(results ,columns=['Movie','MovieCode','Year','Crew 1_name','Code 1','Crew 2_name','Code2', 'Tie strength']) 
  dfresult.to_csv('./files/output-result.csv') 
  crewDF.to_csv('./files/output-matrix.csv') 
except ValueError:
  print("Error in saving the results." )  