import pandas as pd
import numpy as np


FinalValues= []
results=[]

# read file
df = pd.read_excel ("./files/input.xlsx") 
# remove duplicat rows 
df.drop_duplicates(inplace=True)

# create unique lists of movies and crews
movieLists = df['Movie'].unique()
crewLists = df['Crew'].unique()
crewDF = pd.DataFrame( 0 , index=crewLists , columns =crewLists )

# group data by movies and process for creating data and final array
y=0
dt = df.groupby(['Movie' ])
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

# remove 0 and 1 ties and creat final results
for row in FinalValues:
  tieval = crewDF[row[3]][row[5]]
  if(tieval>1):
    row[7]=tieval
    results.append(row)

# save matrix and final result in file  
dfresult = pd.DataFrame(results ,columns=['Movie','MovieCode','Year','Crew 1_name','Code 1','Crew 2_name','Code2', 'Tie strength']) 
dfresult.to_csv('./files/output-result.csv') 
crewDF.to_csv('./files/output-matrix.csv') 