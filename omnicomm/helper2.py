import pandas as pd

dfB = pd.DataFrame({'X':[1,2,3],'Y':[1,2,3], 'Time':[10,20,30]})
dfA = pd.DataFrame({'X':[1,1,2,2,2,3],'Y':[1,1,2,2,2,3], 'ONSET_TIME':[5,7,9,16,22,28],'COLOR': ['Red','Blue','Blue','red','Green','Orange']})

#create one single table
mergeDf = pd.merge(dfA, dfB, left_on = ['X','Y'], right_on = ['X','Y'])
#remove rows where time is less than onset time
filteredDf = mergeDf[mergeDf['ONSET_TIME'] < mergeDf['Time']]
#take min time (closest to onset time)
groupedDf = filteredDf.groupby(['X','Y']).max()

print(filteredDf)