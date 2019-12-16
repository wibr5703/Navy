#import necessary things
import numpy as np
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from datetime import datetime
import re
import sys
    
def ROTC_Func():
    #take command line argument for data
    filename = sys.argv[1]

    #read in all files
    df_HeightWeight = pd.read_excel(filename, sheet_name='HeightWeight')
    df_Input = pd.read_excel(filename, sheet_name='Input',converters={'Run':str})
    df_PRT = pd.read_excel(filename, sheet_name='PRT',converters={'RunMin':str,'RunMax':str})
    df_PassFail = pd.read_excel(filename, sheet_name='PassFail',converters={'MinScore':int,'MaxScore':int,'PassFail':str})
    df_error = pd.DataFrame(columns = list(df_Input.columns))


    #Each of these blocks of code removes incorrect data types for each column and append them to 
    #a new dataframe with all the entries that have errors
    #LAST
    indeces_LastFilter = df_Input[(df_Input.applymap(type)['Last']!=str)].index.values
    errors_last = df_Input.iloc[list(indeces_LastFilter)]

    df_error=df_error.append(errors_last)

    df_Input.drop(indeces_LastFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #FIRST
    indeces_FirstFilter = df_Input[(df_Input.applymap(type)['First']!=str)].index.values
    errors_first = df_Input.iloc[list(indeces_FirstFilter)]

    df_error=df_error.append(errors_first)

    df_Input.drop(indeces_FirstFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #STATUS
    indeces_StatusFilter = df_Input[(df_Input.Status!='Midshipman')&(df_Input.Status!='Active')].index.values
    errors_status = df_Input.iloc[list(indeces_StatusFilter)]

    df_error=df_error.append(errors_status)

    df_Input.drop(indeces_StatusFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #AGE
    df_Input['Age'] = df_Input['Age'].fillna('Blank')

    indeces_AgeFilter = df_Input[((df_Input.applymap(type)['Age']!=int)&((df_Input.applymap(type)['Age']!=float)))].index.values
    errors_age = df_Input.iloc[list(indeces_AgeFilter)]

    df_error=df_error.append(errors_age)

    df_Input.drop(indeces_AgeFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #GENDER
    indeces_GenderFilter = df_Input[(df_Input.Gender!='M')&(df_Input.Gender!='F')].index.values
    errors_gender = df_Input.iloc[list(indeces_GenderFilter)]

    df_error=df_error.append(errors_gender)

    df_Input.drop(indeces_GenderFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #HEIGHT
    df_Input['Height'] = df_Input['Height'].fillna('Blank')

    indeces_HeightFilter = df_Input[((df_Input.applymap(type)['Height']!=int)&((df_Input.applymap(type)['Height']!=float)))].index.values
    errors_height = df_Input.iloc[list(indeces_HeightFilter)]

    df_error=df_error.append(errors_height)

    df_Input.drop(indeces_HeightFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #WEIGHT
    df_Input['Weight'] = df_Input['Weight'].fillna('Blank')

    indeces_WeightFilter = df_Input[((df_Input.applymap(type)['Weight']!=int)&((df_Input.applymap(type)['Weight']!=float)))].index.values
    errors_weight = df_Input.iloc[list(indeces_WeightFilter)]

    df_error=df_error.append(errors_weight)

    df_Input.drop(indeces_WeightFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #PUSH
    df_Input['Push'] = df_Input['Push'].fillna('Blank')

    indeces_PushFilter = df_Input[((df_Input.applymap(type)['Push']!=int)&((df_Input.applymap(type)['Push']!=float)))].index.values
    errors_push = df_Input.iloc[list(indeces_PushFilter)]

    df_error=df_error.append(errors_push)

    df_Input.drop(indeces_PushFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #CRUNCH
    df_Input['Crunch'] = df_Input['Crunch'].fillna('Blank')

    indeces_CrunchFilter = df_Input[((df_Input.applymap(type)['Crunch']!=int)&((df_Input.applymap(type)['Crunch']!=float)))].index.values
    errors_crunch = df_Input.iloc[list(indeces_CrunchFilter)]

    df_error=df_error.append(errors_crunch)

    df_Input.drop(indeces_CrunchFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #RUN
    indeces_RunFilter = df_Input[(df_Input.applymap(type)['Run']!=str)].index.values
    errors_run = df_Input.iloc[list(indeces_RunFilter)]

    df_error=df_error.append(errors_run)

    df_Input.drop(indeces_RunFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    def runFormat(entry):
        runTime = entry['Run']
        if (runTime.count(':') == 1) | (runTime.count(':') == 2):
            string = runTime
            l = re.findall('[^\s:]+', string)
            if 'AM' in l:
                l.remove('AM')
            if 'PM' in l:
                l.remove('PM')
            val = True
            for i in l:
                if i.isnumeric() == False:
                    val = False    
            return val
        else:
            return False 
        
    df_Input['Run Format'] = df_Input.apply(runFormat, axis=1)

    indeces_RunFormatFilter = df_Input[(df_Input['Run Format']==False)].index.values
    errors_runformat = df_Input.iloc[list(indeces_RunFormatFilter)]
    errors_runformat = errors_runformat.drop(columns=['Run Format'])

    df_error=df_error.append(errors_runformat)

    df_Input.drop(indeces_RunFormatFilter, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    df_Input = df_Input.drop(columns=['Run Format'])


    #Hangle ages and heights out of range
    #AGES
    indeces_AgeRange = df_Input[(df_Input.Age<17)|(df_Input.Age>30)].index.values
    errors_agerange = df_Input.iloc[list(indeces_AgeRange)]
    df_error=df_error.append(errors_agerange)
    df_Input.drop(indeces_AgeRange, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #HEIGHTS
    indeces_HeightRange = df_Input[(df_Input.Height<51)|(df_Input.Height>86)].index.values
    errors_heightrange = df_Input.iloc[list(indeces_HeightRange)]
    df_error=df_error.append(errors_heightrange)
    df_Input.drop(indeces_HeightRange, inplace = True)
    df_Input.reset_index(drop=True, inplace=True)

    #Covert each column to appropriate type
    df_HeightWeight['Height'] = df_HeightWeight['Height'].astype(int)
    df_HeightWeight['MenMaxWeight'] = df_HeightWeight['MenMaxWeight'].astype(int)
    df_HeightWeight['WomenMaxWeight'] = df_HeightWeight['WomenMaxWeight'].astype(int)

    df_Input['Last'] = df_Input['Last'].astype(str)
    df_Input['First'] = df_Input['First'].astype(str)
    df_Input['Status'] = df_Input['Status'].astype(str)
    df_Input['Age'] = df_Input['Age'].astype(int)
    df_Input['Gender'] = df_Input['Gender'].astype(str)
    df_Input['Height'] = df_Input['Height'].astype(int)
    df_Input['Weight'] = df_Input['Weight'].astype(float)
    df_Input['Push'] = df_Input['Push'].astype(int)
    df_Input['Crunch'] = df_Input['Crunch'].astype(int)

    df_PRT['AgeMin'] = df_PRT['AgeMin'].astype(int)
    df_PRT['AgeMax'] = df_PRT['AgeMax'].astype(int)
    df_PRT['CrunchMin'] = df_PRT['CrunchMin'].astype(int)
    df_PRT['CrunchMax'] = df_PRT['CrunchMax'].astype(int)
    df_PRT['PushMin'] = df_PRT['PushMin'].astype(int)
    df_PRT['PushMax'] = df_PRT['PushMax'].astype(int)

    df_PRT['Status'] = df_PRT['Status'].astype(str)
    df_PRT['Gender'] = df_PRT['Gender'].astype(str)
    df_PRT['Score'] = df_PRT['Score'].astype(int)

    def removeRunMin(entry):
        RunMin = entry['RunMin']
        if (RunMin.count(':') > 1) & ((RunMin.count('AM')==1)|(RunMin.count('PM')==1)):
            return(RunMin[:-6])
        elif RunMin.count(':') > 1:
            return(RunMin[:-3])
        else:
            return(RunMin)

    def removeRunMax(entry):
        RunMin = entry['RunMax']
        if (RunMin.count(':') > 1) & ((RunMin.count('AM')==1)|(RunMin.count('PM')==1)):
            return(RunMin[:-6])
        elif RunMin.count(':') > 1:
            return(RunMin[:-3])
        else:
            return(RunMin)  

    def removeRun(entry):
        RunMin = entry['Run']
        if (RunMin.count(':') > 1) & ((RunMin.count('AM')==1)|(RunMin.count('PM')==1)):
            return(RunMin[:-6])
        elif RunMin.count(':') > 1:
            return(RunMin[:-3])
        else:
            return(RunMin)    

    #under assumption that heights are integers
    def findHeightWeightResult(entry):
        
        height = entry['Height']
        weight = entry['Weight']
        
        if entry['Gender'] == 'M' and type(entry['Height']) == int and type(entry['Weight']) == float:
            maxWeight = df_HeightWeight.loc[df_HeightWeight['Height']==height,['MenMaxWeight']].values[0][0]
            if weight <= maxWeight:
                return 'PASS'
            else:
                return 'FAIL'
        elif entry['Gender'] == 'F' and type(entry['Height']) == int and type(entry['Weight']) == float:
            maxWeight = df_HeightWeight.loc[df_HeightWeight['Height']==height,['WomenMaxWeight']].values[0][0]
            if weight <= maxWeight:
                return 'PASS'
            else:
                return 'FAIL'
            
        else:
            return 'FAIL'

    def findPushScore(entry):
        status = entry['Status']
        
        #Midshipman are scored within age range 20-24
        if (status == 'Midshipman'):
            age = 21
        else:
            age = entry['Age']
            
        gender = entry['Gender']
        push = entry['Push']
        score = df_PRT.loc[(df_PRT['AgeMin']<=age)&
                           (df_PRT['AgeMax']>=age)&
                           (df_PRT['Status']==status)&
                           (df_PRT['Gender']==gender)&
                           (df_PRT['PushMin']<=push)&
                           (df_PRT['PushMax']>=push),['Score']].values[0][0]
        return score

    def findCrunchScore(entry):
        status = entry['Status']
        
        #Midshipman are scored within age range 20-24
        if (status == 'Midshipman'):
            age = 21
        else:
            age = entry['Age']
            
        gender = entry['Gender']
        crunch = entry['Crunch']
        score = df_PRT.loc[(df_PRT['AgeMin']<=age)&(df_PRT['AgeMax']>=age)&
                           (df_PRT['Status']==status)&(df_PRT['Gender']==gender)&
                           (df_PRT['CrunchMin']<=crunch)&(df_PRT['CrunchMax']>=crunch),['Score']].values[0][0]
        return score

    def findRunScore(entry):
        status = entry['Status']
        
        #Midshipman are scored within age range 20-24
        if (status == 'Midshipman'):
            age = 21
        else:
            age = entry['Age']
            
        gender = entry['Gender']
        run = entry['Run']
        score = df_PRT.loc[(df_PRT['AgeMin']<=age)&(df_PRT['AgeMax']>=age)&
                           (df_PRT['Status']==status)&(df_PRT['Gender']==gender)&
                           (df_PRT['RunMin']<=run)&(df_PRT['RunMax']>=run),['Score']].values[0][0]
        return score

#    def getOverallScore(entry):
#        score = entry['Overall Score']
#        return df_PassFail.loc[(df_PassFail.MinScore<=score) & (df_PassFail.MaxScore>score)]['PassFail'].values[0]

    def getOverallScore(entry):
        if entry['Height Weight Result'] == 'PASS':
            score = entry['Overall Score']
            return df_PassFail.loc[(df_PassFail.MinScore<=score) & (df_PassFail.MaxScore>score)]['PassFail'].values[0]
        else:
            return 'FAIL'


    #Apply all functions
    df_Input['Height Weight Result'] = df_Input.apply(findHeightWeightResult, axis=1)
    df_Input['Push Score'] = df_Input.apply(findPushScore, axis=1)
    df_Input['Crunch Score'] = df_Input.apply(findCrunchScore, axis=1)

    df_PRT['RunMin'] = df_PRT.apply(removeRunMin, axis=1)
    df_PRT['RunMax'] = df_PRT.apply(removeRunMax, axis=1)
    df_Input['Run'] = df_Input.apply(removeRun, axis=1)

    df_PRT['RunMin'] = df_PRT['RunMin'].astype(str)
    df_PRT['RunMin'].replace(regex=True,inplace=True,to_replace=':',value='')
    df_PRT['RunMin'] = df_PRT['RunMin'].astype(int)

    df_PRT['RunMax'] = df_PRT['RunMax'].astype(str)
    df_PRT['RunMax'].replace(regex=True,inplace=True,to_replace=':',value='')
    df_PRT['RunMax'] = df_PRT['RunMax'].astype(int)

    df_Input['Run'].replace(regex=True,inplace=True,to_replace=':',value='')
    df_Input['Run'] = df_Input['Run'].astype(int)

    df_Input['Run Score'] = df_Input.apply(findRunScore, axis=1)
    df_Input['Overall Score'] = (df_Input['Push Score']+df_Input['Crunch Score']+df_Input['Run Score'])*(1/3)
    df_Input['Final Result'] = df_Input.apply(getOverallScore, axis=1)


    #Add back the : to run syntax
    def addRun(entry):
        run = str(entry['Run'])
        length = len(str(run))
        if length == 3:
            return run[0]+':'+run[1:3]
        elif length == 4:
            return run[0:2]+':'+run[2:4]
        else:
            return run
        
    df_Input['Run'] = df_Input.apply(addRun, axis=1)


    #Append the error entries to the input sheet
    df_error.reset_index(drop=True, inplace=True)
    df_error['Height Weight Result']=''
    df_error['Push Score']=''
    df_error['Crunch Score']=''
    df_error['Run Score']=''
    df_error['Overall Score']=''
    df_error['Final Result']='FAIL'

    df_Input=df_Input.append(df_error)
    df_Input.reset_index(drop=True, inplace=True)


    #Output the file with a timestamp
    df_Input.to_excel('ROTC_Automation'+'_'+str(pd.Timestamp.today())+'.xlsx')
    return

try:
    ROTC_Func()
except:
    print('ERROR: CHECK FORMAT, RULES, & NAME OF INPUT SHEET')
    
