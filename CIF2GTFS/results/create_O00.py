import sys
import os

sys.path.append(os.path.join(os.path.dirname(__file__), "src"))
import pandas as pd

import get_inputs as gi
path = os.path.dirname(__file__)
input_path = os.path.join(path, "input\\inputs.csv")
runID = gi.getRunID(input_path)

parquetCompression = 'snappy'

def create_O00():
    VJ_list = Visum.Workbench.Lists.CreateVehJourneyList
    for col in ['No', 'ATOC', r'FIRST:VEHJOURNEYSECTIONS\DAYVECTOR']:
        VJ_list.AddColumn(col)

    dfVJs = pd.DataFrame(VJ_list.SaveToArray(), columns=['No', 'ATOC', 'DayVector'])

    fromDate = pd.to_datetime(Visum.Net.CalendarPeriod.AttValue('ValidFrom'), format="%d.%m.%Y")
    toDate = pd.to_datetime(Visum.Net.CalendarPeriod.AttValue('ValidUntil'), format="%d.%m.%Y")
    
    dateRange = pd.date_range(fromDate, toDate).strftime("%d-%m-%Y").tolist()

    for i, date in enumerate(dateRange):
        dfVJs[f'{date}'] = dfVJs.DayVector.str[i].astype(int)
        dfVJs = dfVJs.copy()
    
    dfVJs.drop(['DayVector'], axis=1, inplace=True)

    dfVJsLong = dfVJs.melt(['No', 'ATOC'], var_name='Date', value_name='Services')
    dfVJsSummary = dfVJsLong.groupby(['ATOC', 'Date'], as_index=False).Services.sum()

    dfVJsSummary.to_parquet(f"results\\{runID}_O00_DailyServicesByTOC.parquet", index=False, compression=parquetCompression)

    del dfVJs
    del dfVJsLong
    del dfVJsSummary


if __name__ == '__main__':
    create_O00()