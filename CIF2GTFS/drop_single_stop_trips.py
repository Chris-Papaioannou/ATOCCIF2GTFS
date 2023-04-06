import os
import pandas as pd

def main():
    #Define path and read oringinal stop times gtfs output from C# CIF to GTFS process
    path = os.path.dirname(__file__)
    df = pd.read_csv(os.path.join(path, 'cached_data\\STOP_TIMES\\full.txt'), low_memory = False)

    #Get pandas DataFrame of unique trip IDs (i.e. single stop trips) only for reporting
    report = df['trip_id'].drop_duplicates(keep = False)
    if len(report) > 0:
        print(f'WARNING (Prio. = High): {len(report)} trips were dropped as they only had one stop. The following trip IDs were affected:')
        print(report)
    else:
        print('NOTE: No trips were dropped as a result of only having one stop.')

    #Drop unique trip IDs (i.e. single stop trips) and output to the final location, and get a unique list of stop IDs
    reduced_df = df[df.duplicated(subset = ['trip_id'], keep = False)].reset_index(drop = True)
    reduced_df.to_csv(os.path.join(path, 'output\\GTFS\\stop_times.txt'), index = False)

    print('Done')
    
if __name__ == "__main__":
    main()