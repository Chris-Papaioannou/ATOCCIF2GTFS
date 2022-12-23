import os
import pandas as pd

path = os.path.dirname(__file__)

df = pd.read_csv(os.path.join(path, 'temp\\stop_times_full.txt'), low_memory = False)

report = df['trip_id'].drop_duplicates(keep = False)

print(str(len(report)) + ' trips were dropped as they only had one stop. The following trip IDs were affected:')
print(report)

reduced_df = df[df.duplicated(subset = ['trip_id'], keep = False)]

reduced_df.to_csv(os.path.join(path, 'output\\stop_times.txt'), index = False)