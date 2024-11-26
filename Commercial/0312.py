# from datetime import datetime, timedelta
# today = datetime.today()
# last_friday = today - timedelta(today.weekday()+3)
# print(last_friday)

import os
import pandas as pd

weekly_tracker_path = r'C:\Users\he kelly\Desktop\FRC Tracker\Final version every week'

all_files = os.listdir(weekly_tracker_path)

weekly_trackers = [f for f in all_files if f.endswith('.xlsx')]
columns_extract = ['Request Number', 'Proposal Code', 'Proposal Start', 'Proposal End']
combined_trackers = pd.DataFrame(columns=columns_extract)

i=0
for tracker in weekly_trackers:
    i += 1
    tracker_path = os.path.join(weekly_tracker_path, tracker)
    df_tracker = pd.read_excel(tracker_path, sheet_name='Request in Proposal Ease', header=0)
    column_start = df_tracker.filter(regex='Proposal Start').columns.values
    column_end = df_tracker.filter(regex='Proposal End').columns.values
    df_tracker = df_tracker.loc[:, ['Request Number', 'Proposal Code', column_start[0], column_end[0]]]
    # df_tracker = df_tracker.dropna()
    print(i)
    print(column_start)
    print(column_end)
    print(df_tracker)
    # df_tracker.to_excel(rf'C:\Users\he kelly\Desktop\FRC Tracker\test_{i}.xlsx')
    combined_trackers = pd.concat([combined_trackers, df_tracker])

# print(combined_trackers)

# combined_trackers.to_excel(r'C:\Users\he kelly\Desktop\FRC Tracker\test_combined.xlsx')

