import pandas as pd

data = pd.read_excel('file.xlsx', index_col=0)
data.dropna(axis=1)

data['Time'] = pd.to_datetime(data['Time'])
data['Time Out'] = pd.to_datetime(data['Time Out'])

data = data.sort_values(by=['Employee Name', 'Time'])

consecutive_days = 1
prev_employee = None
prev_time_out = None

ans1 = []
ans2 = []
ans3 = []

for index, row in data.iterrows():
    if prev_employee == row['Employee Name']:
        # Calculate the time between the end of the previous shift and the start of the current shift
        time_between_shifts = (row['Time'] - prev_time_out).total_seconds() / 3600

        if time_between_shifts < 10 and time_between_shifts > 1:
            ans2.append(f"{row['Employee Name']} has less than 10 hours between shifts on {row['Time']}.")

        if (row['Time Out'] - row['Time']).total_seconds() / 3600 > 14:
            ans3.append(f"{row['Employee Name']} worked for more than 14 hours on {row['Time']}.")

        consecutive_days += 1
    else:
        consecutive_days = 1

    prev_employee = row['Employee Name']
    prev_time_out = row['Time Out']

    if consecutive_days == 7:
        ans1.append(f"{row['Employee Name']} has worked for 7 consecutive days.")


print("a) who has worked for 7 consecutive days.")
for employee in ans1:
    print(employee)

print("\n")

print("b) who have less than 10 hours of time between shifts but greater than 1 hour")
for employee in ans2:
    print(employee)

print("\n")

print("c) Who has worked for more than 14 hours in a single shift")
for employee in ans3:
    print(employee)