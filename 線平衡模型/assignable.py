import pulp as pl
import pandas as pd

t = ['T1', 'T2', 'T3', 'T4', 'T5']
s = ['s1', 's2', 's3', 's4', 's5']
r = ['R1', 'R2', 'R3', 'R4', 'R5', 'R6']

task_times = {'T1': 10, 'T2': 13, 'T3': 20, 'T4': 26, 'T5': 50}
task_times2 = {'R1': 5, 'R2': 3, 'R3': 7, 'R4': 12, 'R5': 9, 'R6': 5}
number = {'s1': 1, 's2': 2, 's3': 3, 's4': 4, 's5': 5}
num_stations = 5
avg_load = (sum(task_times.values()) + sum(task_times2.values())) / num_stations
model = pl.LpProblem("WorkstationBalancing", pl.LpMinimize)
load = pl.LpVariable.dicts("load", s, lowBound=0, cat=pl.LpContinuous)
abs_diff = pl.LpVariable.dicts("abs_diff", s, lowBound=0, cat=pl.LpContinuous)
task_assignment = pl.LpVariable.dicts("task_assignment", [(ti, si) for ti in t for si in s], cat=pl.LpBinary)
task_assignment2 = pl.LpVariable.dicts("task_assignment2", [(ri, si) for ri in r for si in s], cat=pl.LpBinary)
model += pl.lpSum([abs_diff[si] for si in s]), "TotalDeviation"
for si in s:
    model += load[si] == (
        pl.lpSum(task_assignment[(ti, si)] * task_times[ti] for ti in t) +
        pl.lpSum(task_assignment2[(ri, si)] * task_times2[ri] for ri in r)
    )
    model += abs_diff[si] >= load[si] - avg_load
    model += abs_diff[si] >= avg_load - load[si]
for ti in t:
    model += pl.lpSum(task_assignment[(ti, si)] for si in s) == 1
for ri in r:
    model += pl.lpSum(task_assignment2[(ri, si)] for si in s) == 1
for si in s:
    model += pl.lpSum(task_assignment[(ti, si)] for ti in t) == 1
for si in s:
    model += pl.lpSum(task_assignment2[(ri, si)] for ri in r) >= 0
model += (
    pl.lpSum(task_assignment[('T1', si)] * number[si] for si in s)
    <= pl.lpSum(task_assignment[('T5', si)] * number[si] for si in s)
)
model += (
    pl.lpSum(task_assignment[('T5', si)] * number[si] for si in s)
    <= pl.lpSum(task_assignment2[('R1', si)] * number[si] for si in s)
)
model += (
    pl.lpSum(task_assignment2[('R3', si)] * number[si] for si in s)
    <= pl.lpSum(task_assignment2[('R2', si)] * number[si] for si in s)
)
model += task_assignment[('T1', 's3')] == 1
model.solve()

station_tasks = {si: [] for si in s}
for ti in t:
    for si in s:
        if task_assignment[(ti, si)].varValue == 1:
            station_tasks[si].append(ti)
for ri in r:
    for si in s:
        if task_assignment2[(ri, si)].varValue == 1:
            station_tasks[si].append(ri)
station_task_df = pd.DataFrame([
    {"工作站": si, "任務": ", ".join(sorted(station_tasks[si]))}
    for si in s
])

print("目標式 (Total Deviation):", pl.value(model.objective))
print("\n工作站負荷:")
for si in s:
    print(f"{si}: {load[si].varValue:.2f} (abs diff: {abs_diff[si].varValue:.2f})")
print("\nT任務>S工作站(Main Tasks):")
for ti in t:
    for si in s:
        if task_assignment[(ti, si)].varValue == 1:
            print(f"{ti} -> {si}")
print("\nR任務>S工作站(Secondary Tasks):")
for ri in r:
    for si in s:
        if task_assignment2[(ri, si)].varValue == 1:
            print(f"{ri} -> {si}")
print(station_task_df.to_string(index=False))

