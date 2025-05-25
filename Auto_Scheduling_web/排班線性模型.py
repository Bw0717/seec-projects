from pulp import LpMaximize, LpProblem, LpVariable, lpSum


employees = ["E1", "E2", "E3","E4","E5","E6","overpeople"]
stations = ["S1", "S2", "S3", "S4", "S5"] 
capabilities = {
    ("E1", "S1"): 1, ("E1", "S2"): 1, ("E1", "S3"): 1, ("E1", "S4"):1, ("E1", "S5"):1,
    ("E2", "S1"): 1, ("E2", "S2"): 1, ("E2", "S3"): 1, ("E2", "S4"):1, ("E2", "S5"):1,
    ("E3", "S1"): 1, ("E3", "S2"): 1, ("E3", "S3"): 1, ("E3", "S4"):1, ("E3", "S5"):1,
    ("E4", "S1"): 1, ("E4", "S2"): 1, ("E4", "S3"): 1, ("E4", "S4"):1, ("E4", "S5"):1,
    ("E5", "S1"): 1, ("E5", "S2"): 1, ("E5", "S3"): 1, ("E5", "S4"):1, ("E5", "S5"):1,
    ("E6", "S1"): 1, ("E6", "S2"): 1, ("E6", "S3"): 1, ("E6", "S4"):1, ("E6", "S5"):1,        
    ("overpeople", "S1"): 0, ("overpeople","S2"): 1, ("overpeople", "S3"): 1,("overpeople", "S4"):1, ("overpeople", "S5"):1
}
preferences = {
    ("E1", "S1"): 1, ("E1", "S2"): 1, ("E1", "S3"): 1, ("E1", "S4"):1, ("E1", "S5"):1,
    ("E2", "S1"): 1, ("E2", "S2"): 1, ("E2", "S3"): 1, ("E2", "S4"):1, ("E2", "S5"):1,
    ("E3", "S1"): 1, ("E3", "S2"): 1, ("E3", "S3"): 1, ("E3", "S4"):1, ("E3", "S5"):1,
    ("overpeople", "S1"): -100, ("overpeople","S2"): -100, ("overpeople", "S3"): -100,("overpeople", "S4"):-100, ("overpeople", "S5"):-100
}
rewards = {"S1": 30, "S2": 30, "S3": 30, "S4":0, "S5":0}
priority_stations = {"S1": 30, "S3": 30,"S2":30}  

for e in employees:
    for s in stations:
        preferences.setdefault((e, s), 0)


model = LpProblem(name="scheduling", sense=LpMaximize)


x = LpVariable.dicts("assign", [(e, s) for e in employees for s in stations], cat="Binary")


model += lpSum(preferences[e, s] * x[e, s] for e in employees for s in stations) + \
         lpSum(rewards[s] * lpSum(x[e, s] for e in employees if e != "overpeople")  for s in stations) + \
         lpSum(priority_stations.get(s, 0) * lpSum(x[e, s] for e in employees if e != "overpeople") for s in stations)


for e in employees :
    if e != "overpeople":
        model += lpSum(x[e, s] for s in stations) == 1


for s in stations:
    if s != "S2":
        model += lpSum(x[e, s] for e in employees) == 1




for e in employees:
    for s in stations:
        model += x[e, s] <= capabilities[e, s]  



model.solve()


print("Optimal Schedule:")
for e in employees:
    for s in stations:
        if x[e, s].value() == 1:
            print(f"Employee {e} assigned to Station {s}")

print("\nTotal Preference and Reward Value:", model.objective.value())
