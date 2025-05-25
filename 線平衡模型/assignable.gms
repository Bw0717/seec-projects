Sets
    t   / T1*T5 /
    s   / s1*s5 /        
    r   / R1*R6 /
;
Parameters
    task_times(t)    / T1 10, T2 13, T3 20, T4 26, T5 50 /  
    task_times2(r)   / R1 5, R2 3, R3 7, R4 12, R5 9, R6 5/
    number(s)        / s1 1, s2 2, s3 3, s4 4, s5 5/
;
Scalar
    num_stations   / 5 /;       

Variable
    load(s)           ! 每個工作站的總負荷
    abs_diff(s)       ! 每個工作站的負載與平均負載的差異絕對值
    obj               ! 目標：所有工作站差異總和最小化

Binary Variables
    task_assignment(t, s)    ！固定任務
    task_assignment2(r, s)   ！彈性任務 ;

Equations
    load_eq(s)              ! 總負荷計算
    abs_diff_pos_eq(s)      ! 絕對值
    abs_diff_neg_eq(s)      ! 絕對值
    obj_eq                  ! 目標式
    task_assignment_eq(t)   ! 固定任務雙生限制式
    task_per_station_eq(s)  ! 固定任務雙生限制式
    task_assignment_eq2(r)  ! 彈性任務雙生限制式
    task_per_station_eq2(s) ! 彈性任務雙生限制式
    task_order_constraint   ! 排程順序關係式T-T
    task_order_constraint2  ! 排程順序關係式R-T
    task_order_constraint3  ! 排程順序關係式R-R
    lock_order(s)           ! 指定任務
;

* Average load
Scalar avg_load;
avg_load = (sum(t, task_times(t)) + sum(r, task_times2(r))) / num_stations;


load_eq(s).. 
    load(s) =e= sum(t, task_assignment(t, s) * task_times(t)) + sum(r, task_assignment2(r, s) * task_times2(r));

abs_diff_pos_eq(s).. 
    abs_diff(s) =g= load(s) - avg_load;

abs_diff_neg_eq(s).. 
    abs_diff(s) =g= avg_load - load(s);

obj_eq.. 
    obj =e= sum(s, abs_diff(s));

task_assignment_eq(t).. 
    sum(s, task_assignment(t, s)) =e= 1;

task_per_station_eq(s).. 
    sum(t, task_assignment(t, s)) =e= 1;  

task_assignment_eq2(r).. 
    sum(s, task_assignment2(r, s)) =e= 1; 

task_per_station_eq2(s)..
    sum(r,task_assignment2(r,s)) =g= 0;
    
task_order_constraint..
    sum(s, task_assignment('T1', s) * number(s)) =l= sum(s, task_assignment('T5', s) * number(s));

task_order_constraint2..
    sum(s, task_assignment('T5', s) * number(s)) =l= sum(s, task_assignment2('R1', s) * number(s));
    
task_order_constraint3..
    sum(s, task_assignment2('R3', s) * number(s)) =l= sum(s, task_assignment2('R2', s) * number(s));
    
lock_order(s).. 
    task_assignment('T1', 's3') =e= 1;

Model WorkstationBalancing /all/;


Solve WorkstationBalancing using mip minimizing obj;


Display load.l, abs_diff.l, obj.l;
Display task_assignment.l, task_assignment2.l;
