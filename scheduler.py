#####################################################
#       TEAM NUMBER: 10                             #
#       TEAM MEMBER: Lingchao Mao                   #
#                    YuQuan Cui                     #
#                    Ziwei Liu                      #
#                    Wenting Zheng                  #
#####################################################

import numpy as np
from datetime import datetime, timedelta
import operator
from sklearn import linear_model, preprocessing
import pandas as pd


#####################################################
#         This part is to load testing Data         #
#####################################################
test = pd.read_excel(r'Tasks_test.xlsx')
testDF = pd.DataFrame(data=test)
test_column_list = testDF.columns.values.tolist()
test_x_array = np.array(testDF[ [test_column_list[3], test_column_list[4], test_column_list[5], test_column_list[6] ] ] )

############################################
#  initialize returned duration dictionary #
############################################
duration = {}
test_id_array = testDF['ID']
length = len(test_id_array)

for i in range(length):
    duration[i] = []

test_category_array = testDF['Category']
category_num = []

# convert the category string into numbers
for category in test_category_array:
    if category == "School":
        category_num.append(2)
    elif category == "Personal":
        category_num.append(1)
    elif category == "Health":
        category_num.append(4)
    elif category == "Work":
        category_num.append(3)

for i in range(length):
    duration[i] = [ category_num[i] ]

list = []
my_array = np.array([])
for key in duration:
    key_value = duration[key] #a list
    value = key_value[0]
    #print(value)
    list.append(value)
my_array = np.array(list)

testDF[ test_column_list[2] ] =  my_array

test_x_array = np.array(testDF[ [test_column_list[2], test_column_list[3], test_column_list[4], test_column_list[5], test_column_list[6] ] ] )
test_x_array_normalized = preprocessing.normalize(test_x_array)
test_x_array_normalized_df = pd.DataFrame(test_x_array_normalized)


#####################################################
#      This part is to train model with history     #
#####################################################
task = pd.read_excel(r'Tasks_Training.xlsx')
taskDF = pd.DataFrame(data=task)
column_list = taskDF.columns.values.tolist()
###############
# build model #
###############
x_array = np.array(taskDF[ [ column_list[10], column_list[3], column_list[4], column_list[5], column_list[6] ] ] )
x_array_normalized = preprocessing.normalize(x_array)


x_array_normalized_df = pd.DataFrame(x_array_normalized)
y_df = taskDF[ [column_list[7]] ]

### Return a dict for each category ###
multi_regr = linear_model.LinearRegression()
multi_regr.fit(x_array_normalized_df, y_df)
predict_y = multi_regr.predict(test_x_array_normalized)
predict_y_list = predict_y.tolist()

for key in duration:
    value_list = duration[key]
    value_list.append(predict_y_list[key][0])
# print(duration)


#################################
#     This function is to       #
#     get the free blocks       #
#################################
def get_free_blocks(lines, schedule):
    for line in lines:
        line = line.replace(" ", "").split("\t")
        for days in line[9].replace("\n", ""):
            for val in schedule[days]:
                sta = datetime.strptime(val[0].strftime("%m/%d/%y ")+line[5], "%m/%d/%y %H:%M")
                end = datetime.strptime(val[0].strftime("%m/%d/%y ")+line[7], "%m/%d/%y %H:%M")
#             working interval = [A,B], event time = [c,d]
#             events in the block: A < c < d < B
                if val[0].time() <= sta.time() and val[1].time() >= end.time():
                    schedule[days].append([end, val[1]])
                    val[1] = val[1].replace(hour=sta.hour, minute=sta.minute)
#             overlap case 1: c < A < d < B
                if val[0].time() > sta.time() and val[0].time() < end.time():
                    val[0] = val[0].replace(hour=end.hour, minute=end.minute)
#             overlap case 2: A < c < B < d
                if val[1].time() > sta.time() and val[1].time() < end.time():
                    val[1] = val[1].replace(hour=sta.hour, minute=sta.minute)
    return schedule
    

#####################################################
#       This part is to read the events file        #
#####################################################
# Reading class schedule in the txt file
file = open("Events_test.txt", "r")
# Read in schedule in an array of dictionary
#  get next week's start date and end date
today = datetime.now().date()
start = today - timedelta(days=today.weekday()-7)
start_time = datetime.strptime(input("Enter the time you prefer to start the work(eg. 13:30): "), "%H:%M").replace(year=start.year, month = start.month, day = start.day)
end_time = datetime.strptime(input("Enter the time you prefer to end the work(eg. 13:30): "), "%H:%M").replace(year=start.year, month = start.month, day = start.day)
free_time = {
    "M":[[start_time, end_time]],
    "T":[[start_time + timedelta(days=1), end_time + timedelta(days=1)]],
    "W":[[start_time + timedelta(days=2), end_time + timedelta(days=2)]],
    "H":[[start_time + timedelta(days=3), end_time + timedelta(days=3)]],
    "F":[[start_time + timedelta(days=4), end_time + timedelta(days=4)]],
}
free_time = get_free_blocks(file, free_time)
file.close()

#################################
#     This function is to get   #
#     duration time             #
#################################
def getduration(times):
    diff = times[1] - times[0]
    return diff.total_seconds()/3600

#################################
#     This function is to get   #
#     possible event options    #
#################################
def get_Event_Options(slot, tasks):
    result = {
        5:[],
        4:[],
        3:[],
        2:[],
        1:[]
    }
    for k, task in tasks.items():
        ratio = task[1]/getduration(slot)
        if slot[0] < task[3] and ratio <= 1:
            result[task[2]].append([k, ratio, task[4]])
#     sort based on ratio
    for k, val in result.items():
        if val:
            result[k] = sorted(result[k], key=lambda x: -x[1])
    return result

#################################
#     This function is to       #
#     select events             #
#################################
def choose(tasks):
    result = []
    current_ratio = 0
    for k, tasks in tasks.items():
        if tasks:
            for task in tasks:
                if current_ratio + task[1] <= 1:
                    result.append(task)
                    current_ratio = current_ratio + task[1]
    return result

#################################
#     This function is to       #
#     print time and tasks      #
#################################
def printTimeTask(time, tasks):
    print(time[0].strftime("%m/%d/%y %H:%M-"), time[1].strftime("%H:%M tasks: "), end="")
    for task in tasks:
        print("( id =",task[0], ")", task[2], end=", ")
    print()

#####################################################
#  This part is to put the tasks in the free blocks #
#####################################################
tasks = np.array(testDF[ [test_column_list[4], test_column_list[7], test_column_list[1] ] ] )
i = 0

# id => [category#, duration, priority, due time, name]
for k,v in duration.items():
    tasks[i][1] = tasks[i][1].to_pydatetime()
    duration[k].extend(tasks[i])
    i = i + 1

i = 0
for time_slots in free_time.values():
    for slot in time_slots: #for each free time block
        possible_tasks = get_Event_Options(slot, duration)
        selected_tasks = choose(possible_tasks)
        if selected_tasks:
            for t in selected_tasks:
                del duration[t[0]]
            printTimeTask(slot,selected_tasks)
        i = i + 1
if duration:
    for k,v in duration.items():
        print("Task: ", v[4], "cannot fit in preferred working block. You have to use your leisure time!")