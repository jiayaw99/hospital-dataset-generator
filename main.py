import numpy as np
import random as rand
import pandas as pd
disease =  {1: "A",2: "B",3: "C",4: "D",5: "E",6:"F"}

def RemovePatientWithDoctor(current_patient_with_doctor):
    if current_patient_with_doctor !=0 and current_patient_with_doctor[3]==0:
        index=current_patient_with_doctor[0]
        return index,True
    else:
        return 0,False

def CheckNoShow(arrival_index, current_waiting_patient):
    random_noshow_rate=rand.randrange(20)/100  #15
    index = current_waiting_patient[0][0]
    if rand.random() < random_noshow_rate and index!=arrival_index:
        return index
    else:
        return 0

def CallingForNoshow(noshow_trial):
    call_success_rate=rand.randrange(30)/100   #30
    if rand.random() < call_success_rate and noshow_trial!=0:
        return False
    else:
        noshow_trial += 1
        return True  #not answering the call

def AdjustWidth(df,colName,writer):
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets[colName].set_column(col_idx, col_idx, column_width)

writer = pd.ExcelWriter('details.xlsx')
writer2 = pd.ExcelWriter('summary.xlsx')

df3 = pd.DataFrame(columns=['Day', 'Doctor number','Patient', 'Gender', 'Age', 'Service time', 'Waiting time', 'Arrival time',
                            'Queue Size When Arrived',"Disease type","Disease condition","Priority Index"])

doctors_number=[]
for i in range(50): #doctor number in 30 days
    doctor_prob = rand.randrange(1000)
    doctor_prob = 5 if doctor_prob > 950 else (4 if doctor_prob > 500 else (3 if doctor_prob > 50 else 2))
    doctors_number.append(doctor_prob)
largest_doctor = max(doctors_number)

df4 = pd.DataFrame(
    columns=['Day', 'Patient average waiting times', 'No show cases',
             'No show rate', 'Emergency case', 'Emergency rate'])
df4_temp= pd.DataFrame()
initial =[]
for i in range(largest_doctor):
   df4.insert(1+i, 'Doctor '+str(i+1)+ " Idle Rate",None)

for days in range(50): # 30 days
    doctor_number=doctors_number[days]
    df = pd.DataFrame(columns=['Clock Size', 'Queue Size', 'Event'])
    for i in range(doctor_number):
        df.insert(2 + i, 'Doctor '+str(i+1)+' Status', None)

    df2 = pd.DataFrame(columns=['Special Case', 'Patient'])
    clocksize = 480  # 8 hours
    current_clock = 0
    all_patient = []
    current_waiting_patient = []
    current_patient_with_doctor = [0]*doctor_number
    calling_patient = [0]*doctor_number
    emergency_patient = []
    emergency_record = []
    no_show_record = []
    doctor_idle_count = [0]*doctor_number
    noshow_trial = [0]*doctor_number
    calling = [False]*doctor_number
    index_noshow = [-1]*doctor_number
    Doctor_Status = [0]*doctor_number
    index = [0]*doctor_number

    start_queue = int(doctor_number * (rand.randrange(20) + 30)/5)
    for i in range(start_queue):
        new_patient = [len(all_patient) + 1, rand.randrange(2), rand.randrange(10, 60),
                       int(np.random.normal(20, 3, 1)[0]), 0, -1,
                       start_queue,
                       rand.randrange(5)+1,rand.randrange(3)]  # patient index, gender, age, service time, waiting time, arrival time, queue size, disease type, disease condition

        priority=new_patient[2]/60+new_patient[7]+new_patient[8]
        new_patient.append(1 if priority < 7 else (2 if priority < 8 else 3)) # priority index
        all_patient.append(new_patient.copy())
        arrival_index = new_patient[0]
        current_waiting_patient.append(new_patient)
        current_waiting_patient.sort(key=lambda x: x[9],reverse=True)

    #print('Clock size    Queue Size          Doctor 1 Status                      Doctor 2 Status                        Event')
    if (start_queue != 0):
        initial =[0, len(current_waiting_patient) , str(doctor_number) + " doctors oncall today, "+ str(start_queue) + " patient reached before hospital open" ]
        for i in range(doctor_number):
            initial.insert(2,'None')
        a = pd.Series(initial, index=df.columns)
        df = df.append(a, ignore_index=True)
    #     print('  %-4s \t\t  %5d \t\t %25s \t\t\t %25s \t\t\t %-20s' % (
    #     "None", len(current_waiting_patient), "None", "None",
    #     str(start_queue) + " patient reached before hospital open"))
    while current_clock < clocksize or any(v != 0 for v in current_patient_with_doctor) == True \
            or len(current_waiting_patient) != 0 or any(v != 0 for v in calling_patient) == True:
        # print(current_patient_with_doctor)
        # print(calling_patient)

        arrival = False
        arrival_index = 0
        left = [False]*doctor_number
        emergency_index = 0
        emergency = False
        if (current_clock <= clocksize):
            random_arrival_rate = doctor_number * rand.randrange(10) / 100
        else:
            random_arrival_rate = 0
        if rand.random() < random_arrival_rate:
            new_patient = [len(all_patient) + 1, rand.randrange(2), rand.randrange(10, 60),
                           int(np.random.normal(20, 3, 1)[0]), 0, current_clock,
                           0, rand.randrange(5)+1,rand.randrange(3)]  # patient index, gender, age, service time, waiting time, arrival time, queue size
            priority = new_patient[2] / 60 + new_patient[7] + new_patient[8]
            new_patient.append(1 if priority < 7 else (2 if priority < 8 else 3))  # priority index
            all_patient.append(new_patient.copy())
            arrival = True
            arrival_index = new_patient[0]
            current_waiting_patient.append(new_patient)
            current_waiting_patient.sort(key=lambda x: x[9], reverse=True)

        if rand.random() < 0.002:  # emergency case
            new_patient = [len(all_patient) + 1, rand.randrange(2), rand.randrange(10, 60),
                           int(np.random.normal(40, 2, 1)[0]), 0, current_clock, -1,
                           6,2,5]
            all_patient.append(new_patient.copy())
            emergency = True
            emergency_index = new_patient[0]
            emergency_patient.append(new_patient)
            emergency_record.append(emergency_index)

        for i in range(len(current_patient_with_doctor)):
            index[i], left[i] = RemovePatientWithDoctor(current_patient_with_doctor[i])
            if (left[i]):
                all_patient[index[i] - 1][4] = current_patient_with_doctor[i][4]
                current_patient_with_doctor[i] = 0

        if len(emergency_patient) != 0:
            for i in range(len(current_patient_with_doctor)):
                if current_patient_with_doctor[i]==0:
                    current_patient_with_doctor[i] = emergency_patient.pop(0)
                    break

        for i in range(doctor_number):
            if calling[i]:
                calling[i] = CallingForNoshow(noshow_trial[i])
                if (calling[i] and noshow_trial[i] == 5):
                    all_patient[index_noshow[i] - 1][3] = 0  # service time equal to 0 if no show
                    all_patient[index_noshow[i] - 1][4] = calling_patient[i][4]
                    calling_patient[i]=0
                    no_show_record.append(index_noshow[i])
                    if len(current_waiting_patient) > 0 and current_patient_with_doctor[i] == 0:
                        current_patient_with_doctor[i] = current_waiting_patient.pop(0)
                elif calling[i] == False:
                    noshow_trial[i] = 0
                    if current_patient_with_doctor[i] == 0:
                        current_patient_with_doctor[i] = calling_patient[i]
                    else:
                        current_waiting_patient.insert(0, calling_patient[i])
                    calling_patient[i]=0

            elif len(current_waiting_patient) != 0:
                if current_patient_with_doctor[i] == 0:
                    index_noshow[i] = CheckNoShow(arrival_index, current_waiting_patient)  ##prevent no-show happen at the arrival time
                    if (index_noshow[i]) != 0:
                        calling[i] = CallingForNoshow(noshow_trial[i])
                        calling_patient[i]=current_waiting_patient.pop(0)
                    else:
                        current_patient_with_doctor[i] = current_waiting_patient.pop(0)
        #
        # if calling2:
        #     calling2 = CallingForNoshow(noshow_trial2)
        #     if (calling2 and noshow_trial2 == 5):
        #         all_patient[index_noshow2 - 1][3] = 0
        #         all_patient[index_noshow2 - 1][4] = calling_patient2.pop(0)[4]
        #         no_show_record.append(index_noshow2)
        #         if len(current_waiting_patient) > 0 and len(current_patient_with_doctor2) == 0:
        #             current_patient_with_doctor2 = current_waiting_patient.pop(0)
        #     elif calling2 == False:
        #         noshow_trial2 = 0
        #         if (len(current_patient_with_doctor2) == 0):
        #             current_patient_with_doctor2 = calling_patient2.pop(0)
        #         else:
        #             current_waiting_patient.insert(0, calling_patient2.pop(0))
        # elif len(current_waiting_patient) != 0:
        #     if len(current_patient_with_doctor2) == 0 and calling2 == False:
        #         index_noshow2 = CheckNoShow(arrival_index, current_waiting_patient)
        #         if (index_noshow2) != 0:
        #             calling2 = CallingForNoshow(noshow_trial2)
        #             calling_patient2.append(current_waiting_patient.pop(0))
        #         else:
        #             current_patient_with_doctor2 = current_waiting_patient.pop(0)

        if len(current_waiting_patient) != 0:
            for patient in current_waiting_patient:
                patient[4] += 1
        for i in range(len(calling_patient)):
            if calling_patient[i] != 0:
                calling_patient[i][4] += 1
        if len(emergency_patient) != 0:
            for patient in emergency_patient:
                patient[4] += 1

        for i in range(len(current_patient_with_doctor)):
            if current_patient_with_doctor[i] != 0:
                current_patient_with_doctor[i][3] -= 1

        for i in range(doctor_number):
            Doctor_Status[i] = "Occupied with patient " + str(current_patient_with_doctor[i][0]) if current_patient_with_doctor[i] != 0 else str("Idle")
            if (Doctor_Status[i] == "Idle"):
                doctor_idle_count[i] += 1

        Event = ""
        if arrival:
            all_patient[arrival_index-1][6]=len(current_waiting_patient)
            Event += "New patient " + str(arrival_index) + " arrived  "
        if emergency:
            Event += "New emergency patient " + str(emergency_index) + " arrived  "
        for i in range(doctor_number):
            if left[i]:
                Event += "Patient " + str(index[i]) + " left  "
            if noshow_trial[i] < 5 and calling[i] == True:
                noshow_trial[i] += 1
                Event += "Calling for Patient " + str(index_noshow[i]) + " " + str(noshow_trial[i]) + " time(s) "
            elif noshow_trial[i] == 5 and index_noshow[i] != 0 and calling[i]:
                noshow_trial[i] = 0
                calling[i] = False
                calling_patient[i]=0
                Event += "Patient " + str(index_noshow[i]) + " no show "

        #print('%4s \t\t  %5d \t\t %25s \t\t\t %25s \t\t\t %-20s' % (
        #current_clock, len(current_waiting_patient), Doctor1_Status, Doctor2_Status, Event))

        current_clock += 1

        if len(current_waiting_patient) > 10 and current_clock > 360:  # do not accept patient anymore if too much queue near the closing hour
            clocksize = 360

        temp_list=[current_clock, len(current_waiting_patient) + len(emergency_patient), Event]
        for i in range(doctor_number):
            temp_list.insert(2+i,Doctor_Status[i])
        a = pd.Series(temp_list, index=df.columns)
        df = df.append(a, ignore_index=True)

    df.to_excel(writer, sheet_name='day '+str(days+1)+' details', index=False, na_rep='NaN')

    # Auto-adjust columns' width
    AdjustWidth(df,'day '+str(days+1)+' details',writer)

    for i in range(len(no_show_record)):
        a = pd.Series(["No-show  " + str(i + 1), str(no_show_record[i]) ], index=df2.columns)
        df2 = df2.append(a, ignore_index=True)
    for i in range(len(emergency_record)):
        a = pd.Series(["Emergency Case " + str(i + 1), str(emergency_record[i]) ], index=df2.columns)
        df2 = df2.append(a, ignore_index=True)
    df2.to_excel(writer, sheet_name='day '+str(days+1)+' special cases', index=False, na_rep='NaN')
    AdjustWidth(df2,'day '+str(days+1)+' special cases',writer)

    # print('\nPatient       Gender        Age    Service time   Waiting time   Arrival time   Queue Size When Arrival')
    # for i in range(len(all_patient)):
    #     print(' %2s \t\t  %5s \t %5s \t\t %5s \t\t  %5s  \t\t %5s  \t\t %6s' % (
    #     all_patient[i][0], "male" if all_patient[i][1] == 0 else "female", all_patient[i][2], all_patient[i][3],
    #     all_patient[i][4], all_patient[i][5], all_patient[i][6]))

    for i in range(len(all_patient)):
        a = pd.Series([days+1,doctor_number,all_patient[i][0], "male" if all_patient[i][1] == 0 else "female", all_patient[i][2], all_patient[i][3],
        all_patient[i][4], all_patient[i][5], all_patient[i][6],disease[all_patient[i][7]],all_patient[i][8],all_patient[i][9]], index=df3.columns)
        df3 = df3.append(a, ignore_index=True)
    temp_b = [days + 1,
                   round(sum(np.array(all_patient)[:, 4]) / len(all_patient), 2),
                   len(no_show_record), round(100 * len(no_show_record) / len(all_patient), 4), len(emergency_record),
                   round(100 * len(emergency_record) / len(all_patient), 4)]
    for i in range(doctor_number):
        temp_b.insert(1+i,round(100 * doctor_idle_count[i] / current_clock, 2))
    if (largest_doctor-doctor_number)!=0:
        for i in range((largest_doctor-doctor_number)):
            temp_b.insert(1+doctor_number+i,'None')

    b = pd.Series(temp_b, index=df4.columns)
    df4 = df4.append(b, ignore_index=True)

    # print("\nDoctor 1 Idle Rate: " + str(
    #     round(100 * doctor1_idle_count / current_clock, 2)) + "   Doctor 2 Idle Rate: " + str(
    #     round(100 * doctor2_idle_count / current_clock, 2))
    #       + "   Patient average waiting times: " + str(round(sum(np.array(all_patient)[:, 4]) / len(all_patient), 2)))
    # print()
    # for i in range(len(no_show_record)):
    #     print("No-show Case " + str(i + 1) + ": Patient " + str(no_show_record[i]))
    # print()
    # for i in range(len(emergency_record)):
    #     print("Emergency Case " + str(i + 1) + ": Patient " + str(emergency_record[i]))

df3.to_excel(writer2, sheet_name='patient record', index=False, na_rep='NaN')
AdjustWidth(df3, 'patient record',writer2)

df4.to_excel(writer2, sheet_name='summary', index=False, na_rep='NaN')
AdjustWidth(df4, 'summary',writer2)

writer.save()
writer2.save()
