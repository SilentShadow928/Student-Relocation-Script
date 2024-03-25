print("Please wait...")
import msvcrt # For instant input
import random
import pandas # To parse excel file
import tkinter # To get file graphically. If anyone want to make a full GUI, go ahead!
from tkinter.filedialog import askopenfilename
#Tk().withdraw()

max_in_group = input("What is the maximum ammount of students in a group? ")
while not max_in_group.isnumeric():
    print("You answer was not a number.")
    max_in_group = input("What is the maximum ammount of students in a group? ")
max_in_group = int(max_in_group)

min_in_group = input("What is the minimum ammount of students in a group? ")
while not min_in_group.isnumeric():
    print("You answer was not a number.")
    min_in_group = input("What is the minimum ammount of students in a group? ")
min_in_group = int(min_in_group)
print()

window = tkinter.Tk()
window.wm_attributes('-topmost', 1)
window.withdraw()
filename = askopenfilename(parent=window, initialdir="", title="Select the Excel file with student preferences", filetypes = (("Excel Files", "*.xl*;*.xm*"), ("All files", "*")))


df = pandas.read_excel(filename)
student_preferences = {}
for index, row in df.iterrows():
    student_preferences[row.iloc[5]] = []
    for tmpVar7 in row[6:10]:
        if str(tmpVar7) != "nan":
            student_preferences[row.iloc[5]].append(tmpVar7)

#Functions:
def map_preferences(student_preferences, student):
    if student not in student_preferences:
        return
    preferences = student_preferences[student]
    for elem in preferences:
        if elem not in mapped_preferences:
            if elem in student_preferences:
                mapped_preferences.append(elem)
                map_preferences(student_preferences, elem)

    for elem in mapped_preferences:
        tmpVar1=-1
        for item in list(student_preferences.values()):
            tmpVar1 += 1
            if elem in item:
                if list(student_preferences.keys())[tmpVar1] not in mapped_preferences:
                    mapped_preferences.append(list(student_preferences.keys())[tmpVar1])
                    map_preferences(student_preferences, list(student_preferences.keys())[tmpVar1])
    for sublist in mapped_preferences:
        if student in sublist:
            return mapped_preferences
    return [student]

def list_groups(mapped_preferences_tmp, include_kicked=False, new_group_size = -1):
    lengths = []
    if include_kicked == True:
        lengths = [1]
    for elem in mapped_preferences_tmp:
        lengths.append(len(elem))
    lengths.sort()
    if new_group_size != -1:
        lengths.remove(new_group_size-1)
        lengths.append(new_group_size)
    first_iter = True
    lengths2 = list(set(lengths))
    lengths2.reverse()
    for elem in lengths2:
        if not first_iter:
            print(", ", end='')
            print(str(lengths.count(elem)) + " group(s) of " + str(elem), end='')
        else:
            first_iter = False
            print(str(lengths.count(elem)) + " group(s) of " + str(elem), end='')
    print()

def compute_dilike(student, moveTo = []):
    # Count max others dissaproval
    num_dislike = 0
    for elem in list(student_preferences.values()):
        if student in elem:
            num_dislike += 1
    # Count max student dissaproval
    student_dislike=len(student_preferences.get(student))
    if moveTo != []:
        # Subtract how many people in the new group want to sit with the student
        for j in moveTo:
            for elem in student_preferences.get(j):
                if elem == student:
                    num_dislike -= 1
        # Subtract how many people the student wants to sit with in the new group
        for j in moveTo:
            for elem in student_preferences.get(student):
                if elem == j:
                    student_dislike -= 1
    return [num_dislike, student_dislike]

all_mapped_preferences = []

for student in student_preferences:
    isStudentInList = False
    for sub_list in all_mapped_preferences:
        for element in sub_list:
            if student in sub_list:
                isStudentInList = True
    if not isStudentInList:
        mapped_preferences = []
        all_mapped_preferences.append(map_preferences(student_preferences, student))
all_mapped_preferences.sort(key=len)
#UI:
while True:
    print("Current groups: ", end='')
    list_groups(all_mapped_preferences)
    #all_mapped_preferences.sort(key=len)
    print("Use (e) to exit any submenus")
    print("Options: Combine groups (c)")
    print("Move/kick students      (m)")
    print("Automatically arrange   (a)")
    print("Sort by length          (s)")
    print("Print out groups        (p)")
    print("Reset                   (r)")
    print("Exit                    (e)")
    print("Press a button: ", end='', flush=True)
    choice=msvcrt.getwch()
    choice = choice.lower()
    print(str(choice))
    if choice == "e":
        print("Thank you for using the Student Relocation Script!")
        print("Have a nice day!")
        break
    elif choice == "p":
        tmpVar2=1
        for sub_list in all_mapped_preferences:
            print("Group " + str(tmpVar2) + ": " + str(sub_list))
            tmpVar2+=1
        print("Press any key to continue: ", end='', flush=True)
        msvcrt.getch()
        print()
    elif choice == "c":
        print("What groups would you like to combine?")
        tmpVar2=1
        for sub_list in all_mapped_preferences:
            print("Group " + str(tmpVar2) + ": " + str(sub_list))
            tmpVar2+=1
        print("Press a button: ", end='', flush=True)
        choice1=msvcrt.getwch()
        print(choice1)
        if choice1.isnumeric():
            choice1=int(choice1)
            print("Press a button: ", end='', flush=True)
            choice2=msvcrt.getwch()
            print(choice2)
            if choice2.isnumeric():
                choice2=int(choice2)
                all_mapped_preferences[choice1-1] = all_mapped_preferences[choice1-1][:] + all_mapped_preferences[choice2-1][:]
                del all_mapped_preferences[choice2-1]
                print("Combined!")
            else:
                print("Your input was not a number, exiting.")
        else:
            print("Your input was not a number, exiting.")
    elif choice == "m":
        tmpVar2=1
        for sub_list in all_mapped_preferences:
            print("Group " + str(tmpVar2) + ": " + str(sub_list))
            tmpVar2+=1
        print("What group do you want to seperate? ", end='', flush=True)
        choice1=msvcrt.getwch()
        print(choice1)
        if choice1.isnumeric():
            print("Group " + str(choice1) + ": " + str(all_mapped_preferences[choice1-1]))
            print("Possible options:", end='')
            times_printed = 1
            option_instructions=[]
            for student in all_mapped_preferences[choice1-1]:
                print()
                all_dislike1 = compute_dilike(student)
                num_dislike1 = all_dislike1[0]
                student_dislike1 = all_dislike1[1]
                students_in_group = {}
                for key, value in student_preferences.items():
                    if key in all_mapped_preferences[choice1-1] and key != student:
                        students_in_group[key] = value
                mapped_group_preferences = []
                for a in students_in_group:
                    isStudentInList = False
                    for sub_list in mapped_group_preferences:
                        for element in sub_list:
                            if a in sub_list:
                                isStudentInList = True
                    if not isStudentInList:
                        mapped_preferences = []
                        mapped_group_preferences.append(map_preferences(students_in_group, a))
                mapped_group_preferences.sort(key=len)
                print(str(times_printed) + ". Kick " + str(student) + ": DS: " + str(num_dislike1 + student_dislike1) + " (" + str(num_dislike1) + " others, " + str(student_dislike1) + " self), group seperated into ", end='')
                option_instructions.append(["K", student, choice1-1])
                times_printed += 1
                list_groups(mapped_group_preferences, include_kicked=True)
                all_mapped_group_preferences = []
                tmpVar4=0
                for i in all_mapped_preferences:
                    if tmpVar4 != choice1-1:
                        all_mapped_group_preferences.append(i)
                    tmpVar4 += 1
                #all_mapped_group_preferences.sort(key=len)
                for elem in mapped_group_preferences:
                    all_mapped_group_preferences.append(elem)
                #all_mapped_group_preferences.sort(key=len)
                tmpVar3=0
                tmpVar5 = list(all_mapped_preferences[choice1-1])
                tmpVar5.remove(student)

                for group in all_mapped_group_preferences:
                    if set(tmpVar5) != set(group):
                        all_dislike = compute_dilike(student, moveTo = group)
                        num_dislike = all_dislike[0]
                        student_dislike = all_dislike[1]
                        if (num_dislike + student_dislike) < (num_dislike1 + student_dislike1):
                            print(str(times_printed) + ". Move " + str(student) + " to (maybe new) group " + str(tmpVar3+1) + ": DS: " + str(num_dislike + student_dislike) + " (" + str(num_dislike) + " others, " + str(student_dislike) + " self), new group sizes: ", end='')
                            option_instructions.append(["M", student, choice1-1, tmpVar3])
                            times_printed += 1
                            #print(all_mapped_group_preferences)
                            list_groups(all_mapped_group_preferences, include_kicked=False, new_group_size=len(all_mapped_group_preferences[tmpVar3])+1)
                            #print(all_mapped_group_preferences)
                        tmpVar3 += 1
            option_chosen = input("Pick an option: ")
            if choice1.isnumeric():
                if option_instructions[option_chosen-1][0] == "K":
                    students_in_group = {}
                    for key, value in student_preferences.items():
                        if key in all_mapped_preferences[option_instructions[option_chosen-1][2]] and key != option_instructions[option_chosen-1][1]:
                            students_in_group[key] = value
                    mapped_group_preferences = []
                    for a in students_in_group:
                        isStudentInList = False
                        for sub_list in mapped_group_preferences:
                            for element in sub_list:
                                if a in sub_list:
                                    isStudentInList = True
                        if not isStudentInList:
                            mapped_preferences = []
                            mapped_group_preferences.append(map_preferences(students_in_group, a))
                    mapped_group_preferences.append([option_instructions[option_chosen-1][1]])
                    del all_mapped_preferences[option_instructions[option_chosen-1][2]]
                    for elem in mapped_group_preferences:
                        all_mapped_preferences.append(elem)
                    
                elif option_instructions[option_chosen-1][0] == "M":
                    students_in_group = {}
                    for key, value in student_preferences.items():
                        if key in all_mapped_preferences[option_instructions[option_chosen-1][2]] and key != option_instructions[option_chosen-1][1]:
                            students_in_group[key] = value
                    mapped_group_preferences = []
                    for a in students_in_group:
                        isStudentInList = False
                        for sub_list in mapped_group_preferences:
                            for element in sub_list:
                                if a in sub_list:
                                    isStudentInList = True
                        if not isStudentInList:
                            mapped_preferences = []
                            mapped_group_preferences.append(map_preferences(students_in_group, a))
                    #mapped_group_preferences.append([option_instructions[option_chosen-1][1]])
                    mapped_group_preferences.sort(key=len)
                    del all_mapped_preferences[option_instructions[option_chosen-1][2]]
                    for elem in mapped_group_preferences:
                        all_mapped_preferences.append(elem)
                    tmpVar6 = list(all_mapped_preferences[option_instructions[option_chosen-1][3]])
                    tmpVar6.append(option_instructions[option_chosen-1][1])
                    all_mapped_preferences[option_instructions[option_chosen-1][3]] = tmpVar6
            else:
                print("Your input was not a number, exiting.")
        else:
            print("Your input was not a number, exiting.")
    elif choice == "a":
        print("Warning: the output of this option may not be perfect. If you are unhappy with the result, you can reset and rerun this option")
        print("Processing...")
        random.shuffle(all_mapped_preferences)
        tmpVar8=0
        could_slice = True
        for group in all_mapped_preferences:
            if len(group) > max_in_group:
                option_instructions=[]
                for student in all_mapped_preferences[tmpVar8]:
                    all_dislike1 = compute_dilike(student)
                    num_dislike1 = all_dislike1[0]
                    student_dislike1 = all_dislike1[1]
                    students_in_group = {}
                    for key, value in student_preferences.items():
                        if key in all_mapped_preferences[tmpVar8] and key != student:
                            students_in_group[key] = value
                    mapped_group_preferences = []
                    for a in students_in_group:
                        isStudentInList = False
                        for sub_list in mapped_group_preferences:
                            for element in sub_list:
                                if a in sub_list:
                                    isStudentInList = True
                        if not isStudentInList:
                            mapped_preferences = []
                            mapped_group_preferences.append(map_preferences(students_in_group, a))
                    mapped_group_preferences.sort(key=len)
                    isGroupTooLarge = False
                    for elem in mapped_group_preferences:
                        if len(elem) > max_in_group:
                            isGroupTooLarge = True
                    if not isGroupTooLarge:
                        option_instructions.append(["K", student, tmpVar8, None, num_dislike1 + student_dislike1])
                        all_mapped_group_preferences = []
                        tmpVar4=0
                        for i in all_mapped_preferences:
                            if tmpVar4 != tmpVar8:
                                all_mapped_group_preferences.append(i)
                            tmpVar4 += 1
                        #all_mapped_group_preferences.sort(key=len)
                        for elem in mapped_group_preferences:
                            all_mapped_group_preferences.append(elem)
                        #all_mapped_group_preferences.sort(key=len)
                        tmpVar3=0
                        tmpVar5 = list(all_mapped_preferences[tmpVar8])
                        tmpVar5.remove(student)

                        for group in all_mapped_group_preferences:
                            if set(tmpVar5) != set(group) and len(group) + 1 <= max_in_group:
                                all_dislike = compute_dilike(student, moveTo = group)
                                num_dislike = all_dislike[0]
                                student_dislike = all_dislike[1]
                                option_instructions.append(["M", student, tmpVar8, tmpVar3, num_dislike + student_dislike])
                                tmpVar3 += 1
                if option_instructions != []:
                    random.shuffle(option_instructions)
                    min_ds = [0, option_instructions[0][4]]
                    tmpVar9=0
                    for elem in option_instructions:
                        if elem[4] < min_ds[1]:
                            min_ds = [tmpVar9, elem[4]]
                        tmpVar9 += 1
                    if option_instructions[min_ds[0]][0] == "K":
                        students_in_group = {}
                        for key, value in student_preferences.items():
                            if key in all_mapped_preferences[option_instructions[min_ds[0]][2]] and key != option_instructions[min_ds[0]][1]:
                                students_in_group[key] = value
                        mapped_group_preferences = []
                        for a in students_in_group:
                            isStudentInList = False
                            for sub_list in mapped_group_preferences:
                                for element in sub_list:
                                    if a in sub_list:
                                        isStudentInList = True
                            if not isStudentInList:
                                mapped_preferences = []
                                mapped_group_preferences.append(map_preferences(students_in_group, a))
                        mapped_group_preferences.append([option_instructions[min_ds[0]][1]])
                        del all_mapped_preferences[option_instructions[min_ds[0]][2]]
                        for elem in mapped_group_preferences:
                            all_mapped_preferences.append(elem)
                            
                    elif option_instructions[min_ds[0]][0] == "M":
                        students_in_group = {}
                        for key, value in student_preferences.items():
                            if key in all_mapped_preferences[option_instructions[min_ds[0]][2]] and key != option_instructions[min_ds[0]][1]:
                                students_in_group[key] = value
                        mapped_group_preferences = []
                        for a in students_in_group:
                            isStudentInList = False
                            for sub_list in mapped_group_preferences:
                                for element in sub_list:
                                    if a in sub_list:
                                        isStudentInList = True
                            if not isStudentInList:
                                mapped_preferences = []
                                mapped_group_preferences.append(map_preferences(students_in_group, a))
                        #mapped_group_preferences.append([option_instructions[min_ds[0]][1]])
                        mapped_group_preferences.sort(key=len)
                        del all_mapped_preferences[option_instructions[min_ds[0]][2]]
                        for elem in mapped_group_preferences:
                            all_mapped_preferences.append(elem)
                        tmpVar6 = list(all_mapped_preferences[option_instructions[min_ds[0]][3]])
                        tmpVar6.append(option_instructions[min_ds[0]][1])
                        all_mapped_preferences[option_instructions[min_ds[0]][3]] = tmpVar6
                else:
                    print("Unable to slice group " + str(group) + ", manually slice group " + str(group) + " and rerun this option.")
                    could_slice = False
                    print("Press any key to continue: ", end='', flush=True)
                    msvcrt.getch()
                    print()
            tmpVar8 += 1
        if could_slice:
            all_small_groups_combined = False
            cannot_combine = []
            while not all_small_groups_combined:
                group_index = 0
                for group in all_mapped_preferences:
                    if len(group) < min_in_group:
                        as_and_option_instructions = []
                        group2_index = 0
                        for group2 in all_mapped_preferences:
                            if len(group) + len(group2) <= max_in_group and group != group2:
                                approval_score=0
                                for student in group:
                                    # Subtract how many people in the new group want to sit with the student
                                    for j in group2:
                                        for elem in student_preferences.get(j):
                                            if elem == student:
                                                approval_score += 1
                                    # Subtract how many people the student wants to sit with in the new group
                                    for j in group2:
                                        for elem in student_preferences.get(student):
                                            if elem == j:
                                                approval_score += 1
                                    as_and_option_instructions.append([approval_score, group2_index])
                            group2_index += 1
                        #print(as_and_option_instructions)
                        if as_and_option_instructions != []:
                            #merge group and best group2
                            max_as = [0, as_and_option_instructions[0][1]]
                            for elem in as_and_option_instructions:
                                if elem[0] > max_as[0]:
                                    max_as = [elem[0], elem[1]]
                            #print(all_mapped_preferences)
                            #print(group_index)
                            #print(max_as[1])
                            all_mapped_preferences[group_index] = all_mapped_preferences[group_index][:] + all_mapped_preferences[max_as[1]][:]
                            del all_mapped_preferences[max_as[1]]
                            #print(all_mapped_preferences)
                        else:
                            cannot_combine.append(group)
                    group_index += 1
                all_small_groups_combined = True
                for group in all_mapped_preferences:
                    if len(group) < min_in_group and group not in cannot_combine:
                        all_small_groups_combined = False
            print("Done!")
            #print("Press any key to continue: ", end='', flush=True)
            #msvcrt.getch()
            #print()
    elif choice == "s":
        all_mapped_preferences.sort(len)
    elif choice == "r":
        all_mapped_preferences = []
        for student in student_preferences:
            isStudentInList = False
            for sub_list in all_mapped_preferences:
                for element in sub_list:
                    if student in sub_list:
                        isStudentInList = True
            if not isStudentInList:
                mapped_preferences = []
                all_mapped_preferences.append(map_preferences(student_preferences, student))
        all_mapped_preferences.sort(key=len)
    else:
        print("That is not an option.")
        print("Press any key to continue: ", end='', flush=True)
        msvcrt.getch()
        print()
    print()
