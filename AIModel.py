import random
import pandas as pd

class Student:
    def __init__(self, student_id, name, *args):
        self.student_id = student_id
        self.name = name
        self.additional_info = args

    def __repr__(self):
        return f"Student({self.student_id}, {self.name})"

def load_students_from_excel(file_path):
    df = pd.read_excel(file_path)
    students = []
    
    for _, row in df.iterrows():
        student = Student(
            row['Student ID'],
            row['Name'],
            row.get('Additional Info', None)
        )
        students.append(student)
        
    return students

def group_students_into_teams(students, max_team_size=4):
    random.shuffle(students)
    
    total_students = len(students)
    num_teams = total_students // max_team_size
    remainder = total_students % max_team_size

    teams = []
    current_index = 0

    for _ in range(num_teams):
        team = students[current_index:current_index + max_team_size]
        teams.append(team)
        current_index += max_team_size

    for i in range(remainder):
        teams[i % len(teams)].append(students[current_index])
        current_index += 1

    return teams

def introduce_random_variation(teams, variation_factor=0.1):
    num_swaps = max(1, int(len(teams) * variation_factor))

    for _ in range(num_swaps):
        team1, team2 = random.sample(teams, 2)
        if team1 and team2:
            student1 = random.choice(team1)
            student2 = random.choice(team2)

            team1.remove(student1)
            team2.remove(student2)
            team1.append(student2)
            team2.append(student1)

    return teams

def save_teams_to_excel(teams, output_file='teams.xlsx'):
    team_data = []
    
    for team_number, team in enumerate(teams, start=1):
        for member in team:
            team_data.append({
                'Team Number': team_number,
                'Student ID': member.student_id,
                'Student Name': member.name
            })
    
    df_teams = pd.DataFrame(team_data)
    df_teams.to_excel(output_file, index=False)


students = load_students_from_excel('students.xlsx')
teams = group_students_into_teams(students)
teams = introduce_random_variation(teams, variation_factor=0.2)
save_teams_to_excel(teams)
