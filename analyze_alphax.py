import csv

with open('AlphaX/AlphaX_Project_Schedule.csv', 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    tasks = list(reader)

print('=== AlphaX SCHEDULE STRUCTURE (167 tasks) ===')
print()

# Group by outline level
outline_1 = [t for t in tasks if t['Outline Level'] == '1']
outline_2 = [t for t in tasks if t['Outline Level'] == '2']
outline_3 = [t for t in tasks if t['Outline Level'] == '3']

print(f'Phase summaries (outline 1): {len(outline_1)}')
print(f'Workstreams (outline 2): {len(outline_2)}')
print(f'Detailed tasks (outline 3): {len(outline_3)}')
print()

print('=== WORKSTREAMS IN ALPHAX (Outline Level 2) ===')
for task in outline_2:
    name = task['Name']
    print(f'  {name}')
