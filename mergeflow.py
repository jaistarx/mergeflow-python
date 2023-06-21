import openpyxl
import subprocess
from datetime import datetime
import git

gitRepo = git.Repo('./')
gitRemotes = [remote.name for remote in gitRepo.remotes]


def validate_date(date_str):
    try:
        datetime.strptime(date_str, '%Y-%m-%d')
        return True
    except ValueError:
        return False

while True:
    remoteName = input('Enter remote name : ').lower()
    if remoteName in gitRemotes:
        break
    else:
        print(f"{remoteName} does not exist. Please try again.")
while True:
    branchName = input('Enter branch name : ').lower()
    if branchName in gitRepo.heads:
        break
    else:
        print(f"{branchName} does not exist in the remote {remoteName}. Please try again.")
while True:
    fromDate = input('Enter from date(YYYY-MM-DD) : ')
    if validate_date(fromDate):
        break
    else:
        print("Invalid date entered. Please try again.")
while True:
    toDate = input('Enter to date(YYYY-MM-DD)(if not entered, will take the current date) : ')
    if toDate == '' or validate_date(toDate):
        break
    else:
        print("Invalid date entered. Please try again.")

print('\n')
print('remote/branch : ' + remoteName + '/' + branchName)
print('From date : ' + fromDate)
if(toDate):
    print('To date : ' + toDate)
print('\n')

print('Process Started...')
print('\n')

print('Fetching Latest...')
print('\n')
subprocess.run(["git", "fetch", remoteName, branchName])
print('\n')
print('Fetch Complete :)')
print('\n')

workBookName = ['MRs']
commandBuilder = ['git', 'log']
if(remoteName and branchName):
    commandBuilder.append(remoteName + '/' + branchName)
    workBookName.append(remoteName + '_' + branchName)
if(fromDate):
    commandBuilder.append('--since=' + fromDate)
    workBookName.append('from_' + fromDate)
if(toDate):
    commandBuilder.append('--until=' + toDate)
    workBookName.append('to_' + toDate)

commandBuilder.append('--reverse')
gitProcess = subprocess.run(commandBuilder, stdout=subprocess.PIPE)
gitProcessResult = gitProcess.stdout.decode('utf-8')
gitData = gitProcessResult.split('\n')

consecutive_lines = []
excel_rows = []
commitMessage = ''
repo = ''
MrInfo = []
for line in gitData:
    if(line == ''):
        consecutive_lines = []
    elif(line.startswith('commit') or line.startswith('Merge:') or line.startswith('Author:') or line.startswith('Date:')):
        consecutive_lines.append(line)
    elif('See merge request' in line):
        repo = line.lstrip()
    else:
        lStrippedLine = line.lstrip()
        if(lStrippedLine and not (lStrippedLine.startswith('Merge branch') or lStrippedLine.startswith('Closes'))):
            commitMessage = lStrippedLine
    if(len(consecutive_lines) == 4):
        MrInfo = consecutive_lines
        consecutive_lines = []
    if(len(MrInfo) == 4 and repo != ''):
        MrInfo.append(commitMessage)
        MrInfo.append(repo)
        excel_rows.append(MrInfo)
        repo = ''
        commitMessage = ''
        MrInfo = []

workbook = openpyxl.Workbook()
worksheet = workbook.active
headings = ['Commit Id', 'Merge', 'Author', 'Date', 'Commit Message','Repo', 'Link']
worksheet.append(headings)
for row in excel_rows:
    cells = []
    cells.append(row[0][7:])
    cells.append(row[1][7:])
    cells.append(row[2][8:])
    cells.append(row[3][8:])
    cells.append(row[4])
    cells.append(row[5][18:])
    cells.append('<write domain>' + row[5][18:].replace('!', '/-/merge_requests/', 1))
    worksheet.append(cells)
workBookName = '_'.join(workBookName)
workBookName = workBookName.replace('/','_')
workBookName = workBookName + '.xlsx'
workbook.save(workBookName)
print('File name : ' + workBookName)
print('\n')
print('Process Complete :)')
print('\n\n')

input("Press Enter to exit...")
