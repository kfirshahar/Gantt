# Gantt Chart

## Goal
Build an excel template for project Gantt

## Inputs
The inputs should be inserted by the user en via excel tab and each one represents a constraint which should be taken into consideration when preparing the plan.
1. Tasks - each task has many columns but the relevant ones are category, priority, complexity, and number of sub-tasks.
2. Time - displayed in week numbers, and each week has a different number of available work days per assignee which derives from general holidays, personal vacations, and individual bandwidth. 
3. Assignees - the person who execute do the task, each with different level of proficiency in execution.
4. Equipment - the equipmemnt necessary execute the task, an assignee requires an available equipment to do the task, and there is a limited number of equipments in the shared 'pool' of equipments which varies over time, so in some weeks there are less equipments than assignees which affects the effective bandwidth. 

## Output
A gantt chart in separate excel tab which dynamically shows allocation of tasks, assignees and equipment allocation over a range of dates by week numbers.
The end results should be a simple to use excel file with few tabs, in which the user can add Tasks, Assignees, Equipment, and available work days (a combined constraint of assignee, equipment, and time)
There should be 2 levels of details for displaying the gantt chart:
1. High-level - show matrix of allocations by assignee, equipment, by week number and by the tasks without sub-items.
2. Deep-Dive - show matrix of allocations by assignee, equipment, by day and by sub-task.

# How to build it
It can be done either directly by generating an excel from claude-code skills or claude-cowork skills, or start by building a small python app (without a DB, backend, UI, etc.) which reprersents the model and the constraints, and the excel with all necessary tabs can be exported from it.
The important part is that the excel template should be the end result which can be exported, and it can be generic enough so it can be duplicated and customized for any type of project, tasks, equipment and assignees.
Brainstorm and ask me few guiding questions on how to build it, but the idea is to keep it simple, without over-complications of dependency between tasks like it is sometimes done in Gantt charts.

# Appendix A - Enhancements
1. A more realistic 'check' column in Tasks and Sub-Tasks tab which reflects a more accurate status of tasks and sub-tasks.
-  When a task is DONE, there is no option to update it's status and 'freeze' it so it won't be re-ranked or re-scheduled.
since the task's start week (and maybe even end week) is before the Config's start week - it appears as 'start outside horizon'
- When a task doesn't fit into the schedule it appears as either 'overruns horizon' or as 'not scheduled' and it is not clear what is the logic behind it. As a result it is not clear how to make changes in order to make the plan converge
- Sugggest a better way to handle such cases so it is possible to use the XLSX to dynamically manage the project after it started and adjust the plan.
2. Easy ingestion of real data to input tabs from JSON files and a matching skill for an agent to update tasks/sub-tasks status on a weekly/daily basis.
- Ingestion should be suported for 1st time ingestion, and also for updates in which new tasks/sub-tasks may appear
- The new/updated JSON data should include updated statuses of tasks/sub-tasks (see item #1 first bullet) with TODO/In-Progress/DONE values, but there should be still an option to edit statuses and everything else in the XLSX as before.
- There should also be an option to export the input-data from existing XLSX, i.e. the data from input tabs such as: Tasks, Sub-Tasks, Assignees, Capacity, Equipment, Holidays and Config.
- Backward compatibility of different versions of the XLSX when exporting/importing data should be as simple as possible, i.e. when there are new fields in new version (for example the status which will be added to tasks/sub-tasks) leave it blank or with a defined default-value,such as TODO, so it is easy to understand what happens.

# Appendix B - Machine and Agent Compatability
1. The python code including the pytest part is planned to be deployed and run on a windows machine which has python3.12 with openpyxl, and  Office 2019.
If I understand correcly only the LibreOffice should be adjusted to make tests pass, analyze what changes are required to be platform independent.
2. The skills and target agent will not necessarily be Claude, but rather Opencode on windows, which has skills support but the PC primitives such as proceeses handling, bash scripts, etc. will need to be replaced by windows PC shell/power-shell scripts. This work should by analyzed and planned by you for the target agent to read and implement using a dedicated skill-creator, but should not be executed by Claude on this PC.
3. Migration, porting and deployment from current MacOS + Claude -> Windows + Opencode should be done via pushing to github remote repo. Real live data exists as XLSX in v4 format in the target platform and should work with it smoothly.
