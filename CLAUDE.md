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