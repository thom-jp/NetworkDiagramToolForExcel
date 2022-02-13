# NetworkDiagramToolForExcel
PERT like chart creation tool

## About Files
You only need **NDT.xlsm** at bin folder.
Any other files in a project are just used for macro development.

## Tutorial
### Cleanup
1. Open **NDT.xlsm** and move to **Draw** sheet.
2. Click **Remove All Shapes** button at **Remove** section of header menu.
3. Move to **Schedule** sheet.
4. Select data rows(start from row 4 to the end) and delete them.

### Register Tasks
1. In Schedule sheet first data row, please input **START** at Task Name column.
2. Add any other **Task Name** as you want. <br /> To understand how task connection works, I reccomend you to put 3 tasks at least.
3. Add taskname **END** at the last data row.

Only Task Name was filled at that moment. START and END is mandatory for proper macro operation.

### Draw Tasks
1. Move to Draw Sheet and click **Plot Tasks** button. <br /> Now you can se several tasks as ovals in expanded header.
2. Layout those tasks on body of the sheet (White Area) by drag & drop.

### Connect Tasks
You can see 4 buttons in **Connect** section of the header.
OK, let's start from straight connection.

1. Select START and 2 more tasks by intentional order by pressing SHIFT.<br /> Now 3 tasks including START is selected, isn't it?
2. Click **Connect Straight** button at header menu.<br /> Then, 3 tasks connected by selected order.

This is straight connection.
Is it bend? Oh sorry I used word "straight" as meaning of not merged nor splitted.

As next explanation readiness, let's once discconect them by clicking **Remove Connections** button at **Remove** category.

Let's move on to split feature.

3. Select START task and 2 more tasks by intentional order by pressing SHIFT.
4. Click **Connect Split** button at header menu.<br /> Then, first selected task START connect to other 2 tasks.
This is split connection.

OK, click **Remove Connections** button again.
Let's move on to marge feature.

5. Select 2 more tasks and END task by intentional order by pressing SHIFT.
6. Click **Connect Marge** button at header menu.<br /> Then, first selected 2 tasks connect to END tasks.
This is marge connection.

Now you know how to connect tasks.

7. Connect All tasks by arrows as you want.

In case you disconnect arrow by unintentional drag & drop, program can't process schedule calcuration.
**Find Disconnection** button is ready for such situation.
OK, let's test this feature.

8. Disconnect any arrow from Oval by dragging arrow body or head or tail.
9. Click **Find Disconnection** button. <br /> Now you can see the disconnected arrows are highlighted by red.
10. Mannually connect arrows to proper oval.
11. Click **Find Disconnection** button again. <br /> Now you can see the re-connected arrow color is turned to brack again.

### Assign Task Numbers
You can assign task numbers by **Draw** sheet or **Schedule** sheet.

In **Draw** sheet, you can select all shapes by intentional order and click **Number Nodes as Selection Order**. <br />Then you can see all tasks are numbered. Schedule Sheet **No.** will be updated simultaneously.

In **Schedule** sheet, you can manually fill **No.** from 0. <br />In this case, you need to click **Plot Tasks** in Draw sheet to apply those numbers on diagram.

By the way, **Unnumber All Nodes** button in Draw sheet works without shape seleciton.

### Plot Schedule
Before proceed this step, you need to connect all shapes properly on Draw sheet and also assign numbers for all tasks without duplication.

In **Schedule** sheet, click "Plot Schedule" button.
Then, **Dependency** and **Planned Start** and **Planned End** will be automatically filled.

But you will notice that **Planned Start** for all tasks are sat to "Today".

This is because **Duration** and **Start Offset** are blank(0).

You can manually put Durations as business days you need to proceed each tasks.<br />(0 means within the day. 1 means next day.)

And also Start Offsets as days you can start tasks from dependent task ends.

You can also fills Duration and Start Offset by default by clicking **Fill Default Duration and Start Offset** buttons.<br .> This fills 1 for all Durations except START and END. This fills 0 for Start Offsets to START, START dependent tasks and END. But 1 for other tasks.

In Planned End calucuration, duration will be calcurated as business days. It means Saturday and Sunday will not counted in duration. You can set additional holidays for **Holidays** sheet.

### Modification
You can modify **Task Names** on schedule sheet manually. To apply that modification to Draw sheet, You can click **Plot Tasks** button.
You can also modify task names on Draw sheet, too. To apply that modification to Schedule sheet, You can click **Plot Schedule** button.

When you change order of tasks in schedule sheet by using autofilter button, worksheet function once will be corraped.
Don't worry. You can click **Plot Schedule** to fix it.

By the way, **Assign To** is not used by macro currently. It's up to you to fill or not.

**Wait for day of the week** does mean wait task completion till next specified day of the week.
You can set from 1(Sunday) through 7(Saturday). In case you set this column, Duration will be simply ignored.

### Operation
When you start a project using this book, you need some protection to avoid unintentional button clicking.

In that case, move to **Config** sheet and change from FALSE to TRUE on question "Would you like to lock design macros?".

You still see buttons. But error will be shown when you click distructive macro buttons after configured as above.

You can set status icons for each tasks in Draw sheet even if design macro locked.

To set/remove icon, select ovals(Make sure not text modification mode) and clcik Completed or Cancelled or Set Progress or Clear Icon button.
This feature can be used for PERT base progress management.

### Other
Order Node Vertical is used to locate several tasks on vertical line.
This is intended for initial rough layout creation before node connection.
It works for selected ovals.

Swap Node Location is for 2 node location swap.
This is sometimes useful for tangled connection resolution.
It works for selected 2 ovals.

## About Worksheets
### Schedule Sheet
This sheet will be synced bi-directionaly with Draw Sheet Diagram by Macro.

#### What you can do here
- Register **Task Name**
- Plan **Duration** days and **Start Offset** days from previous task
- **Plot Schedule** by auto calculation

#### What you should NOT do here
- Do not modify **Shape Object Name** by manually as it's used to find Draw Sheet Oval object.
- Do not modify **Dependency** by manually as it's automatically filled by macro. Manual change won't be applied to worksheet function.

### Draw Sheet
This sheet will be synced bi-directionaly with Schedule Sheet Diagram by Macro.
I hope everything you can do was explained in tutorial.

### Icon Sheet
This is used by Status category macros in Draw sheet.
I don't recommend you touch without macro knowledge.

Technically those icons are shapes in chart object. On each time you run status macrom, icon will be saved as bitmap in temporary folder on your PC and read for oval background.

### Holidays Sheet
You can input your national or personal holidays.

### Config Sheet
Currently used for locking some macro.
