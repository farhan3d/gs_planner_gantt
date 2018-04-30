# gs_planner_gantt
Automated simple Gantt and resource allocation weightage timeline from a pre-defined roadmap on your Google Sheet.

This sheet can be useful for projects where resource allocation over a timeline needs to be assessed to check whether anyone on the project is over or under allocated.

The template Google Sheet is available below:

https://docs.google.com/spreadsheets/d/1WCKA-byAZGyNoxHjfLPMqtWbR0MNDPGHcpzv7tbkXgM/edit?usp=sharing

The 'Timeline' sheet is where the resource allocation weightage and project milestone timelines are plotted. The Timeline sheet is not to be tampered with.

Follow these steps to use the sheet to plan your project:

1) Make a copy of the above sheet and add to your Google Drive
2) Go to 'Tools > Script Editor' and create a new project
3) Paste the code inside the code.js file in this repository in the editor and save the script
4) Go to the 'Team Planning' sheet, and add your project milestones with their start and end dates
5) Make sure to add a project code starting from 1 so that the script can associate timelines with the milestones
6) The sheet has a predefined list of team members (21). You need to change them to your team member names on BOTH the sheets
7) Now each resource has to work on a particular milestone. Assign amount of involvement for a particular resouce for a milestone by entering values from 1 - 3. These will be automatically color coded. Generally, 1 means minimal involvement in the development of that paricular project item, 2 is moderate, and 3 is heavy involvement with zero possiblity for other work
8) Make sure you have added the start and end dates for each milestone, otherwise nothing will be plotted on the Timeline sheet
10) Now go to the 'settings' sheet and make sure that the cell notations correctly correspond to the the right cells on both sheets. Also check the numbers at the bottom of the list. Generally you would only want to edit the row number of the project codes (second last), and the number of team members (last) in the list
11) Make sure to update the cell B2 in the 'Team Planning' sheet with the right date you need to start your Timeline sheet to start from
12) Now go to your script, select 'main' from the dropdown, and press play
13) Your timeline should be magically populated if all information had been added correctly.

The weightage timeline is color coded from yellow to very dark red, the latter meaning very heavily overloaded. The script works by adding up the weightages for that particular resource in the timeframe and writes in each cell. The colors are then added by conditional formatting to give a quick visual feedback.

Let me know if this works for you and if you found it useful in any way. I would be happy to help with any issues, so please drop an email at : farhan3d [ @ ] gmail [ dot ] com

Happy planning!
