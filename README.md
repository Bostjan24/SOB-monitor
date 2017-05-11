## SOB-monitor

SOB - State Of Being

This is a little program I've written for myself. Fell free to use it.

It collects levels of happiness, energy and focus on hourly level.

You can choose whether you would save data to a spreadsheet or to the database.

<b>Saving data to spreadsheet</b>

When you run the program for the first time, you need to run it with create spreadsheet option (see below), to create an empty spreadsheet, wich will be created in the same folder where program is located.

<b>Saving data to the database</b>

If you want to save the data to the database, you need to create database yourself and do some tweaks in main program. Script to create the tables is located in the script.sql document. <b>I run the database on MariaDB server. I haven't tried any other servers, so I suggest that you use MariaDB as well. If not, there is a chance that the program won't work.</b>

  `-a, --show-averages=> Shows averages for selected sheet.`<br>
  `-c, --create-spreadsheet => Creates spreadsheet named SOBm-data.xlsx in current folder`<br>
  `--sheet-name [sheet_name] => Name of the sheet in wich data will be saved. Default: 'data'`<br>
  `--database-mode => Use database (you need to create it) instead of spreadsheet.`
  
