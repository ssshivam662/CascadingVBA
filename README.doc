# CascadingVBA
DOCUMENTATION FOR ADVANCE VBA +DASHBOARD EXERCISE
1.	First thing to do is to create a unique list of all the nations, which will be used in the combo box for NATION. This will create Nation Combo box.







2.	Next two radio buttons are created for Geography and Employees, which will toggle between geographical, and employee based information.
3.	The cell link for those values named is Toggle.
1 for geography, 
2 for employee









4.	Next dynamic ranges are created for region and manager from VBA (using pivot tables) and they have been saved as region and manager.
5.	A new dynamic range named region_or_manager is created which will give region if toggle is 1 and manager if toggle is 2.
6.	Region_or_manager will be given to second combo box region/manager.


Toggle is 2
	
            

Toggle is 1
 

7.	Next VBA code for copy pasting ranges is done using if condition for toggle. If toggle is 1 copy paste all the regions based on nation to calculation tab otherwise copy paste manager information based on nation filter.
8.	In similar fashion territory and rep will interchange based on toggle 1 and 2.

Territory Filter based region and nations cascading


Rep Filter based manager and nations cascading


9.	To create cell link for combo boxes which will dynamically change this formula is used
=IF(toggle=1,INDEX(region,cell_value),All)
If toggle is 1 (geography) then based on cell value provided by user from combo box and   named range the value at that index in geography measure will be showed. If toggle is 2, it will simply go to All.
10.	After cascading part for nation, region/manager, territory/rep, one combo box is created for metrics (value/volume) based on which either value or volume metric will be shown. It is independent from cascading.
11.	The final subroutine total  is created which will copy paste all the values based on nation, region/manager, territory/rep, value/volume and toggle. It is copy pasting values from pivot table with similar filters.
12.	If toggle is 1, then user will see nation, region, territory otherwise they will see nation, manager, rep.
13.	NOTE- there are some cases where after applying all various selections from combo box no value/volume is obtained (nothing to copy paste). For those cases a message box will show which will ask to double click geography filter or start again by selecting nations as All.
 
14.	The table is further used to create chart (stacked area-line chart) to show products metrics per month. 
15.	If for some filters number of products in table decreases, the chart will distort/show empty values. So to avoid that we are dynamically hiding the rows which are empty 
16.	A simple function =if(cell=””,1,””) is created which will show 1 if cell is empty then in VBA Sheet4.Range("A33:A41").Special Cells(xlCellTypeFormulas, 1).EntireRow.Hidden = True
This formula is applied which is telling to hide rows which have value =1 in the cell otherwise do nothing. (Error can occur is there is nothing to show so use on error resume next)
17.	At the end navigations are created at each page which will help us toggle between all the pages.
