Visual Basic Exercises

Sentence Breaker - A quick and fun exercise using Visual Basic that identifies the Value of a Range in the 'User Sentence' and splits that value into an index. The user enters numeric values in the 'A' column. Once the button is pushed, the value in the 'B' column will populate with the corresponding the user's inputed values that match the index value of the User Sentence.

Choose your Story - The user enters a 'Choice' numeric value and based on the value of the choice, the Sub will run an If/Else Statement that pushes 1 of 5 distinct messages to the Message Box.

Hornets Nest - This exercise combines for looping with if statements and incrementation. The loop runs through each cell and identifies the cell's value. If the iterrated cell contains the value "Hornet",the variable HornetCount's value increases by 1. If the variable BugCount's value is greater than one, the value of that cell will thereby be assigned the value 'Bug' and BugCount's value will decrease by one. Once BugCount's value is 0, any cells within the loop that contain 'Hornet' will be replaced by the value, 'Bee', which in turn will decrease the variable Beecount's value by 1. The loop will continue until it completes or BeeCount is 0.

Stock Analysis - This a more involved combination of the prior 3 exercises. All code occurs within a For each loop that iterrates through each sheet of the workbook. As the loop iterates through the rows, it first looks to the ticker symbol. 

If the ticker symbol from the previous iterration matches the current iterration, the total volume aggregates. 

Once a new symbol is dectected, final total volume is aggregated. The change for each day (row) is calculated by subtracting the closing value from the opening and then converting that value into a percentage. Furthermore:

    The last ticker symbol value captured before the aformentioned change in pushed to the I Column
    The total change for all rows with the corresponding symbol is pushed to the J Column. 
        If the amount of the change is greater than zero, the cell is highlighted green. 
        If not, the cell is highlighted red.
    The calculated percent change is pushed to the K Column
    The total volume is pushed in the L Column

After these values are pushed, both the Change and Total values are reset and the loop begins again.
    