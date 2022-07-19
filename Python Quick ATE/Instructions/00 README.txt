4/1/2019
Devon Murdza

The Quick ATE.py file will open the Test.xlsx workbook and send cmds/qrys based on the columns
The Test.xlsx file has example cmds/qrys listed in the columns. These can be changed to build a script
The Test_RESULTS.xlsx files has example results in place showing the filled in PASS/FAIL column
*Make sure the Quick ATE.py file and the Test.xlsx file are in the same folder, this is also where the RESULTS file will save



Procedure:
Step 1: 
    Create script in Test.xlsx based on Command Type descriptions listed below and rename file to whatever you like
Step 2:
    Edit the Blank.py file User Setup:
        1) USER_IP: Enter IP Address default(USER_IP or you may enter it each time the script is run
	2) DFLT_DELAY: Enter default delay
	3) DFLT_PROTOCOL: Specify default protocol if you wish to change it (default:TCP)
	4) LONG_TERM_LOG: If using the Loop function, this will save a results file on each iteration (On="1", Off="0")
	5) LOOP_AMOUNT: If using the Loop function, enter how many times you would like the program to loop ("inf" loops infinitely)
        5) EXCEL_SOURCE: Change the name of the workbook to be opened 
        6) EXCEL_RESULTS: Change the name of the RESULTS file to be saved per protocol
        7) Save the py file matching the name of the workbook
Step 3:
    Run the py file and follow the prompts to run the script
    If the entire script runs all the way through, the saved RESULTS file will contain the responses/results
    *Make sure the Blank.py file and the Blank.xlsx file are in the same folder, this is also where the RESULTS file will save



How to create the test:
Column descriptions Blank.xlsx:

Command Type:
    Command- Send command in the opened socket 
    Query- Send a query in the opened socket 
    CommandF- Send a command with 0 delay in the opened socket	
    Comment- Print message in the shell, may be used as a prompt for the Pause function 
    Pause- Pause the script until Enter is pressed     
    Wait- Causes a delay for specified time entered in the Command List column 
    Loop- Go back to the beginning of the list, may enter a delay in the Command List column (None/default=0)

Command List: 
    Used for any SCPI command or query to be sent to the power supply, 
    for entering delay length (Wait/Loop), or for entering a comment.

Response:
    In the RESULTS file the query response will be pasted here. 
    If using the Loop function, the total loop counter will print here.

Expected Response:
    If you would like to compare the query response with some expected response, enter it in this column.
    If you dont care about the response, leave the column blank.
    It may be necessary to put an aspostrophe in front of the expected response if it is a number in order to compare properly.

Pass/Fail:
    RED- Fail (Response does not match expected response/Time-out)
    GREEN- Pass (Response matches expected response)
    BLUE- Dont Care