# VBA_password
 

program do hakowania has³a w VBA project

1. The VBE will call a system function to create the password dialog box.
2. If user enters the right password and click OK, this function returns 1. If user     enters the wrong password or click Cancel, this function returns 0.
3. After the dialog box is closed, the VBE checks the returned value of the system function
4. if this value is 1, the VBE will "think" that the password is right, hence the locked VBA project will be opened.
5. The code below swaps the memory of the original function used to display the password dialog with a user defined function that will always return 1 when being called.

USAGE
Open the file(s) that contain your locked VBA Projects

Create a new xlsm file and store this code in Module1

Paste this code (SUB unprotected) under the above code in Module1 and run it