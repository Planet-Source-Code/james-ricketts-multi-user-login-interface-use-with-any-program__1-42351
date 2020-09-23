This is the Main Form of the Program, this is the form that you would replace with the one from your application, applying similar menus to it to access the userlist and to log off.
---------------------------------------------

This Program is made to be easily adapted into your own VB applications. It can handle an unlimited number of users, and along with usernames and passwords it can store information such as program settings specific to each individual user. 

All data including settings is Encrypted so that it can only be altered from inside the program,and passwords cannot be simply read from the data file.

To help you, I have put all the subroutines and functions inside the mod_login.bas module. From there by simply altering a couple of constants you can change the length of records contained within the data file to your own requirements.

The program also has an Authorisation Password (contained within a constant inside the module). When you first log into the program, and the data file is EMPTY then it will prompt you to type the authorisation password to add a new user. This is so that if someone was to simply delete the datafile containing all the user information inside it then they would not be able to set up a new account and enter the program unless they had the authorisation password. 

Also The Title screen (frm_passwordscreen) has labels describing the program, such as developer, version and title. These can be altered in the project properties of VB so that the title screen will display your own program information. 

I only ask that you attribute my work in an about dialog and in the code of the program if you are to use it and to modify it.

Thanks and I hope you like it
James Ricketts
---------------------------------------------
