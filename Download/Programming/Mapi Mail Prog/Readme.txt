PSC Database
Author JA
Date: 31/10/2000
Verion 1.0.0
------------------------------------

***Introduction
PSC database is a complete application that helps to manage the mailing lists that Planet source code sends to every member every day.
The application looks in the 'INBOX' of your mail program and creates a database with all the data that the PSC mailing lists contain about the new code entries.
I use Outlook Express to manage my inbox, and I can verify that this program works good when you set the 'Default mail handler' to Microsoft Outlook Express. I don't know what happens when you use an other mail application.

***How does it work
When yo start the application you can see direcrtly the code entries that exist in the database
If you want to look in the Inbox for new mailing lists from PSC then you must click 'Add' and follow the wizard steps.
The wizard will look for every code entry that does not exist in the database.
It will pick any usefull data that is posible (Title, date submitted, listdate, Level, category,Compatibility, Times accessed, and the location in the PSC where you can find the actual code to download it)
You can add some more data like your personal comment/value, the local directory where you have stored the code and some other stuff.
The most usefull thing is that you can search this database using many types of criteria. The search results appear in a Ritch text box and you save these results , you can format the way they appear and you can open from there the folder that the code is stored, the URL or you cab see the data of the database.
You can create and use filters for the data that will be displayed, using the filters wizard, and you can print the search results.
The most dificult part was to make a  custom parser function to get the data from the mailing lists, and there may be some bugs in there.
For example if a description of a code entry uses a number bullet
e.g.
1)lkjlkjsdf
2)lkjlkdsfjg
3)kjflksjdf
The parser will be confused in this mailing list
Also if Ian Ipolito changes the formating of the Mailing list....
it will be a disaster
Please Ian don't do that
The commenting of the code is not too good (not to much time)
I believe that peaple who understand will understand

****Thanks
Thanks to 
T De Lange, e-mail: tomdl@attglobal.net
for his wonderfull code of Ariel Color Combo
I used it as it was

-----------------------------------
Version2 notes
Look in the Description.txt file to see the changes since Version1
