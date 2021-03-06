form_filler V3.0
~*~*~*~*~*~*~*~*~*~*~*~*~
New this version:

This version brings with it the addition of saving and another refactoring of
the data structure to accomodate it. The change means that there is no longer
a limit placed on the amount of assignments the user can create and upon
creation of the assignments they are now saved for future use.

The editing feature has not yet been implemented to the program and attempts
to edit the data will only allow for a selection of an assignment and nothing
else.

Most error handling has now been taken care of and there should be no(to the
degree that is expectable) sudden crashes.

To Do:
- Add the view & edit feature
- Ability to cancel mistaken assignment creation once undertaken
~*~*~*~*~*~*~*~*~*~*~*~*~
form_filler V2.0
~*~*~*~*~*~*~*~*~*~*~*~*~
New this version:

The entire data structure has been refactored to use dictionaries.
This change will allow for greater flexibility in data handling.
The door is now open to further feature implementations such as saving the
data for subsequent sessions and in-app data management.

i.e. viewing, deleting, editing data objects
~*~*~*~*~*~*~*~*~*~*~*~*~

A simple form filler program written in python3.

Dependencies:
openpyxl

Compile the source on the commandline and follow the prompts to fill out a
template form (an example of the handled template is included in the
repository).

The code is created to handle the included template and other excel
spreadsheet forms were not considered during the creation of this code.

This code is provided under no warranty, explicit or implied, free of charge
and freely available to use, change, or whatever your heart desires.


How to use:

Make sure you have python3 installed with the template excel file you will use
in the same directory (folder). Then, on a commandline, simply:

python form.py

and follow the on screen instructions.

Future Version Roadmap:
-add a save feature
-add data manipulation (editing, viewing, etc.)
