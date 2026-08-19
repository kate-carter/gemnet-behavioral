# gemnet-behavioral
Repository for data structure conversion and validation for GEMNET's Behavioral component
everything is formatted for mac

## Task 1
- qual2an converts the annoying qualtrics format to a much nicer format for usage with seaborn and R
- tbh I forgot what the other one does

## Task 2
- task-2-imp.py is the **CURRENT** version of the validation script
- task-2-cleaning is the (old) validation script that performs several important steps when deciding whether or not to pay mturkers

1. Did they complete the survey (give more than 10 answers)
2. Did they press random buttons to hit the attention checks (screens for  more than 3 impossible features for a given participant) For example, if they put one impossible term for one concept and then four for another but none for the third that's >3, so all of their data gets thrown out as "incoherent".
3. Did they pass the attention check in the demographics section of qualtrics (can only be applied to data from after 8/3/2026) For other data I just comment this section out, but I'll probably add a flagging system at some point.

To run, import all the dependencies and then "python task-2-imp.py". The file paths for input and output are set in the file at line 10ish. search "filepath" if you can't find it.

It also organizes the data for easy vetting of the rest of the actual subtypes and returns which impossible features they gave per concept.
