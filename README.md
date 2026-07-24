# gemnet-behavioral
Repository for data structure conversion and validation for GEMNET's Behavioral component
everything is formatted for mac

## Task 1
- qual2an converts the annoying qualtrics format to a much nicer format for usage with seaborn and R
- tbh I forgot what the other one does

## Task 2
- task-2-cleaning is a validation script that performs several important steps when deciding whether or not to pay mturkers

1. Did they complete the survey (give more than 10 answers)
2. Did they press random buttons to hit the attention checks (screens for impossible features for a given concept)
3. Did they pass the attention check in the demographics section of qualtrics

It also organizes the data for easy vetting of the rest of the actual subtypes and returns which impossible features they gave per concept.
