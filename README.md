# gemnet-behavioral
Repository for data structure conversion and validation for GEMNET's Behavioral component
everything is formatted for mac

## Task 1
- qual2an converts the annoying qualtrics format to a much nicer format for usage with seaborn and R
- tbh I forgot what the other one does

## Task 2
- task-2-cleaning is a validation script that performs several important steps when deciding whether or not to pay mturkers

1. Did they complete the survey (give more than 10 answers)
2. Did they include the "critical" dummy check item (a feature all subconcepts should have) \
   **Critical MUST be manually configured for each subconcept** should be around line 20

It also organizes the data for easy vetting of the rest of the actual subtypes
