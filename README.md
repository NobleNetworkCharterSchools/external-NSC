# extern-NSC
Tools to process NSC reports

## Processing NSC File
Note, to interact with Salesforce, you'll need to create a file in the
main directory named 'salesforce_secrets.py' that has the following
three lines defining string variables:

```python
SF_LIVE_USERNAME = 'your_username'
SF_LIVE_PASSWORD = 'your_password'
SF_LIVE_TOKEN = 'your_token'
```

## Steps for processing:

1. Stash the current files in the "stash_dir" directory by running "step_1_save_tables.py"
2. Run import_nsc.py on the new NSC file
3. If we're using a separate system for a portion of Comer students, remove them from from import_nsc_output.csv
4. Also, check for N/As in the Student__c and College__c columns in import_nsc_output.csv; you'll resolve either by
   deleting them or by adding records to Salesforce and replacing the N/A with that new ID
5. If you're running on Noble's main database or Comer, check to make sure the True/False flag in nsc_modules.enrollment_match
   is set to the right system
6. Run merge_nsc.py
