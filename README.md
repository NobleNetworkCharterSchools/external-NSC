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
## before processing, create a virtual environment and install libraries:

*(Note, all of this assumes that you've installed Python and Git and are*
*accessing these commands via a Git Bash shell)*

```python
python -m venv .env
source .env/Scripts/activate
pip install -r requirements.txt
```
## If you need to refresh from the repository, type:

```python
git pull https://github.com/NobleNetworkCharterSchools/external-NSC.git
```
## If you want to leave this environment:

```python
deactivate
```
*To reenter the environment, type the "source .env/Scripts/activate" line above*

## Steps for processing:

1. Stash the current files in the "stash_dir" directory by running "step_1_save_tables.py":
```python
python step_1_save_tables.py
```

2. Run import_nsc.py on the new NSC file (set the date for the effective date of the NSC file
```python
python import_nsc.py -date 11/19/2019
```
(find the NSC file by clicking in the box in the upper right then hit OK. Make sure the NSC file is a CSV)

3. When you run this process, it might warn about missing_degrees. Look at the instructions on the console
   message and add new rows to the inputs/degreelist.csv file per it's instructions.
   
4. Similarly, there might be a warning about missing colleges. Add those per instructions as well.

5. Once the above step is complete, you will have an output labelled "import_nsc_output.csv" Check that file
   for N/As in the Student__c and College__c columns; you'll resolve these either by
   deleting them or by adding records to Salesforce and replacing the N/A with that new ID (make sure
   the ID is 18 digits and not 15)

6. Run merge_nsc.py:
```python
python merge_nsc.py
```

7. Take the 3 outputs of this file and load them to Salesforce with dataloader.io:
   - new_enr_<date>.csv: create new enrollment records with this info
   - enr_update_<date>.csv: update enrollments with this info
   - con_update_<date>.csv: update contacts with this info
