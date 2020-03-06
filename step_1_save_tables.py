#!python3
"""
First step in the process for working with NSC files (run this first)

Saves the contacts, enrollments, and accounts tables locally to 
stash_dir
"""
import sys

from modules import sf_interface

def main(cred_file):
    sf = sf_interface.get_sf(cred_file[:-3])
    for fn, obj in (
                    ('contacts.csv', sf.Contact),
                    ('accounts.csv', sf.Account),
                    ('enrollments.csv', sf.Enrollment__c)):
        print('Downloading and saving '+fn, flush=True)
        df = sf_interface.get_table(sf, obj)
        df.to_csv('stash_dir/'+fn, index_label='Id')


if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1]) # pass the secrets file
    else:
        main('salesforce_secrets.py')
