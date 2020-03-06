#!python3
"""
Module for all interaction with Salesforce API
"""
import sys
import importlib
import pandas as pd
from simple_salesforce import Salesforce

def get_sf(cred_module):
    """
    Returns a Salesforce Object after being passed a secrets module
    """
    creds = importlib.import_module(cred_module)
    return Salesforce(username=creds.SF_LIVE_USERNAME,
                      password=creds.SF_LIVE_PASSWORD,
                      security_token=creds.SF_LIVE_TOKEN)

def get_query(sf, query_text, verbose=True):
    """
    Returns a list of lists based on a SOQL query with the fields as
    the header column in the first list/row
    """
    # execute query for up to 2,000 records
    gc = sf.query(query_text)

    records = gc['records']

    if verbose:
        print('Reading from %s object' % records[0]['attributes']['type'],
                flush=True)
    headers = list(records[0].keys())[1:] # get the headers
    return_table = [ [record[heading] for heading in headers]
                        for record in records]
    return_table.insert(0, headers)
    # the above is complete unless there are >2,000 records
    total_read_so_far = len(records)
    while not gc['done']:
        if verbose:
            print('Progress: {} records out of {}'.format(
                total_read_so_far, gc['totalSize']), flush=True)
        gc = sf.query_more(gc['nextRecordsUrl'], True)
        records = gc['records']
        total_read_so_far += len(records)
        next_table = [ [record[heading] for heading in headers]
                        for record in records]
        return_table.extend(next_table)
    return return_table

def get_fields(obj):
    """
    Returns a tuple:
    0: list of the fields in a given Salesforce object
    1: the 'soapType' for each of the passed fields
    """
    fields = obj.describe()['fields']
    return ([field['name'] for field in fields],
            [field['soapType'] for field in fields])

def get_table(sf, obj):
    """
    Requests an entire table for Salesforce API and returns as a Pandas
    DataFrame with Id as the index
    """
    # Get the field info and run the query
    field_names, field_types = get_fields(obj)
    query_text = 'SELECT ' + ', '.join(field_names) + ' FROM '+obj.name
    l_o_l = get_query(sf, query_text)

    # Convert the list of lists to DataFrame, converting where necessary
    df = pd.DataFrame(l_o_l[1:], columns=l_o_l[0]).set_index('Id')
    for i, var_type in enumerate(field_types):
        if var_type == 'xsd:dateTime':
            df[field_names[i]] = pd.to_datetime(df[field_names[i]])
        elif var_type == 'xsd:double':
            df[field_names[i]] = df[field_names[i]].astype('float64')
        elif var_type == 'xsd:boolean':
            df[field_names[i]] = df[field_names[i]].astype('bool')
        ''' # These have been commented out to avoid 'None' entries
        elif var_type == 'xsd:date':
            df[field_names[i]] = pd.to_datetime(
                                df[field_names[i]]).dt.date
        elif var_type == 'xsd:int':
            pass # this conversion can't handle NaNs
            #df[field_names[i]] = df[field_names[i]].astype('int64')
        elif var_type == 'xsd:string':
            pass # turns NaN to "None"
            #df[field_names[i]] = df[field_names[i]].astype('str')
        '''
    return df
