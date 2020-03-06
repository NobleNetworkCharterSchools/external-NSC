#!python3
'''
Functions for matching names with rosters
'''
import sys
import xlsxwriter
from botutils.tabletools import tableclass as tc
from botutils.tabletools import tabletools as tt
from botutils.tkintertools import tktools
from botutils.ADB import AlumniDatabaseInterface as aDBi

def get_field_names(match_table):
    fields = [None]*4
    headers = match_table.get_header_dict()
    for x in ['First Name', 'First name', 'first name', 'FirstName',
              'First_Name', 'firstname']:
        if x in headers: fields[0] = x
    for x in ['Last Name', 'Last name', 'last name', 'LastName',
              'Last_Name', 'lastname']:
        if x in headers: fields[1] = x
    for x in ['High_School', 'High_School__c', 'High School', 'Campus',
              'HS', 'high school', 'HighSchool']:
        if x in headers: fields[2] = x
    for x in ['HS_Class', 'HS_Class__c', 'Class', 'Grad Year', 'grad_yr',
              'hs_class', 'hs class', 'class']:
        if x in headers: fields[3] = x
    if None in fields[:3]:
        print('Unable to find all matching fields:\n'+
              'First Name: %s\nLast Name: %s\n'+
              'High School: %s\nHS Class: %s\n'+
              'From headers: %s' % tuple(fields)+(str(headers),))
        sys.exit()
    else:
        return fields

def get_SF_contacts():
    '''Returns a list of list with a specific set of columns:
    Id, LastName, FirstName, HS_Class__c, High_School__c
    '''
    try:
        contact_table = aDBi.getContactFields_Roster()
    except:
        print('Error attempting to grab data from Salesforce:',
                sys.exc_info()[0])
        print('#### Detailed error message follows ####')
        raise
    return contact_table

def get_csv_table_contacts(temp_contact):
    '''Ignores extra info, but fails if it doesn't find the following columns:
    Id, LastName, FirstName, HS_Class__c, High_School__c
    '''
    try:
        short_table = temp_contact.new_subtable(['Id',
                                                 'LastName',
                                                 'FirstName',
                                                 'HS_Class__c',
                                                 'High_School__c'])
    except:
        print('Roster table must contain the following columns:')
        print('Id\nLastName\nFirstName\nHS_Class__c\nHigh_School__c')
        print('#### Detailed error message follows ####')
        raise
    return short_table.get_full_table()

def short_name(orig, cut=5):
    '''Returns shorter version of string after first checking bounds'''
    if len(orig) >= cut:
        return orig[:cut]
    else:
        return orig

def short_names(orig, compare, cut=5):
    '''Returns shorter versions of the strings after first checking bounds'''
    if len(orig) >= cut:
        new_orig = orig[:cut]
    else:
        new_orig = orig
    shorter_cut = len(new_orig)
    if len(compare) >= shorter_cut:
        new_compare = compare[:shorter_cut]
    else:
        new_compare = compare
    return (new_orig, new_compare)

def find_matches(match_table, roster_table, 
                 fn_label, ln_label, hs_label, class_label):
    '''Adds columns to the match_table based on automated matching
    labels refer to the column headers for the matched fields within
    the match_table (roster_table labels are standard); class_label
    may be None, in which case will attempt to match without classes'''
    # first add placeholder fields
    num_records = len(match_table)
    match_table.add_column('Id', ['NotFound']*num_records)
    match_table.add_column('MatchCode',['Not matched']*num_records)
    
    #match table then has columns referenced as follows
    id_col = match_table.c('Id')
    match_code_col = match_table.c('MatchCode')
    fn_col = match_table.c(fn_label)
    ln_col = match_table.c(ln_label)
    hs_col = match_table.c(hs_label)
    if class_label:
        class_col = match_table.c(class_label)
        match_table.apply_func(class_label, lambda x: int(x))
    else:
        class_col = None

    #roster_table has these columns:
    # Id, LastName, FirstName, HS_Class__c, High_School__c
    roster_table.apply_func('HS_Class__c', lambda x: int(x))

    for row in match_table:
        hs = row[hs_col]
        hs_class = row[class_col] if class_col else None
        ln = row[ln_col].lower().strip()
        fn = row[fn_col].lower().strip()

        # compare with students for whom we think the hs (and class) matches
        if hs_class:
            compare_set = [comp for comp in roster_table if
                           comp[3] == hs_class and comp[4] == hs]
        else:
            compare_set = [comp for comp in roster_table if comp[4] == hs]

        for comp in compare_set:
            # Perfect match
            if comp[2].lower() == fn and comp[1].lower() == ln:
                row[id_col] = comp[0]
                row[match_code_col] = 'Perfect match'
                break
        if row[match_code_col] == 'Not matched':
            # Short first name match
            for comp in compare_set:
                fn_orig, fn_comp = short_names(fn, comp[2].lower(), 5)
                if fn_comp == fn_orig and comp[1].lower() == ln:
                    row[id_col] = comp[0]
                    row[match_code_col] = 'Short first name match'
                    break
        if row[match_code_col] == 'Not matched':
            # Short first and last name
            for comp in compare_set:
                fn_orig, fn_comp = short_names(fn, comp[2].lower(), 5)
                ln_orig, ln_comp = short_names(ln, comp[1].lower(), 5)
                if fn_orig == fn_comp and ln_orig == ln_comp:
                    row[id_col] = comp[0]
                    row[match_code_col] = 'Short first and last name match'

        if row[match_code_col] == 'Not matched' and hs_class:
            compare_set = [comp for comp in roster_table if
                           comp[4] == hs and comp[1].lower() == ln]
            # Full first and last without year
            for comp in compare_set:
                if comp[2].lower() == fn:
                    row[id_col] = comp[0]
                    row[match_code_col] = 'Perfect match on name, but not year'
                break

        if row[match_code_col] == 'Not matched' and hs_class:
            # Full last and short first without year
            for comp in compare_set:
                fn_orig, fn_comp = short_names(fn, comp[2].lower(), 5)
                if fn_comp == fn_orig:
                    row[id_col] = comp[0]
                    row[match_code_col] = 'Short first name match, but no year'
                    break
        

def get_alumni_match(match_table, roster_source_table=None):
    '''Function takes the name of a match Table that has at least these columns:
    First_Name, Last_Name, High_School, (various spellings allowed) and
    optionally HS_Class (also various spellings)
    And trys to append ID matches to the match_table
    If roster_file is empty, it will attempt to pull from Salesforce'''

    # First check for the proper fields in the match_table (a Table)
    fn, ln, hs, hs_class = get_field_names(match_table)
    
    # Load the bulk roster from Salesforce or the roster Table
    if roster_source_table: # either way, raw_contacts is list of lists
        raw_contacts = get_csv_table_contacts(roster_source_table)
    else:
        raw_contacts = get_SF_contacts()

    # Now pick a single high school from the roster:
    hs_contacts = tktools.reduce_table_by_checkbox(
            raw_contacts,
            'High_School__c',
            'Pick which High School(s) to use in the roster search',
            default=False)

    if not hs_class: # the match table doesn't have a HS class field
        roster = tktools.reduce_table_by_checkbox(
                hs_contacts,
                'HS_Class__c',
                'Pick which Classes to include in the roster search',
                default=False)
        roster_table = tc.Table(roster)
    else:
        roster_table = tc.Table(hs_contacts)

    # Now call the matching function, which changes the Table in place
    find_matches(match_table, roster_table, fn, ln, hs, hs_class)
    return match_table
