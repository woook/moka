""" 

This script imports patients from a txt file into Moka
It then creates an array test for them and feeds back information to the user in the txt file 

##### DEBUGGING #####

Use --debug to get more verbose error messages printed to the terminal 

All SQL's are commented with DEBUG if returns need to be checked 
If a file has failed at the Patients stage, check DEBUG PATIENT
If a file has failed at the Patients stage, check DEBUG TEST



"""

import os 
import pandas as pd
import datetime
from ConfigParser import ConfigParser 
import pyodbc 
import getpass # to get username 
import socket # to get computer name 
import ImportTxtToDBConfig as config 
import sys
import argparse 
import numpy as np

# Read config file(must be called config.ini and stored in the same directory as script)
config_parser = ConfigParser()
print_config = config_parser.read(os.path.join(os.path.dirname(os.path.realpath(__file__)), "config.ini"))

class MokaConnector(object):
    """
    pyodbc connection to Moka database for use by other functions
    
   
    """
    def __init__(self):
        self.cnxn = pyodbc.connect('DRIVER={{SQL Server}}; SERVER={server}; DATABASE={database};'.format(
            server=config_parser.get("MOKA", "SERVER"),
            database=config_parser.get("MOKA", "DATABASE")
            ),
            autocommit=True
        )
        self.cursor = self.cnxn.cursor()

    def __del__(self):
        """
        Close connection when object destroyed
        """
        self.cnxn.close()

    def execute(self, sql):
        """
        Execute SQL, without catching return values (INSERT, UPDATE etc.)
        """
        self.cursor.execute(sql)

    def fetchall(self, sql):
        """
        Execute SQL catching all return records (SELECT etc.)
        """
        return self.cursor.execute(sql).fetchall()

    def fetchone(self, sql):
        """
        Execute SQL catching one returned record (SELECT etc.)
        """
        return self.cursor.execute(sql).fetchone()




'''=================================================== SCRIPT FUNCTIONS=================================================== ''' 

# for debugging flag
def arg_parse():	
	   """
	   Parses arguments supplied by the command line.
	       :return: (Namespace object) parsed command line attributes
	
	   Creates argument parser, defines command line arguments, then parses supplied command line arguments using the
	   created argument parser.
	   """
	   parser = argparse.ArgumentParser()
	   parser.add_argument('-d', '--debug', action='store_true', help="Run this mode for increase erroring printing when running scriptrr")
	   return parser.parse_args()




# Patient booking function ==========================================
def import_check_patients_table_moka(patient_error, i):
    '''
    This function takes a row of a df as an input to check if the patient is in moka
    then inserts them into moka if required and adds to the patient log
    It returns the row, as well as an error handling message if any processing errors occured 

    '''
    try:
        sql_check_patient = ("SELECT [Patients].[InternalPatientID]" # Check if patient is in the patients table
                "FROM ([dbo].[gwv-patientlinked] INNER JOIN [Patients] ON"
                "[dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID] )"
                "INNER JOIN [dbo].[gwv-dnaspecimenlinked] ON ([dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID])"
                "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}'"
        ).format(
            SpecimenTrust_ID = df.loc[i, "SpecimenTrustID"] 
        )
        sql_check_patient_return = mc.fetchall(sql_check_patient) # Check if SELECT has returned any rows 
        # print(sql_check_patient_return) DEBUG PATIENT
        if len(sql_check_patient_return) > 1: # If this returns >1 there's two patients with this spec number, NOT GOOD           
            error = "ERROR: Patient details in Moka Patients table twice!" # For debugging
            error_list.append(error)
            df.loc[i,'Patient_Moka_status'] = 'Failed'
            patient_error = True # Flag  
        else: 
            if len(sql_check_patient_return) == 1: # Patient already in Patients table
                message = "Success"
                df.loc[i,'Patient_Moka_status'] = message # Add status to df for logging 
            else: 
                sql_get_patient_ID = ("SELECT [dbo].[gwv-patientlinked].[PatientTrustID]"
                    "FROM [dbo].[gwv-dnaspecimenlinked] INNER JOIN [dbo].[gwv-patientlinked] "
                    "ON [dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID]"
                    "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}' "
                ).format(
                    SpecimenTrust_ID = df.loc[i,"SpecimenTrustID"] 
                )
                #print(sql_get_patient_ID) DEBUG PATIENT
                sql_check_patient_ID_return = mc.fetchall(sql_get_patient_ID) # Run SQL, returns a list
                if len(sql_check_patient_ID_return) != 1: # If does not return 1, there's no patient in GW 
                    error = "ERROR: Can't find patient in GW" # For debugging
                    error_list.append(error)
                    df.loc[i,'Patient_Moka_status'] = 'Failed' # Add status to df for logging 
                    patient_error = True # Flag                    
                else: 
                    patient_error = False
                    PatientTrustID = sql_check_patient_ID_return[0][0] # Assign the first part of the  returned tuple to PatientTrustID to a variable 
                    sql_insert_patient = ("INSERT INTO [Patients] ([PatientID], [s_StatusOverall], [BookinLastName],"
                                        "[BookinFirstName], [BookinSex], [MokaCreated], [MokaCreatedBy], [MokaCreatedPC])" 
                                        "VALUES ('{Patient_ID}', '{Patient_Status}', '{Last_Name}', '{First_Name}',"
                                        "'{Sex}', '{Created_date}' , '{Staff_username}', '{Staff_PC}')" 
                    ).format(
                        Patient_ID = PatientTrustID, # Use the tuple return to fill the insert query 
                        Patient_Status = config.status_inprogress,
                        Last_Name = df.loc[i,"LastName"],
                        First_Name = df.loc[i,"FirstName"],
                        Sex = df.loc[i,"Gender"],
                        Created_date= date_time,
                        Staff_username = username,
                        Staff_PC = computer_name
                    )
                    #print(sql_insert_patient) DEBUG PATIENT
                    mc.execute(sql_insert_patient) # Insert patient into patients table 
                    Internal_ID_return = mc.fetchone("SELECT @@IDENTITY")[0] # Return the primary key (InternalPatientID) for new insert
                    sql_insert_patient_log = ("INSERT INTO PatientLog([InternalPatientID], [LogEntry], [Date], [Login], [PCName]) "
                                "VALUES ('{Patient_ID}', 'New Patient added to Patients table {Patient_ID} "
                                "using the Automating booking in arrays script version {Script_version}',"
                                "'{Created_date}', '{Staff_username}' , '{Staff_PC}')" 
                    ).format(
                        Patient_ID = Internal_ID_return, # New InternalPatientID
                        Script_version = config.scriptversion, 
                        Created_date= date_time,
                        Staff_username = username,
                        Staff_PC = computer_name
                    )  
                    mc.execute(sql_insert_patient_log) # Insert into patients log
                    sql_check_patient = ("SELECT [Patients].[InternalPatientID] "
                                    "FROM [Patients]" 
                                    "WHERE [PatientID]='{Patient_ID}' AND [BookinLastName]= '{Last_Name}' "
                                    "AND [BookinFirstName]= '{First_Name}' AND [BookinSex]= '{Sex}' "
                                    "AND [MokaCreated]='{Created_date}'"
                ).format(
                    Patient_ID = PatientTrustID, # Use the tuple return to fill the insert query 
                    Last_Name = df.loc[i,"LastName"],
                    First_Name = df.loc[i,"FirstName"],
                    Sex = df.loc[i,"Gender"],
                    Created_date= date_time,
                )
                    #print(sql_check_patient) DEBUG PATIENT
                    sql_check_patient_return = mc.fetchall(sql_check_patient) 
                    if len(sql_check_patient_return) == 1: # If this returns 1, the patient has been added successfully 
                        message = "Success"
                        df.loc[i,'Patient_Moka_status'] = message # Add status to df
                    else: 
                        error = "ERROR: Inserting patient into Patient's table failed" # For debugging
                        error_list.append(error)
                        df.loc[i,'Patient_Moka_status'] = 'Failed'
                        patient_error = True
    except:
        patient_error = True # If the function fails, there was an error 
        error = "ERROR: Patient booking made an unexpected erro" # For debugging
        error_list.append(error)
    return(patient_error, df) # return error flag status and df

# Test booking function ========================================================================

def import_check_array_table_moka(test_error, i):
    '''
    This function checks if there is a test in array test for the Patient 
    If there is and its status 2,4 or 5 another is added or if the test is ongoing, nothing is added 
    '''
    try:
        sql_check_ArrayTest = ("SELECT [ArrayTest].[ArrayTestID]" # Check a test for this SpecimenTrustID is already in the ArrayTest table 
                                "FROM ( [Patients] INNER JOIN [ArrayTest] ON [Patients].[InternalPatientID] = [ArrayTest].[InternalPatientID])"
                                "INNER JOIN ([dbo].[gwv-dnaspecimenlinked] INNER JOIN [dbo].[gwv-patientlinked] ON"
                                "[dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID])"
                                "ON  [dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID]"
                                "WHERE ([dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}'"
                                "AND [ArrayTest].[StatusID] NOT IN (2,4,5))"  # Incase a patient with the same specimennumber has another test requested
                        ).format(
                            SpecimenTrust_ID = df.loc[i,"SpecimenTrustID"] 
                        )
        #print(sql_check_ArrayTest) DEBUG TEST
        sql_check_ArrayTest_return = mc.fetchall(sql_check_ArrayTest) 
        if len(sql_check_ArrayTest_return) == 1: # If this returns 1, there is an ongoing test already 
                message = "Test already booked in & status is not completed/pending/not possible"
                df.loc[i,'Booking_in_sample_status'] = message # add status to df
        else: # Patient is either in the DNA table with a completed/pending/ not possible status or not in there at all, to be inserted!
            sql_get_Array_ID = ("SELECT [Patients].[InternalPatientID], [dbo].[gwv-dnaspecimenlinked].[SpecimenID],  "
                " [dbo].[gwv-dnaspecimenlinked].[CreatedDate]" # Get data to form insert statement below 
                "FROM (([dbo].[gwv-dnanumberlinked] INNER JOIN [dbo].[gwv-dnaspecimenlinked] "
                "ON [dbo].[gwv-dnaspecimenlinked].[SpecimenID] = [dbo].[gwv-dnanumberlinked].[SpecimenID])"
                "INNER JOIN [dbo].[gwv-patientlinked] ON [dbo].[gwv-dnanumberlinked].[PatientID]= [dbo].[gwv-patientlinked].[PatientID])"
                "INNER JOIN [Patients] ON [dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID]"
                "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}'"
        ).format(
            SpecimenTrust_ID = df.loc[i,"SpecimenTrustID"] 
        )
            #print(sql_get_Array_ID) DEBUG TEST
            sql_ArrayTest_insert_return = mc.fetchall(sql_get_Array_ID)
            sql_get_check1ID = ("SELECT [Checker].[Check1ID]" # Get Check1ID of person running script, this is needed for ArrayTest table 
                            "FROM [Checker] "
                            "WHERE [Checker].[UserName]='{Staff_username}'"
            ).format(
                Staff_username = username
            )
            #print(sql_get_check1ID) DEBUG TEST
            sql_get_check1ID_return = mc.fetchall(sql_get_check1ID)
            sql_insert_ArrayTest = ("INSERT INTO [ArrayTest] ([InternalPatientID], [GWSpecID], [ReferralID],[StatusID], "
                                        "[RequestedDate], [BookedByID]) "
                                        " VALUES ('{Patient_ID}','{GW_Spec_no}', '{Patient_referral}','{Patient_Status}', "
                                        " '{Requested_date}',  '{Booked_in_by}')" 
            ).format(
                Patient_ID = sql_ArrayTest_insert_return[0][0], # The query returns a tuple, within a list
                GW_Spec_no = sql_ArrayTest_insert_return[0][1],
                Patient_Status = config.status_arraytobebookedin,
                Patient_referral = config.referral_array,
                Requested_date = sql_ArrayTest_insert_return[0][2],
                Staff_PC = computer_name,
                Booked_in_by = sql_get_check1ID_return[0][0]
            )
            mc.execute(sql_insert_ArrayTest)
            Array_test_ID_return = mc.fetchone("SELECT @@IDENTITY")[0] # return the newly generated ArrayTestID
            sql_insert_patient_log = ("INSERT INTO PatientLog([InternalPatientID], [LogEntry], [Date], [Login], [PCName])"
                                    "VALUES ('{Patient_ID}', 'New ArrayTest {Array_test_ID}"
                                    " using the Automating booking in arrays script version {Script_version}' ,"
                                    " '{Created_date}', '{Staff_username}' , '{Staff_PC}')" 
            ).format(
                Patient_ID = sql_ArrayTest_insert_return[0][0], 
                Array_test_ID = Array_test_ID_return,
                Script_version = config.scriptversion,
                Created_date = date_time,
                Staff_username = username,
                Staff_PC = computer_name
            )  
            mc.execute(sql_insert_patient_log) # insert into patients log
            sql_check_arraytest_after_insert = ("SELECT [ArrayTest].[ArrayTestID]"
                                        "FROM [ArrayTest]"
                                        "WHERE [ArrayTest].[ArrayTestID] = '{Array_test_ID}'" 
            ).format(
                Array_test_ID = Array_test_ID_return
            )
            #print(sql_check_arraytest_after_insert) DEBUG TEST 
            sql_check_arraytest_return_after_insert = mc.fetchall(sql_check_arraytest_after_insert) # run sql query
            if len(sql_check_arraytest_return_after_insert) != 1: # If this returns 1, the patient has successfully been booked in 
                error = "ERROR: Inserting Test into ArrayTest table failed" # For debugging
                error_list.append(error)
                df.loc[i,'Booking_in_sample_status'] = 'Failed'   
                test_error = True
            else:
                test_error = False
                df.loc[i,'Booking_in_sample_status'] = 'Success' # add status to df 
                sql_update_status = ("UPDATE [Patients] "   # Update Patients status in the Patients table to show new test. Check at UAT
                                    " SET [s_StatusOverall] = '{Patient_Status}'"
                                    " WHERE [PatientID] = '{Patient_ID}' "    
                ).format(       
                    Patient_Status = config.status_array,
                    Patient_ID = sql_ArrayTest_insert_return[0][0]
                )
                #print(sql_update_status) DEBUG TEST 
                mc.execute(sql_update_status)
                sql_insert_patient_log = ("INSERT INTO PatientLog([InternalPatientID], [LogEntry], [Date], [Login], [PCName]) "
                                        "VALUES ('{Patient_ID}', 'Patient status updated to Array " # Check at UAT
                                        " using the Automating booking in arrays script version {Script_version}',"
                                        " '{Created_date}', '{Staff_username}' , '{Staff_PC}')" 
                ).format(
                    Patient_ID = sql_ArrayTest_insert_return[0][0], 
                    Script_version = config.scriptversion,
                    Created_date = date_time,
                    Staff_username = username,
                    Staff_PC = computer_name
                )  
                mc.execute(sql_insert_patient_log) # insert into patients log
    except: 
        test_error = True # If the function fails, there was an error 
        error = "ERROR: Test booking made an unexpected error " # For debugging
        error_list.append(error)
    return(test_error, df)


# Save & move txt file ========================================================================

def save_move_txt(df):

    df.to_csv(to_process_path, sep ='\t', index = False)   # save the df to the txt file            
    if error_occured != True: # If no errors, move file to /Booked directory
        processed_path = os.path.join(config.processed_path+"/"+txt_file)
        try: # Try to move the completed txt file 
            os.rename((to_process_path), (os.path.join(config.processed_path+"/"+txt_file)))
        except:
            message ='Could not move file to /Booked folder. Is there a file with the same name in that folder?'
            df['Move'] = message 
            df.to_csv(to_process_path, sep ='\t', index = False) 
            print('All test succesfully completed but theres another file already in the /Booked folder with this file name. Please move mannually') # to print to Moka                    
    return(df)

# Error handling ========================================================================

def error_handling(df):
    count = 0
    if args.debug == True: # Debug mode activated
        print(error_list) # print error list 
    for i in range(len(df)): # create a loop to go through the df
        if df.loc[i,'Processed_status'] == 'FAILED':
            count = count + 1 # Count the number of failed rows 
        else: 
            count = count + 0                    
        if i == len(df) - 1: # if this is the last row of the df, get the count of failed rows
            if count >= 1:
                print('An error occured in ' +str(count)+' sample/s. Please see the txt file for error messages') # Print this to Moka for user 
            else:
                print('All sample/s imported into Moka with no errors' ) # Print this to Moka for user 
    return(df)



'''================== Define variables =========================== '''

mc = MokaConnector() # instantiate moka connector 

username = getpass.getuser() # get username 
computer_name = socket.gethostname() # get computer name 

t = datetime.datetime.now() # datetime.datetime.now returns six sig figs and Moka tables needs three sig figs
if t.microsecond % 1000 >= 500:  # check if there will be rounding up
    t = t + datetime.timedelta(milliseconds=1)  # manually round up
date_time= t.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]


error_list = [] # list to collect errors for debugging
patient_error_status = False # Patient error flag
test_error_status = False # Test error flag
error_occured = False # Error occured in script flag
args = arg_parse()



'''================== Import information from txt file =========================== ''' 
#try:
for txt_file in os.listdir(config.path): 
    if txt_file.endswith(".txt"): # Look for all .txt files in the folder 
        to_process_path = os.path.join(config.path+"/"+txt_file)
        df = pd.read_csv(to_process_path, delimiter = "\t")
        print(df)
        if 'Processed_status' not in df.columns:
            df['Processed_status'] = 'to process' # Add column to df
        for i in range(len(df)): # create a loop to go through the df
            if (df.loc[i,'Processed_status'] != 'Completed') or (df.loc[i,'Processed_status'] != 'FAILED') : # Don't want to re run a row that's already been done        
                #print('completed or failed')
            #else:
                #print('to process')   
                df['Gender'] = df['Gender'].replace(['Female','Male', np.nan ],['F','M', 'unknown']) # Change to fit in with Patients table column requirements 
                patient_error_status, df = import_check_patients_table_moka(False, i) # Attempt to book Patient into Moka
                if patient_error_status == True: # There was an error during processing patient
                    df.loc[i,'Processed_status'] = 'FAILED'
                    error_occured = True
                else: 
                    test_error_status, df  = import_check_array_table_moka(False, i) # Run ArrayTable function for those rows which didn't error
                    if test_error_status == True: # There was an error during processing test
                        df.loc[i,'Processed_status'] = 'FAILED'
                        error_occured = True
                    else:
                        df.loc[i,'Processed_status'] = 'Completed'
            if i == len(df) - 1: # check if this is the last row of the data frame
                save_move_txt(df) # run save and move
                error_handling(df) # Run error handling               
'''
except: 
    print('ERROR: Script not run')
    if args.debug == True:
        print('Script unexpectedly broken', sys.exc_info()[0]) # print the error to the terminals 
    df['Processed_status'] = 'Error'
    df.to_csv(to_process_path, sep ='\t', index = False) 
''' 

