""" 

This script imports patients from a txt file into Moka
It then creates an array test for them and feeds back information to the user in the txt file 

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



mc = MokaConnector() # instantiate moka connector 

username = getpass.getuser() # get username 
computer_name = socket.gethostname() # get computer name 

t = datetime.datetime.now() # datetime.datetime.now returns six sig figs and Moka tables needs three sig figs
if t.microsecond % 1000 >= 500:  # check if there will be rounding up
    t = t + datetime.timedelta(milliseconds=1)  # manually round up
date_time= t.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]

'''================== Import information from txt file =========================== ''' 
#try:
for txt_file in os.listdir(config.path): 
    if txt_file.endswith(".txt"): # Look for all .txt files in the folder 
        print(txt_file)
        df = pd.read_csv(txt_file, delimiter = "\t")
        sep = '.'
        stripped_txt = txt_file.split(sep,1) [0]
        print(stripped_txt) # take the .txt off the end of the filename 
        print(df)
        if 'Processed_status' in df.columns:
            print('Has processed status ')
        else:
            print('Adding processed status column')
            df['Processed_status'] = 'to process'
            print(df)    
        for i in range(len(df)): # create a loop to go through the df
            print(i)
            if df.loc[i,'Processed_status'] == 'Completed':              
                print('This row is completed')
            else:
                print('This row is to procoess')    
                df['Gender'] = df['Gender'].replace(['Female','Male'],['F','M']) # change to fit in with Patients table column requirements
            # ======================= Book patient into MOKA ======================================= ''' 
                print('processing patient row')
                sql_check_patient = ("SELECT [Patients].[InternalPatientID]" #Check if patient is in the patients table
                        "FROM ([dbo].[gwv-patientlinked] INNER JOIN [Patients] ON "
                        "[dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID] )"
                        "INNER JOIN [dbo].[gwv-dnaspecimenlinked] ON ([dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID])"
                        "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}' "
                ).format(
                    SpecimenTrust_ID = df.loc[i, "SpecimenTrustID"] 
                )
                mc.execute(sql_check_patient) # run sql query 
                sql_check_patient_return = mc.fetchall(sql_check_patient) # Check if SELECT has returned any rows 
                #print(sql_check_patient)
                if len(sql_check_patient_return) > 1: # If this returns >1 there's two patients with this spec number, NOT GOOD
                    print("Patient already in patients table twice")
                    message = "ERROR: Patient details already in moka Patients table twice"
                    df.loc[i,'Patient_Moka_status'] = message
                    patient_error = True # Flag to ensure things are passed later on 
                else:
                    patient_error = False 
                    if len(sql_check_patient_return) == 1: # Patient already in patients tabley
                        print("Patient already in patients table")
                        message = "Patient details already in moka"
                        df.loc[i,'Patient_Moka_status'] = message # add status to df for logging 
                    else: 
                        print("Adding patient to patients table")
                        # find the PatientTrustID for the patient from GW to use in the patients insert
                        sql_get_patient_ID = ("SELECT [dbo].[gwv-patientlinked].[PatientTrustID]"
                            "FROM [dbo].[gwv-dnaspecimenlinked] INNER JOIN [dbo].[gwv-patientlinked] "
                            "ON [dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID]"
                            "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}' "
                        ).format(
                            SpecimenTrust_ID = df.loc[i,"SpecimenTrustID"] 
                        )
                        #print('Adding pt'+sql_get_patient_ID)
                        mc.execute(sql_get_patient_ID)
                        sql_check_patient_ID_return = mc.fetchall(sql_get_patient_ID)
                        if len(sql_check_patient_ID_return) != 1: # If does not return 1, there's no patient in GW
                            print("Patient not in GW")
                            message = "ERROR: Can't find patient in GW"
                            df.loc[i,'Patient_Moka_status'] = message # add status to df for logging 
                            patient_error = True # Flag to ensure things are passed later on
                        else: 
                            print("Patient in GW: to add to Patients table")
                            for PatientTrustID_return in mc.fetchall(sql_get_patient_ID): # Patient is in GW, continue
                            #sql_check_patient_ID_return instead fo the mc.fetchall above 
                            # len == 2 what happens ?  
                                PatientTrustID = PatientTrustID_return[0] # assign the first part of the  returned tuple to PatientTrustID to a variable 
                                sql_insert_patient = ("INSERT INTO [Patients] ([PatientID], [s_StatusOverall], [BookinLastName], "
                                                        "[BookinFirstName], [BookinSex], [MokaCreated], [MokaCreatedBy], [MokaCreatedPC]) " 
                                                        "VALUES ('{Patient_ID}', '{Patient_Status}', '{Last_Name}', '{First_Name}', "
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
                                #print('Patinet table insert')
                                #print(sql_insert_patient)
                                mc.execute(sql_insert_patient) # insert patient into patients table 
                                Internal_ID_return = mc.fetchone("SELECT @@IDENTITY")[0]
                                print('Patient ID generated')
                                print(Internal_ID_return)
                                sql_insert_patient_log = ("INSERT INTO PatientLog([InternalPatientID], [LogEntry], [Date], [Login], [PCName]) "
                                            "VALUES ('{Patient_ID}', 'New Patient added to Patients table {Patient_ID} added using the Automating booknig in arrays script',"
                                            " '{Created_date}', '{Staff_username}' , '{Staff_PC}')" 
                                ).format(
                                    Patient_ID = Internal_ID_return, 
                                    Created_date= date_time,
                                    Staff_username = username,
                                    Staff_PC = computer_name
                                )  
                                mc.execute(sql_insert_patient_log) # insert into patients log
                                sql_check_patient = ("SELECT [Patients].[InternalPatientID] "
                                                    "FROM [Patients]" 
                                                    "WHERE [PatientID]='{Patient_ID}' AND [BookinLastName]= '{Last_Name}' "
                                                    " AND [BookinFirstName]= '{First_Name}' AND [BookinSex]= '{Sex}' "
                                                    " AND [MokaCreated]='{Created_date}'"
                            ).format(
                                Patient_ID = PatientTrustID, # Use the tuple return to fill the insert query 
                                Last_Name = df.loc[i,"LastName"],
                                First_Name = df.loc[i,"FirstName"],
                                Sex = df.loc[i,"Gender"],
                                Created_date= date_time,
                            )
                                mc.execute(sql_check_patient) # run sql query 
                                sql_check_patient_return = mc.fetchall(sql_check_patient) 
                                #print(sql_check_patient)
                                if len(sql_check_patient_return) == 1: # If this returns 1, the patient has been added successfully 
                                    print("Patient successfully added to Patients table")
                                    message = "Patient successfully added into moka"
                                    df.loc[i,'Patient_Moka_status'] = message # add status to df
                                else: 
                                    message = "ERROR: Patient not added into moka"
                                    df.loc[i,'Patient_Moka_status'] = message
            #======================= Book in sample ===============================================  
                    #for idx, row in df.iterrows(): # Check if DNA for this SpecimenTrustID is already in the DNA table 
                if patient_error == True:
                    print('Patient error. Skip')
                else:

                    sql_check_ArrayTest = ("SELECT [ArrayTest].[ArrayTestID]"
                                            "FROM ( [Patients] INNER JOIN [ArrayTest] ON [Patients].[InternalPatientID] = [ArrayTest].[InternalPatientID])"
                                            "INNER JOIN ([dbo].[gwv-dnaspecimenlinked] INNER JOIN [dbo].[gwv-patientlinked] ON "
                                            " [dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID]) "
                                            " ON  [dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID] "
                                            "WHERE ([dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}' "
                                            "AND [ArrayTest].[StatusID] NOT IN  (2,4,5))"  # Incase a patient with the same specimennumber has another test requested
                                    ).format(
                                        SpecimenTrust_ID = df.loc[i,"SpecimenTrustID"] 
                                    )
                    mc.execute(sql_check_ArrayTest) # run sql query 
                    sql_check_ArrayTest_return = mc.fetchall(sql_check_ArrayTest) # run sql query
                    print('ArrayTest Check')
                    #print(sql_check_ArrayTest)

                    if len(sql_check_ArrayTest_return) == 1: # If this returns 1, there is a patient in the table already
                        print("Patient already booked into ArrayTest table")
                        message = "Patient sample already booked in & status is not completed/pending/not possible"
                        df.loc[i,'Booking_in_sample_status'] = message # add status to df
                    else: # Patient is either in the DNA table with a completed/pending/ not possible status or not in there at all, to be inserted!
                        sql_get_DNA_ID = ("SELECT [Patients].[InternalPatientID], [dbo].[gwv-dnaspecimenlinked].[SpecimenID],  "
                            " [dbo].[gwv-dnaspecimenlinked].[CreatedDate]" # Get data to form insert statement below 
                            "FROM (([dbo].[gwv-dnanumberlinked] INNER JOIN [dbo].[gwv-dnaspecimenlinked] "
                            "ON [dbo].[gwv-dnaspecimenlinked].[SpecimenID] = [dbo].[gwv-dnanumberlinked].[SpecimenID])"
                            "INNER JOIN [dbo].[gwv-patientlinked] ON [dbo].[gwv-dnanumberlinked].[PatientID]= [dbo].[gwv-patientlinked].[PatientID])"
                            "INNER JOIN [Patients] ON [dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID]"
                            "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}'"
                    ).format(
                        SpecimenTrust_ID = df.loc[i,"SpecimenTrustID"] 
                    )
                        #print('THIS IS THE GET DNA QUERY')
                        #print(sql_get_DNA_ID)
                        mc.execute(sql_get_DNA_ID)
                        sql_ArrayTest_insert_return = mc.fetchall(sql_get_DNA_ID)
                        print(sql_ArrayTest_insert_return)
                        print(sql_ArrayTest_insert_return[0][0])
                        print(sql_ArrayTest_insert_return[0][1])
                        print(sql_ArrayTest_insert_return[0][2])
                        #print(sql_ArrayTest_insert_return.split[1])
                        sql_insert_ArrayTest = ("INSERT INTO [ArrayTest] ([InternalPatientID], [GWSpecID], [ReferralID],[StatusID], "
                                                    "[RequestedDate], [BookedByID]) "
                                                    " VALUES ('{Patient_ID}','{GW_Spec_no}', '{Patient_referral}','{Patient_Status}', "
                                                    " '{Requested_date}',  '{Booked_in_by}')" 
                        ).format(
                            Patient_ID = sql_ArrayTest_insert_return[0][0], # the query returns a tuple, within a list this access the first element
                            GW_Spec_no = sql_ArrayTest_insert_return[0][1],
                            Patient_Status = config.status_arraytobebookedin,
                            Patient_referral = config.referral_array,
                            Requested_date = sql_ArrayTest_insert_return[0][2],
                            Staff_PC = computer_name,
                            Booked_in_by = 1201865448 # Need to change this, perhaps a query to checker table?
                        )
                        #print(sql_insert_ArrayTest) 
                        # return the ArrayTestID for this insert
                        mc.execute(sql_insert_ArrayTest)
                        Array_test_ID_return = mc.fetchone("SELECT @@IDENTITY")[0] # return the newly generated ArrayTestID
                        print ('new Array test ID')
                        print(Array_test_ID_return)
                        sql_insert_patient_log = ("INSERT INTO PatientLog([InternalPatientID], [LogEntry], [Date], [Login], [PCName]) "
                                                "VALUES ('{Patient_ID}', 'New ArrayTest {Array_test_ID} added using the Automating booknig in arrays script',"
                                                " '{Created_date}', '{Staff_username}' , '{Staff_PC}')" 
                        ).format(
                            Patient_ID = sql_ArrayTest_insert_return[0][0], 
                            Array_test_ID = Array_test_ID_return,
                            Created_date = date_time,
                            Staff_username = username,
                            Staff_PC = computer_name
                        )  
                        #print(sql_insert_patient_log)
                        mc.execute(sql_insert_patient_log) # insert into patients log
                        print('added to log')
                        sql_check_arraytest_after_insert = ("SELECT [ArrayTest].[ArrayTestID]"
                                            "FROM [ArrayTest] "
                                            "WHERE [ArrayTest].[ArrayTestID] =  '{Array_test_ID}' " 
                        ).format(
                            Array_test_ID = Array_test_ID_return
                        )
                        mc.execute(sql_check_arraytest_after_insert) # run sql query 
                        sql_check_arraytest_return_after_insert = mc.fetchall(sql_check_arraytest_after_insert) # run sql query
                        #print(sql_check_patient)
                        if len(sql_check_arraytest_return_after_insert) == 1: # If this returns 1, the patient has successfully been booked in 
                            print("Patient ArrayTest successfully booked into moka")
                            message = "Patient ArrayTest test successfully booked into moka"
                            df.loc[i,'Booking_in_sample_status'] = message # add status to df
                        else: 
                            print("ERROR: Patient ArrayTest not booked into moka")
                            message = "ERROR: Patient ArrayTest test not booked into moka"
                            df.loc[i,'Booking_in_sample_status'] = message    
                        # add to txt file 
                        #Update Patients status in the PatientsTable 
                        sql_update_status = ("UPDATE [Patients] "
                                            " SET [s_StatusOverall] = '{Patient_ID}'"
                                            " WHERE [PatientID] = '{Patient_ID}' "    
                        ).format(       
                            Patient_Status = config.status_array,
                            Patient_ID = sql_ArrayTest_insert_return[0]
                        )
                        mc.execute(sql_update_status)
                        print('Updated patient status for new array test in Patients table')
                        sql_insert_patient_log = ("INSERT INTO PatientLog([InternalPatientID], [LogEntry], [Date], [Login], [PCName]) "
                                                "VALUES ('{Patient_ID}', 'Patient status updated to Array added using the Automating booknig in arrays script',"
                                                " '{Created_date}', '{Staff_username}' , '{Staff_PC}')" 
                        ).format(
                            Patient_ID = sql_ArrayTest_insert_return[0][0], 
                            Created_date = date_time,
                            Staff_username = username,
                            Staff_PC = computer_name
                        )  
                        print(sql_insert_patient_log)
                        mc.execute(sql_insert_patient_log) # insert into patients log
                        print('added to log')
                    # insert new bits into txt file for user feedback 
                    #======================= Update txt file with messages =============================================== 
            message ='Completed'
            df.loc[i,'Processed_status'] = message 
            df.to_csv(stripped_txt+'.txt', sep ='\t', index = False)  
            df_check = pd.read_csv(txt_file, delimiter = "\t")

        if patient_error == False: # don't allow the file to be move if an error has occured
            if i != len(df) - 1:  # Check if this is the end of the dataframe 
                print('!!This is not the end!!')
            else:
                print('!!At the last row of the df!!')       
                if 'Processed_status' in df_check.columns:
                    pass # everything's been save to the txt file don't need to anything else 
                else:
                    print('back up save')
                    message ='As the original file was still open and couldnt be written too, this additional file was made. If the original file has since been closed, the text should now appear '
                    df_check['Back_up'] = message 
                    df_check.to_csv(stripped_txt+'_log.txt', sep ='\t', index = False)  # print to a new file incase the original one is open
                    new_path_log = processed_folder+stripped_txt+'_log.txt' 
                    old_path_log = os.path.abspath(stripped_txt+'_log.txt')
                    try: 
                        os.rename((os.path.abspath(stripped_txt+'_log.txt')), (config.processed_path+stripped_txt+'_log.txt'))
                    except:
                        message ='Could not move file to /Booked folder. Is there a file with the same name in that folder?'
                        df_check['Move'] = message 
                        df_check.to_csv(stripped_txt+'_log.txt', sep ='\t', index = False) 
                try: # try to move file to the /Booked directory 
                    os.rename((os.path.abspath(txt_file)), (config.processed_path+txt_file))
                except: 
                    print('theres already a file with that name in that folder!')
                    message ='Could not move file to /Booked folder. Is there a file with the same name in that folder?'
                    df['Move'] = message 
                    df.to_csv(stripped_txt+'.txt', sep ='\t', index = False) # Save error message to txt file    
'''                          
except: 
    print('AN ERROR HAS OCCURRED with the sample')
    print('this is the error', sys.exc_info()[0]) # print the error to the terminals 
    message ='ERROR OCCURRED IN PROCESSING'
    df['Processed_statuss'] = message 
    df.to_csv(stripped_txt+'.txt', sep ='\t', index = False)   
    df.to_csv(stripped_txt+'_log.txt', sep ='\t', index = False)
'''
    
