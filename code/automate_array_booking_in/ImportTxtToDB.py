
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

# Read config file(must be called config.ini and stored in the same directory as script)
config = ConfigParser()
print_config = config.read(os.path.join(os.path.dirname(os.path.realpath(__file__)), "config.ini"))

class MokaConnector(object):
    """
    pyodbc connection to Moka database for use by other functions
    
   
    """
    def __init__(self):
        self.cnxn = pyodbc.connect('DRIVER={{SQL Server}}; SERVER={server}; DATABASE={database};'.format(
            server=config.get("MOKA", "SERVER"),
            database=config.get("MOKA", "DATABASE")
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
date_time = datetime.datetime.now() # get current date time 

'''================== Import information from txt file =========================== ''' 
#try:
for txt_file in os.listdir("H:/04_automating_moka_project"): 
    if txt_file.endswith(".txt"): # Look for all .txt files in the folder 
        print(txt_file)
        df = pd.read_csv(txt_file, delimiter = "\t")
        sep = '.'
        stripped_txt = txt_file.split(sep,1) [0]
        print(stripped_txt) # take the .txt off the end of the filename 
        print(df)
        if 'Processed' in df.columns: # check if the txt file has been processed by the script already
            print('processed')
        else: 
            print('not processed')       
            #df['Gender'] = df['Gender'].replace(['Female','Male'],['F','M']) # change to fit in with Patients table column requirements
            '''======================= Book patient into MOKA ======================================= ''' 
            for idx, row in df.iterrows():  ## Check if the patient is in the patients table 
                sql_check_patient = ("SELECT [Patients].[PatientID]"
                        "FROM ([dbo].[gwv-patientlinked] INNER JOIN [Patients] ON "
                        "[dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID] )"
                        "INNER JOIN [dbo].[gwv-dnaspecimenlinked] ON ([dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID])"
                        "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}' "
                ).format(
                    SpecimenTrust_ID = row["SpecimenTrustID"] 
                )
                mc.execute(sql_check_patient) # run sql query 
                sql_check_patient_return = mc.fetchall(sql_check_patient) # Check if SELECT has returned any rows 
                #print(sql_check_patient)
                if len(sql_check_patient_return) == 1: # If this returns 1, there is a patient in the table already
                    print("Patient already in patients table")
                    message = "Patient details already in moka"
                    df['Patient_Moka_status'] = message # add status to df for logging 
                else: 
                    print("Adding patient to patients table")
                    # find the PatientTrustID for the patient from GW to use in the patients insert
                    sql_get_patient_ID = ("SELECT [dbo].[gwv-patientlinked].[PatientTrustID]"
                        "FROM [dbo].[gwv-dnaspecimenlinked] INNER JOIN [dbo].[gwv-patientlinked] "
                        "ON [dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID]"
                        "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}' "
                ).format(
                    SpecimenTrust_ID = row["SpecimenTrustID"] 
                )
                    mc.execute(sql_get_patient_ID)
                    sql_check_patient_ID_return = mc.fetchall(sql_get_patient_ID)
                    if len(sql_check_patient_ID_return) == 1: # If does not return 1, there's no patient in GW
                        print("Patient not in GW")
                        message = "ERROR: Can't find patient in GW"
                        df['Patient_Moka_status'] = message # add status to df for logging 
                    else: 
                        print("Patient in GW to add to Patients table")
                        for PatientTrustID_return in mc.fetchall(sql_get_patient_ID): # Patient is in GW, continue
                            PatientTrustID = PatientTrustID_return[0] # assign the first part of the  returned tuple to PatientTrustID to a variable 
                            sql_insert_patient = ("INSERT INTO [Patients] ([PatientID], [s_StatusOverall], [BookinLastName], "
                                                    "[BookinFirstName], [BookinSex], [MokaCreated], [MokaCreatedBy], [MokaCreatedPC]) " 
                                                    "VALUES ('{Patient_ID}', 3, '{Last_Name}', '{First_Name}', "
                                                    "'{Sex}', '{Created_date}', '{Staff_username}', '{Staff_PC}')" 
                            ).format(
                                Patient_ID = PatientTrustID, # Use the tuple return to fill the insert query 
                                Last_Name = row["LastName"],
                                First_Name = row["FirstName"],
                                Sex = row["Gender"],
                                Created_date= date_time,
                                Staff_username = username,
                                Staff_PC = computer_name
                            )
                            #print(sql_insert_patient)
                            mc.execute(sql_insert_patient) # insert patient into patients table 
                            for idx, row in df.iterrows(): # Check if patient has been inserted into patients table
                                sql_check__insert_patient = (("SELECT [Patients].[PatientID]"
                                        "FROM ([dbo].[gwv-patientlinked] INNER JOIN [Patients] ON "
                                        "[dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID] )"
                                        "INNER JOIN [dbo].[gwv-dnaspecimenlinked] ON ([dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID])"
                                        "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}' "
                                ).format(
                                    SpecimenTrust_ID = row["SpecimenTrustID"] 
                                        ) 
                                )
                                mc.execute(sql_check_patient) # run sql query 
                                sql_check__insert_patient_return = mc.fetchall(sql_check__insert_patient) 
                                #print(sql_check_patient)
                                if len(sql_check__insert_patient_return) == 1: # If this returns 1, the patient has been added successfully 
                                    print("Patient successfully added to Patients table")
                                    message = "Patient successfully added into moka"
                                    df['Patient_Moka_status'] = message # add status to df
                                else: 
                                    message = "ERROR: Patient not added into moka"
                                    df['Patient_Moka_status'] = message
            '''======================= Book in sample =============================================== '''         
            for idx, row in df.iterrows(): # Check if DNA for this SpecimenTrustID is already in the DNA table 
                sql_check_DNA = ("SELECT [DNA].[DNANumber]"
                                        "FROM ( [Patients] INNER JOIN [DNA] ON [Patients].[InternalPatientID] = [DNA].[InternalPatientID]) "
                                        "INNER JOIN ([dbo].[gwv-dnaspecimenlinked] INNER JOIN [dbo].[gwv-patientlinked] ON "
                                        " [dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID]) "
                                        " ON  [dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID] "
                                        "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}' "
                                        "AND [Patients].[s_StatusOverall] != '4' " # Incase a patient with the same specimennumber has another test requested
                                ).format(
                                    SpecimenTrust_ID = row["SpecimenTrustID"] 
                                )
                mc.execute(sql_check_DNA) # run sql query 
                sql_check_DNA_return = mc.fetchall(sql_check_DNA) # run sql query
                #print('DNA Check')
                #print(sql_check_DNA)
                if len(sql_check_DNA_return) == 1: # If this returns 1, there is a patient in the table already
                    print("Patient already booked into DNA table")
                    message = "Patient sample already booked in & status is not completed"
                    df['Booking_in_sample_status'] = message # add status to df
                else: # Patient is either in the DNA table with a completed table or not in there at all 
                    for idx, row in df.iterrows(): # check if this is second request with the same specimentrustID
                        sql_check_DNA = ("SELECT [DNA].[DNANumber]"
                                                "FROM ( [Patients] INNER JOIN [DNA] ON [Patients].[InternalPatientID] = [DNA].[InternalPatientID]) "
                                                "INNER JOIN ([dbo].[gwv-dnaspecimenlinked] INNER JOIN [dbo].[gwv-patientlinked] ON "
                                                " [dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID]) "
                                                " ON  [dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID] "
                                                "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}' "
                                                "AND [Patients].[s_StatusOverall] = '4' " # Incase a patient with the same specimennumber has another test requested
                                        ).format(
                                            SpecimenTrust_ID = row["SpecimenTrustID"] 
                                        )
                        mc.execute(sql_check_DNA) # run sql query 
                        sql_check_DNA_return = mc.fetchall(sql_check_DNA) # run sql query
                        #print('DNA Check')
                        #print(sql_check_DNA)
                        if len(sql_check_DNA_return) == 1: # If this returns 1, patient is having another test requested with the same specimenID
                            print("Patient already booked into DNA table & status is complete")
                            message = "Patient in DNA table, status completed. Adding new DNA test"
                            df['Booking_in_sample_status'] = message # add status to df
                        else: 
                            print('Patient not booked into to DNA table') # No DNA booking at all!
                    sql_get_DNA_ID = ("SELECT [Patients].[InternalPatientID], [dbo].[gwv-dnanumberlinked].[DNANo]," # Get data to form insert statement below 
                        "[dbo].[gwv-dnaspecimenlinked].[Concentration] "
                        "FROM (([dbo].[gwv-dnanumberlinked] INNER JOIN [dbo].[gwv-dnaspecimenlinked] "
                        "ON [dbo].[gwv-dnaspecimenlinked].[SpecimenID] = [dbo].[gwv-dnanumberlinked].[SpecimenID])"
                        "INNER JOIN [dbo].[gwv-patientlinked] ON [dbo].[gwv-dnanumberlinked].[PatientID]= [dbo].[gwv-patientlinked].[PatientID])"
                        "INNER JOIN [Patients] ON [dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID]"
                        "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}'"
                ).format(
                    SpecimenTrust_ID = row["SpecimenTrustID"] 
                )
                    #print('THIS IS THE GET DNA QUERY')
                    #print(sql_get_DNA_ID)
                    mc.execute(sql_get_DNA_ID)
                    for DNA_ID_reutrn in mc.fetchall(sql_get_DNA_ID): # return the PatientID 
                        if DNA_ID_reutrn[2] == None: # If the SQL return has no concentration for this DNA
                            print('no concentration found')
                            sql_insert_DNA_no_conc = ("INSERT INTO [DNA] ([InternalPatientID], [DNANumber], [DNAComment]) "
                                                " VALUES ('{Patient_ID}','{DNANo}', 'TEST')" 
                            ).format(
                                Patient_ID = DNA_ID_reutrn[0], 
                                DNANo = DNA_ID_reutrn[1]
                            )
                            print(sql_insert_DNA_no_conc) 
                            mc.execute(sql_insert_DNA_no_conc)
                        #print(sql_insert_DNA_no_conc) # run sql query 
                        else: # if there is a concentration 
                            print('Concentration found')
                            sql_insert_DNA_conc = ("INSERT INTO [DNA] ([InternalPatientID], [DNANumber],[Concentration],  [DNAComment]) "
                                                " VALUES ('{Patient_ID}','{DNANo}', '{Conc}', 'TEST' )" 
                            ).format(
                                Patient_ID = DNA_ID_reutrn[0], 
                                DNANo = DNA_ID_reutrn[1],
                                Conc = DNA_ID_reutrn[2]
                            )
                            print(sql_insert_DNA_conc) 
                        #try:
                            mc.execute(sql_insert_DNA_conc)
                            print('Sample added to DNA table') # run sql query
                        #except: 
                            #print('patient already in DNA table?')
                            #!! put in error handling for if patient and DNA sample no mixed up
                            # #print(sql_insert_DNA_no_conc)
                        # Insert in patient log 
                        sql_insert_patient_log = ("INSERT INTO PatientLog([InternalPatientID], [LogEntry], [Date], [Login], [PCName]) "
                                                "VALUES ('{Patient_ID}', 'New DNA Number {DNANo} added',"
                                                " '{Created_date}', '{Staff_username}' , '{Staff_PC}')" 
                        ).format(
                            Patient_ID = DNA_ID_reutrn[0], 
                            DNANo = DNA_ID_reutrn[1],
                            Created_date= date_time,
                            Staff_username = username,
                            Staff_PC = computer_name
                        )
                        #print(sql_insert_patient_log)
                        mc.execute(sql_insert_patient_log) # run sql query
                        for idx, row in df.iterrows(): 
                                    sql_check_DNA_after_insert = ("SELECT [DNA].[DNANumber]" # check if DNA havs been added to the DNA table 
                                            "FROM (([dbo].[gwv-dnaspecimenlinked] INNER JOIN [dbo].[gwv-patientlinked] "
                                            "ON [dbo].[gwv-patientlinked].[PatientID] = [dbo].[gwv-dnaspecimenlinked].[PatientID]) "
                                            " INNER JOIN [Patients] ON [dbo].[gwv-patientlinked].[PatientTrustID] = [Patients].[PatientID]) "
                                            " INNER JOIN DNA ON [Patients].[InternalPatientID] = [DNA].[InternalPatientID] "
                                            "WHERE [dbo].[gwv-dnaspecimenlinked].[SpecimenTrustID]='{SpecimenTrust_ID}' "
                                        # " AND [dbo].[gwv-dnaspecimenlinked].[Active] = '0' " # to add back in 
                                    ).format(
                                        SpecimenTrust_ID = row["SpecimenTrustID"] 
                                    )
                                    mc.execute(sql_check_DNA_after_insert) # run sql query 
                                    sql_check_DNA_return_after_insert = mc.fetchall(sql_check_DNA_after_insert) # run sql query
                                    #print(sql_check_patient)
                                    if len(sql_check_DNA_return_after_insert) == 1: # If this returns 1, the patient has successfully been booked in 
                                        print("Patient DNA test successfully booked into moka")
                                        message = "Patient DNA test successfully booked into moka"
                                        df['Booking_in_sample_status'] = message # add status to df
                                    else: 
                                        print("ERROR: Patient DNA not booked into moka")
                                        message = "ERROR: Patient DNA test not booked into moka"
                                        df['Booking_in_sample_status'] = message    
                                        # add to txt file 
            #print(df)
            # insert new bits into txt file for user feedback 
        '''======================= Update txt file with messages =============================================== '''   
        message ='Processing completed please check outputs for words ERROR for errors'
        df['Processed'] = message 
        df.to_csv(stripped_txt+'.txt', sep ='\t', index = False)  
        df_check = pd.read_csv(txt_file, delimiter = "\t")
        
        print(txt_file)
        print(df_check) # Re open the txt file to check if everything's been written to it 
        if 'Patient_Moka_status' in df_check.columns:
            pass # everything's been save to the txt file don't need to anything else 
        else:
            print('back up save')
            message ='As the original file was still open and couldn't be written too, this additional file was made. If the original file has since been closed, the text should now appear '
            df_check['Back_up'] = message 
            df_check.to_csv(stripped_txt+'_log.txt', sep ='\t', index = False)  # print to a new file incase the original one is open
'''
except: 
    print('AN ERROR HAS OCCURRED')
    message ='ERROR OCCURRED IN PROCESSING. Patient first sample set to complete?'
    df['Processed'] = message 
    df.to_csv(stripped_txt+'.txt', sep ='\t', index = False)   
    df.to_csv(stripped_txt+'_log.txt', sep ='\t', index = False)
