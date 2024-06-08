import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import *
import datetime as dt
import string
#d_parser = lambda x: pd.datetime.strptime(x, '%Y-%m-%d-%I-%p')

def CSV_to_Excel():
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')))
        df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
        df.insert(0, 'S/N.', df.index + 1)
        df.insert(4, 'Patient id', df.index + 1)
        df['Household Unique No']=''
        df['Received OVC Service?']=''
        df.loc[df['Sex']=="M",'Sex']='Male'
        df.loc[df['Sex']=="F",'Sex']='Female'
        df['DaysOfARVRefil']=df['DaysOfARVRefil'].astype('float')
        df['DaysOfARVRefil'] = (df['DaysOfARVRefil'] / 30).round(1)
        df['TPT Type']=''
        df['TPT Completion date (yyyy-mm-dd)']=''
        df['Date of Regimen Switch/ Substitution (yyyy-mm-dd)']=''
        df['Date of Full Disclosure (yyyy-mm-dd)']=''
        df['Number of Support Group (OTZ Club) meeting attended']=''
        df['Number of OTZ Modules completed']=''
        df['VL Result After VL Sample Collection (c/ml)']=''
        df['Date of VL Result After VL Sample Collection (yyyy-mm-dd)']=''
        df['Status at Registration']=''
        #df['VL Result After VL Sample Collection (c/ml)']=''
        df['RTT']=''
        df['If Dead, Cause of Dead']=''
        df['VA Cause of Dead']=''
        df['If Transferred out, new facility']=''
        df['Reason for Transfer-Out / IIT / Stooped Treatment']=''
        df['ART Enrollment Setting']=''
        df['Date Commenced DMOC (yyyy-mm-dd)']=''
        df['Type of DMOC']=''
        df['Date of Return of DMOC Client to Facility (yyyy-mm-dd)']=''
        df['Date of Commencement of EAC (yyyy-mm-dd)']=''
        df['Number of EAC Sessions Completed']=''
        df['Date of 3rd EAC Completion (yyyy-mm-dd)']=''
        df['Date of Extended EAC Completion (yyyy-mm-dd)']=''
        df['Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)']=''
        df['Co-morbidities']=''
        df['Date of Cervical Cancer Screening (yyyy-mm-dd)']=''
        df['Cervical Cancer Screening Type']=''
        df['Cervical Cancer Screening Method']=''
        df['Result of Cervical Cancer Screening']=''
        df['Date of Precancerous Lesions Treatment (yyyy-mm-dd)']=''
        df['Precancerous Lesions Treatment Methods']=''
        df['IIT Chance (%)']=''
        df['Date calculated (yyyy-mm-dd)']=''
        df['Case Manager']=''
        
        #rearrange columns
        df = df[['S/N.',
                    'State',
                    'LGA',
                    'FacilityName',
                    'Patient id',
                    'PatientHospitalNo',
                    'PatientUniqueID',
                    'Household Unique No',
                    'Received OVC Service?',
                    'Sex',
                    'CurrentWeight(Kg)',
                    'DateOfBirth',
                    'ARTStartDate',
                    'LastPickupDate',
                    'DaysOfARVRefil',
                     'LastINHDispensedDate',
                     'TPT Type',
                     'TPT Completion date (yyyy-mm-dd)',
                     'InitialRegimenLine',
                     'InitialRegimen',
                     'CurrentRegimenLine',
                     'CurrentRegimen',
                     'Date of Regimen Switch/ Substitution (yyyy-mm-dd)',
                     'PregnancyStatus',
                     'Date of Full Disclosure (yyyy-mm-dd)',
                     'OTZEnrollmentDate',
                     'Number of Support Group (OTZ Club) meeting attended',
                     'Number of OTZ Modules completed',
                     'ViralLoadSampleCollectionDate',
                     'CurrentViralLoad(c/ml)',
                     'ViralLoadEncounterDate',
                     'ViralLoadIndication',
                     'VL Result After VL Sample Collection (c/ml)',
                     'Date of VL Result After VL Sample Collection (yyyy-mm-dd)',
                     'Status at Registration',
                     'EnrollmentDate',
                     'ARTStatusPreviousQuarter',
                     'PatientOutcomeDatePreviousQuarter',
                     'CurrentARTStatusWithPillBalance',
                     'PatientOutcomeDate',
                    'RTT',
                     #
                    'If Dead, Cause of Dead',
                    'VA Cause of Dead',
                    'If Transferred out, new facility',
                    'Reason for Transfer-Out / IIT / Stooped Treatment',
                    'ART Enrollment Setting',
                    'Date Commenced DMOC (yyyy-mm-dd)',
                    'Type of DMOC',
                    'Date of Return of DMOC Client to Facility (yyyy-mm-dd)',
                    'Date of Commencement of EAC (yyyy-mm-dd)',
                    'Number of EAC Sessions Completed',
                    'Date of 3rd EAC Completion (yyyy-mm-dd)',
                    'Date of Extended EAC Completion (yyyy-mm-dd)',
                    'Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)',
                    'Co-morbidities',
                    'Date of Cervical Cancer Screening (yyyy-mm-dd)',
                    'Cervical Cancer Screening Type',
                    'Cervical Cancer Screening Method',
                    'Result of Cervical Cancer Screening',
                    'Date of Precancerous Lesions Treatment (yyyy-mm-dd)',
                    'Precancerous Lesions Treatment Methods',
                    'BiometricCaptureDate',
                    'ValidCapture',
                    'IIT Chance (%)',
                    'Date calculated (yyyy-mm-dd)',
                    'Case Manager']]
        
        #Convert Date Objects to Date
        df['DateOfBirth'] = pd.to_datetime(df['DateOfBirth']).dt.date
        df['ARTStartDate'] = pd.to_datetime(df['ARTStartDate']).dt.date
        df['LastPickupDate'] = pd.to_datetime(df['LastPickupDate']).dt.date
        df['LastINHDispensedDate'] = pd.to_datetime(df['LastINHDispensedDate']).dt.date
        df['TPT Completion date (yyyy-mm-dd)'] = pd.to_datetime(df['TPT Completion date (yyyy-mm-dd)']).dt.date
        df['Date of Regimen Switch/ Substitution (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Regimen Switch/ Substitution (yyyy-mm-dd)']).dt.date
        df['Date of Full Disclosure (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Full Disclosure (yyyy-mm-dd)']).dt.date
        df['ViralLoadSampleCollectionDate'] = pd.to_datetime(df['ViralLoadSampleCollectionDate']).dt.date
        df['ViralLoadEncounterDate'] = pd.to_datetime(df['ViralLoadEncounterDate']).dt.date
        df['Date of VL Result After VL Sample Collection (yyyy-mm-dd)'] = pd.to_datetime(df['Date of VL Result After VL Sample Collection (yyyy-mm-dd)']).dt.date
        df['EnrollmentDate'] = pd.to_datetime(df['EnrollmentDate']).dt.date
        df['PatientOutcomeDatePreviousQuarter'] = pd.to_datetime(df['PatientOutcomeDatePreviousQuarter']).dt.date
        df['PatientOutcomeDate'] = pd.to_datetime(df['PatientOutcomeDate']).dt.date
        df['Date Commenced DMOC (yyyy-mm-dd)'] = pd.to_datetime(df['Date Commenced DMOC (yyyy-mm-dd)']).dt.date
        df['Date of Return of DMOC Client to Facility (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Return of DMOC Client to Facility (yyyy-mm-dd)']).dt.date
        df['Date of Commencement of EAC (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Commencement of EAC (yyyy-mm-dd)']).dt.date
        df['Date of 3rd EAC Completion (yyyy-mm-dd)'] = pd.to_datetime(df['Date of 3rd EAC Completion (yyyy-mm-dd)']).dt.date
        df['Date of Extended EAC Completion (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Extended EAC Completion (yyyy-mm-dd)']).dt.date
        df['Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)']).dt.date
        df['Date of Cervical Cancer Screening (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Cervical Cancer Screening (yyyy-mm-dd)']).dt.date
        df['Date of Precancerous Lesions Treatment (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Precancerous Lesions Treatment (yyyy-mm-dd)']).dt.date
        df['BiometricCaptureDate'] = pd.to_datetime(df['BiometricCaptureDate']).dt.date
        df['Date calculated (yyyy-mm-dd)'] = pd.to_datetime(df['Date calculated (yyyy-mm-dd)']).dt.date
        
        
        #rename columns
        df.columns = ['S/No.',
                      'State',
                      'LGA',
                      'Facility',
                      'Patient id',
                      'Hospital Number',
                      'Unique ID',
                      'Household Unique No',
                      'Received OVC Service?',
                      'Sex',
                      'Current Weight (Kg)',
                      'Date of Birth (yyyy-mm-dd)',
                      'ART Start Date (yyyy-mm-dd)',
                      'Last Pickup Date (yyyy-mm-dd)',
                      'Months of ARV Refill',
                      'Date of TPT Start (yyyy-mm-dd)',
                      'TPT Type',
                      'TPT Completion date (yyyy-mm-dd)',
                      'Regimen Line at ART Start',
                      'Regimen at ART Start',
                      'Current Regimen Line',
                      'Current ART Regimen',
                      'Date of Regimen Switch/ Substitution (yyyy-mm-dd)',
                      'Pregnancy Status','Date of Full Disclosure (yyyy-mm-dd)',
                      'Date Enrolled on OTZ (yyyy-mm-dd)',
                      'Number of Support Group (OTZ Club) meeting attended',
                      'Number of OTZ Modules completed',
                      'Date of Viral Load Sample Collection (yyyy-mm-dd)',
                      'Current Viral Load (c/ml)',
                      'Date of Current Viral Load (yyyy-mm-dd)',
                      'Viral Load Indication',
                      'VL Result After VL Sample Collection (c/ml)',
                      'Date of VL Result After VL Sample Collection (yyyy-mm-dd)',
                      'Status at Registration',
                      'Date of Enrollment/Transfer-In (yyyy-mm-dd)',
                      'Previous ART Status',
                      'Confirmed Date of Previous ART Status (yyyy-mm-dd)',
                      'Current ART Status',
                      'Date of Current ART Status (yyyy-mm-dd)',
                      'RTT',
                        'If Dead, Cause of Dead',
                        'VA Cause of Dead',
                        'If Transferred out, new facility',
                        'Reason for Transfer-Out / IIT / Stooped Treatment',
                        'ART Enrollment Setting',
                        'Date Commenced DMOC (yyyy-mm-dd)',
                        'Type of DMOC',
                        'Date of Return of DMOC Client to Facility (yyyy-mm-dd)',
                        'Date of Commencement of EAC (yyyy-mm-dd)',
                        'Number of EAC Sessions Completed',
                        'Date of 3rd EAC Completion (yyyy-mm-dd)',
                        'Date of Extended EAC Completion (yyyy-mm-dd)',
                        'Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)',
                        'Co-morbidities',
                        'Date of Cervical Cancer Screening (yyyy-mm-dd)',
                        'Cervical Cancer Screening Type',
                        'Cervical Cancer Screening Method',
                        'Result of Cervical Cancer Screening',
                        'Date of Precancerous Lesions Treatment (yyyy-mm-dd)',
                        'Precancerous Lesions Treatment Methods',
                        'Date Biometrics Enrolled (yyyy-mm-dd)',
                        'Valid Biometrics Enrolled?',
                        'IIT Chance (%)',
                        'Date calculated (yyyy-mm-dd)',
                        'Case Manager']
        
        #format and export
        output_file_name = input_file_path.split("/")[-1][:-4]
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")
        df.to_excel(writer, sheet_name="NMRS-RADET", startrow=1, header=False, index=False)
        
        workbook = writer.book
        worksheet = writer.sheets["NMRS-RADET"]
        
        # Add a header format.
        header_format = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "bottom",
                "fg_color": "#D7E4BC",
                "border": 1,
            }
        )
        
        # Write the column headers with the defined format.
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num + 0, value, header_format)
        
        # Close the Pandas Excel writer and output the Excel file.
        writer.close()

def BASELINE_FILE():

        input_Bseline = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')))
        dfbaseline = pd.read_excel(input_Bseline,sheet_name=0, dtype=object)
            
        
        
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')))


        dfbaseline = pd.read_excel(input_Bseline,sheet_name=0, dtype=object)        
        df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
        dfbaseline['unique identifiers'] = dfbaseline["LGA"].astype(str) + dfbaseline["Facility"].astype(str) + dfbaseline["Hospital Number"].astype(str) + dfbaseline["Unique ID"].astype(str)
        df['unique identifiers'] = df["LGA"].astype(str) + df["FacilityName"].astype(str) + df["PatientHospitalNo"].astype(str) + df["PatientUniqueID"].astype(str)
        
        #remove duplicates
        dfbaseline = dfbaseline.drop_duplicates(subset=['unique identifiers'], keep=False)
        d = dict(enumerate(string.ascii_uppercase))
        m = df.duplicated(['unique identifiers'], keep=False)
        df.loc[m, 'unique identifiers'] += '_' + df[m].groupby(['unique identifiers']).cumcount().map(d)
        
        df.insert(0, 'S/N.', df.index + 1)
        df['Patient id']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Patient ID'])
        df['Patient id'] = df['Patient id'].fillna(df['unique identifiers'])
        #df.insert(4, 'Patient id', df.index + 1)
        df['Household Unique No']=''
        df['Received OVC Service?']=''
        df.loc[df['Sex']=="M",'Sex']='Male'
        df.loc[df['Sex']=="F",'Sex']='Female'
        df['DaysOfARVRefil']=df['DaysOfARVRefil'].astype('float')
        df['DaysOfARVRefil'] = (df['DaysOfARVRefil'] / 30).round(1)
        df['TPT Type']=''
        df['TPT Completion date (yyyy-mm-dd)']=''
        df['Date of Regimen Switch/ Substitution (yyyy-mm-dd)']=''
        df['Date of Full Disclosure (yyyy-mm-dd)']=''
        df['Number of Support Group (OTZ Club) meeting attended']=''
        df['Number of OTZ Modules completed']=''
        df['VL Result After VL Sample Collection (c/ml)']=''
        df['Date of VL Result After VL Sample Collection (yyyy-mm-dd)']=''
        df['Status at Registration']=''
        #df['VL Result After VL Sample Collection (c/ml)']=''
        df['RTT']=''
        df['If Dead, Cause of Dead']=''
        df['VA Cause of Dead']=''
        df['If Transferred out, new facility']=''
        df['Reason for Transfer-Out / IIT / Stooped Treatment']=''
        df['ART Enrollment Setting']=''
        df['Date Commenced DMOC (yyyy-mm-dd)']=''
        df['Type of DMOC']=''
        df['Date of Return of DMOC Client to Facility (yyyy-mm-dd)']=''
        df['Date of Commencement of EAC (yyyy-mm-dd)']=''
        df['Number of EAC Sessions Completed']=''
        df['Date of 3rd EAC Completion (yyyy-mm-dd)']=''
        df['Date of Extended EAC Completion (yyyy-mm-dd)']=''
        df['Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)']=''
        df['Co-morbidities']=''
        df['Date of Cervical Cancer Screening (yyyy-mm-dd)']=''
        df['Cervical Cancer Screening Type']=''
        df['Cervical Cancer Screening Method']=''
        df['Result of Cervical Cancer Screening']=''
        df['Date of Precancerous Lesions Treatment (yyyy-mm-dd)']=''
        df['Precancerous Lesions Treatment Methods']=''
        df['IIT Chance (%)']=''
        df['Date calculated (yyyy-mm-dd)']=''
        df['Case Manager']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Case Manager'])
        
        #rearrange columns
        df = df[['S/N.',
                    'State',
                    'LGA',
                    'FacilityName',
                    'Patient id',
                    'PatientHospitalNo',
                    'PatientUniqueID',
                    'Household Unique No',
                    'Received OVC Service?',
                    'Sex',
                    'CurrentWeight(Kg)',
                    'DateOfBirth',
                    'ARTStartDate',
                    'LastPickupDate',
                    'DaysOfARVRefil',
                     'LastINHDispensedDate',
                     'TPT Type',
                     'TPT Completion date (yyyy-mm-dd)',
                     'InitialRegimenLine',
                     'InitialRegimen',
                     'CurrentRegimenLine',
                     'CurrentRegimen',
                     'Date of Regimen Switch/ Substitution (yyyy-mm-dd)',
                     'PregnancyStatus',
                     'Date of Full Disclosure (yyyy-mm-dd)',
                     'OTZEnrollmentDate',
                     'Number of Support Group (OTZ Club) meeting attended',
                     'Number of OTZ Modules completed',
                     'ViralLoadSampleCollectionDate',
                     'CurrentViralLoad(c/ml)',
                     'ViralLoadEncounterDate',
                     'ViralLoadIndication',
                     'VL Result After VL Sample Collection (c/ml)',
                     'Date of VL Result After VL Sample Collection (yyyy-mm-dd)',
                     'Status at Registration',
                     'EnrollmentDate',
                     'ARTStatusPreviousQuarter',
                     'PatientOutcomeDatePreviousQuarter',
                     'CurrentARTStatusWithPillBalance',
                     'PatientOutcomeDate',
                    'RTT',
                     #
                    'If Dead, Cause of Dead',
                    'VA Cause of Dead',
                    'If Transferred out, new facility',
                    'Reason for Transfer-Out / IIT / Stooped Treatment',
                    'ART Enrollment Setting',
                    'Date Commenced DMOC (yyyy-mm-dd)',
                    'Type of DMOC',
                    'Date of Return of DMOC Client to Facility (yyyy-mm-dd)',
                    'Date of Commencement of EAC (yyyy-mm-dd)',
                    'Number of EAC Sessions Completed',
                    'Date of 3rd EAC Completion (yyyy-mm-dd)',
                    'Date of Extended EAC Completion (yyyy-mm-dd)',
                    'Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)',
                    'Co-morbidities',
                    'Date of Cervical Cancer Screening (yyyy-mm-dd)',
                    'Cervical Cancer Screening Type',
                    'Cervical Cancer Screening Method',
                    'Result of Cervical Cancer Screening',
                    'Date of Precancerous Lesions Treatment (yyyy-mm-dd)',
                    'Precancerous Lesions Treatment Methods',
                    'BiometricCaptureDate',
                    'ValidCapture',
                    'IIT Chance (%)',
                    'Date calculated (yyyy-mm-dd)',
                    'Case Manager']]
        
        #Convert Date Objects to Date
        df['DateOfBirth'] = pd.to_datetime(df['DateOfBirth']).dt.date
        df['ARTStartDate'] = pd.to_datetime(df['ARTStartDate']).dt.date
        df['LastPickupDate'] = pd.to_datetime(df['LastPickupDate']).dt.date
        df['LastINHDispensedDate'] = pd.to_datetime(df['LastINHDispensedDate']).dt.date
        df['TPT Completion date (yyyy-mm-dd)'] = pd.to_datetime(df['TPT Completion date (yyyy-mm-dd)']).dt.date
        df['Date of Regimen Switch/ Substitution (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Regimen Switch/ Substitution (yyyy-mm-dd)']).dt.date
        df['Date of Full Disclosure (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Full Disclosure (yyyy-mm-dd)']).dt.date
        df['ViralLoadSampleCollectionDate'] = pd.to_datetime(df['ViralLoadSampleCollectionDate']).dt.date
        df['ViralLoadEncounterDate'] = pd.to_datetime(df['ViralLoadEncounterDate']).dt.date
        df['Date of VL Result After VL Sample Collection (yyyy-mm-dd)'] = pd.to_datetime(df['Date of VL Result After VL Sample Collection (yyyy-mm-dd)']).dt.date
        df['EnrollmentDate'] = pd.to_datetime(df['EnrollmentDate']).dt.date
        df['PatientOutcomeDatePreviousQuarter'] = pd.to_datetime(df['PatientOutcomeDatePreviousQuarter']).dt.date
        df['PatientOutcomeDate'] = pd.to_datetime(df['PatientOutcomeDate']).dt.date
        df['Date Commenced DMOC (yyyy-mm-dd)'] = pd.to_datetime(df['Date Commenced DMOC (yyyy-mm-dd)']).dt.date
        df['Date of Return of DMOC Client to Facility (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Return of DMOC Client to Facility (yyyy-mm-dd)']).dt.date
        df['Date of Commencement of EAC (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Commencement of EAC (yyyy-mm-dd)']).dt.date
        df['Date of 3rd EAC Completion (yyyy-mm-dd)'] = pd.to_datetime(df['Date of 3rd EAC Completion (yyyy-mm-dd)']).dt.date
        df['Date of Extended EAC Completion (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Extended EAC Completion (yyyy-mm-dd)']).dt.date
        df['Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)']).dt.date
        df['Date of Cervical Cancer Screening (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Cervical Cancer Screening (yyyy-mm-dd)']).dt.date
        df['Date of Precancerous Lesions Treatment (yyyy-mm-dd)'] = pd.to_datetime(df['Date of Precancerous Lesions Treatment (yyyy-mm-dd)']).dt.date
        df['BiometricCaptureDate'] = pd.to_datetime(df['BiometricCaptureDate']).dt.date
        df['Date calculated (yyyy-mm-dd)'] = pd.to_datetime(df['Date calculated (yyyy-mm-dd)']).dt.date
        
        
        #rename columns
        df.columns = ['S/No.',
                      'State',
                      'LGA',
                      'Facility',
                      'Patient id',
                      'Hospital Number',
                      'Unique ID',
                      'Household Unique No',
                      'Received OVC Service?',
                      'Sex',
                      'Current Weight (Kg)',
                      'Date of Birth (yyyy-mm-dd)',
                      'ART Start Date (yyyy-mm-dd)',
                      'Last Pickup Date (yyyy-mm-dd)',
                      'Months of ARV Refill',
                      'Date of TPT Start (yyyy-mm-dd)',
                      'TPT Type',
                      'TPT Completion date (yyyy-mm-dd)',
                      'Regimen Line at ART Start',
                      'Regimen at ART Start',
                      'Current Regimen Line',
                      'Current ART Regimen',
                      'Date of Regimen Switch/ Substitution (yyyy-mm-dd)',
                      'Pregnancy Status','Date of Full Disclosure (yyyy-mm-dd)',
                      'Date Enrolled on OTZ (yyyy-mm-dd)',
                      'Number of Support Group (OTZ Club) meeting attended',
                      'Number of OTZ Modules completed',
                      'Date of Viral Load Sample Collection (yyyy-mm-dd)',
                      'Current Viral Load (c/ml)',
                      'Date of Current Viral Load (yyyy-mm-dd)',
                      'Viral Load Indication',
                      'VL Result After VL Sample Collection (c/ml)',
                      'Date of VL Result After VL Sample Collection (yyyy-mm-dd)',
                      'Status at Registration',
                      'Date of Enrollment/Transfer-In (yyyy-mm-dd)',
                      'Previous ART Status',
                      'Confirmed Date of Previous ART Status (yyyy-mm-dd)',
                      'Current ART Status',
                      'Date of Current ART Status (yyyy-mm-dd)',
                      'RTT',
                        'If Dead, Cause of Dead',
                        'VA Cause of Dead',
                        'If Transferred out, new facility',
                        'Reason for Transfer-Out / IIT / Stooped Treatment',
                        'ART Enrollment Setting',
                        'Date Commenced DMOC (yyyy-mm-dd)',
                        'Type of DMOC',
                        'Date of Return of DMOC Client to Facility (yyyy-mm-dd)',
                        'Date of Commencement of EAC (yyyy-mm-dd)',
                        'Number of EAC Sessions Completed',
                        'Date of 3rd EAC Completion (yyyy-mm-dd)',
                        'Date of Extended EAC Completion (yyyy-mm-dd)',
                        'Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)',
                        'Co-morbidities',
                        'Date of Cervical Cancer Screening (yyyy-mm-dd)',
                        'Cervical Cancer Screening Type',
                        'Cervical Cancer Screening Method',
                        'Result of Cervical Cancer Screening',
                        'Date of Precancerous Lesions Treatment (yyyy-mm-dd)',
                        'Precancerous Lesions Treatment Methods',
                        'Date Biometrics Enrolled (yyyy-mm-dd)',
                        'Valid Biometrics Enrolled?',
                        'IIT Chance (%)',
                        'Date calculated (yyyy-mm-dd)',
                        'Case Manager']
        
        #format and export
        output_file_name = input_file_path.split("/")[-1][:-4]
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")
        df.to_excel(writer, sheet_name="NMRS-RADET", startrow=1, header=False, index=False)
        
        workbook = writer.book
        worksheet = writer.sheets["NMRS-RADET"]
        
        # Add a header format.
        header_format = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "bottom",
                "fg_color": "#D7E4BC",
                "border": 1,
            }
        )
        
        # Write the column headers with the defined format.
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num + 0, value, header_format)
        
        # Close the Pandas Excel writer and output the Excel file.
        writer.close()


# Creating Main Window
root = tk.Tk()
root.title("NETO's NMRS TO RADET CONVERTER V001")
root.geometry("600x400")
root.config(bg="#f0f0f0")

# Adding a Button to the Window
convert_button = tk.Button(root, text="SELECT FILE & CONVERT", command=CSV_to_Excel, font=("Helvetica", 14), bg="#4caf50", fg="#ffffff")
convert_button.pack(pady=10)
text1 = tk.Label(root, text="Output's empty patiend id! (requires only NMRS file)")
text1.pack(pady=1)
convert_button = tk.Button(root, text="SELECT FILE & CONVERT", command=BASELINE_FILE, font=("Helvetica", 14), bg="#4caf50", fg="#ffffff")
convert_button.pack(pady=10)
text1 = tk.Label(root, text="Output's Patient Id (requires: baseline Radet and NMRS file)")
text1.pack(pady=1)
convert_button1 = tk.Button(root, text="EXIT CONVERTER", command=root.destroy, font=("Helvetica", 14), bg="red", fg="#ffffff")
convert_button1.pack(pady=40)
text = tk.Label(root, text="Welcome! to NMRS to RADET Converter!")
text.pack(pady=1)
text3 = tk.Label(root, text="you will be prompted to select required files and the location you want to save the converted file")
text3.pack(pady=1)

text2 = tk.Label(root, text="Contacts: email: chinedum.pius@gmail.com, phone: +2348134453841")
text2.pack(pady=30)


# Adding File Dialog
filedialog = tk.filedialog 

# Running the GUI
root.mainloop()


#input_file_path = filedialog.askopenfilename()
#df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
#pyinstaller nmrs-radet.py