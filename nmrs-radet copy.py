import pandas as pd
import tkinter as tk
from tkinter import filedialog
import datetime as dt
#d_parser = lambda x: pd.datetime.strptime(x, '%Y-%m-%d-%I-%p')

input_file_path = filedialog.askopenfilename()
df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)

#insert missing columns
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
output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=output_file_name)
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