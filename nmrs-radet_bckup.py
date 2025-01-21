from tkinter import ttk
import time
import pandas as pd
import numpy as np
from tkcalendar import DateEntry
from datetime import datetime
from dateutil import parser
import tkinter as tk
from tkinter import filedialog
from tkinter import *
import string
import csv
import os

# Declare global variables
start_date = None
end_date = None

def get_selected_date():
    global start_date, end_date
    start_date = cal.get_date()
    end_date = cal2.get_date()
    start_date_label.config(text=f"Start Date: {start_date}")
    end_date_label.config(text=f"End Date: {end_date}")
    startDate = start_date
    endDate = end_date
    todayDate = endDate
    todayDate = str(todayDate)
    return startDate, endDate

#def parse_date(date):
    #try:
        #return parser.parse(date, fuzzy=True, ignoretz=True)
    #except (parser.ParserError, TypeError):
        #return pd.NaT
    
def parse_date(date):
    date_formats = ["%Y-%m-%d", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y", "%Y.%m.%d", "%Y-%b-%d"]  # Add necessary formats here
    for fmt in date_formats:
        try:
            return pd.to_datetime(date, format=fmt).date()
        except (ValueError, TypeError):
            continue
    try:
        return parser.parse(date, fuzzy=True, ignoretz=True).date()
    except (parser.ParserError, TypeError, ValueError):
        return pd.NaT
    
# lamis and nmrs facility names
emrData = {
    'STATE': ['Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra'],
    'LGA': ['Aguata', 'Aguata', 'Aguata', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra West', 'Anaocha', 'Anaocha', 'Anaocha', 'Anaocha', 'Anaocha', 'Anaocha', 'Awka North', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Ayamelum', 'Dunukofia', 'Ekwusigo', 'Ekwusigo', 'Ekwusigo', 'Ekwusigo', 'Ekwusigo', 'Idemili North', 'Idemili North', 'Idemili North', 'Idemili North', 'Idemili North', 'Idemili North', 'Idemili North', 'Idemili South', 'Idemili South', 'Ihiala', 'Ihiala', 'Njikoka', 'Nnewi North', 'Nnewi North', 'Nnewi North', 'Nnewi South', 'Nnewi South', 'Nnewi South', 'Nnewi North', 'Nnewi North', 'Ogbaru', 'Ogbaru', 'Onitsha North', 'Onitsha North', 'Onitsha North', 'Onitsha South', 'Orumba North', 'Orumba South', 'Orumba South', 'Oyi', 'Oyi', 'Oyi', 'Anambra East', 'Ayamelum'],
    'Name on Lamis': ['Catholic Visitation Hospital and Maternity, Umuchu', 'Ekwulobia General Hospital', 'Igboukwu Diocesan Hospital', 'Aguleri Immaculate Heart Hospital', 'an Enugwu Aguleri Primary Health Center', 'an Ikem Nando Primary Health Center', 'an Mama Eliza Maternity Home', 'an Otuocha Maternity and Child Hospital', 'an Umuoba-Anam Primary Health Center', 'Umueri Diocesan Hospital and Maternity', 'Umueri General Hospital', 'an Abube Agu Nando Primary Health Centre', 'an Rex Inversorium Hospital, Mmiata (Referral Health center Oroma-Etiti )', 'Adazi St Joseph Hospital', 'an Akwaeze Primary Health Centre', 'an Ichida Primary Health Centre', 'an neni 1 Primary Health Centre', 'an Neni Comprehensive Health Centre NAUTH', 'An Rise Clinic, Adazi Ani', 'Lancet Hospital LTD', 'an God is Able Maternity Home', 'an Nibo Primary Health Center', 'an Nise Primary Health Center', 'an Okpuno Primary Health Center', 'Anambra State One Stop Shop (OSS), Awka', 'Chuwuemeka Odumegwu Ojukwu University Teaching Hospital', 'Faith Hospital and Maternity, Awka', 'Isiagu Primary Health Centre', 'Regina Caeli Hospital', 'an Maranatha Caring Mission Hospital', 'Ukpo Comprehensive Health Centre', 'an Ichi Referral Health Centre', 'an Joint Hospital Ozubulu', 'an Ozubulu Referral Health Center', 'Evans Specialist Hospital', 'Orafite General Hospital', 'an Eziowelle Primary Health Center', 'an Madueke Memorial Hospital and Maternity', 'Immaculate Heart Hospital (Nkpor)', 'Iyi-Enu Hospital', 'Nkpor Crown Hospital and Maternity', 'Obosi St Martins Hospital', 'an Obinwane Hospital and Maternity', '''an Ojoto Uno Primary Health Centre
''', 'Oba Comprehensive Health Centre (Trauma Centre)', 'an Eziani Health Centre', 'Our Lady of Lourdes Hospital', 'General Hospital Enugu-Ukwu', 'an Otolo Umuenem Primary Health Centre', 'Nnamdi Azikiwe University Teaching Hospital', 'Nnewi Diocesan Hospital', 'Amichi Diocesan Hospital', 'Osumenyi Visitation Hospital', 'Ukpor General Hospital', 'an Accucare Foundation Hospital', 'an St Felix Hospital and Maternity', 'Okpoko St Lukes Diocesan Hospital', 'Patricia Memorial Hospital and Maternity Aka baby', 'Holy Rosary Hospital', 'Onitsha General Hospital', 'St Charles Borromeo Hospital', 'Onitsha Pieta Hospital', 'Community Hospital Oko', 'an Trinitas International Hospital Umuchukwu', 'an Umunze Immaculate Heart Hospital', 'an Awkuzu Primary Health Centre', 'an Ogbunike Primary Health Centre', 'Umunya Comprehensive Health Centre (NAUTH)', 'an Ogbu Primary Health Centre', 'an St. Joseph Hospital and Maternity, Ifite Ogwari'],
    'Name on NMRS': ['', 'Ekwulobia General Hospital', 'Igboukwu DiocesHospital','','','','','','', 'Umueri DiocesHospital and Maternity','','','Oroma-Etiti Referral Health Centre', 'Adazi St Joseph Hospital','','','', 'Neni Comprehensive Health Centre NAUTH', 'Rise Clinic, Adazi Ani', 'Lancet Hospital LTD','','','','', 'Anambra State One Stop Shop (OSS) Site, Awka', 'Chuwuemeka Odumegwu Ojukwu University Teaching Hospital', 'Faith Hospital and Maternity, Awka','', 'Regina Caeli Hospital -  Awka', 'Maranatha Caring Mission Hospital', 'Ukpo Comprehensive Health Centre','','','', 'Evans Specialist Hospital', 'Orafite General Hospital','','', 'Immaculate Heart Hospital (Nkpor)', 'Iyi-Enu Hospital', 'Nkpor Crown Hospital and Maternity', "Obosi St Martin's Hospital",'', 'Ojotu Uno Primary Health Centre', 'Oba Comprehensive Health Centre (Trauma Centre)','', 'Our Lady of Lourdes Hospital', 'Enugu-Ukwu General Hospital', 'Otolo Umuenem Primary Health Centre','', 'Nnewi Diocesan Hospital', 'Amichi Diocesan Hospital', 'Osumenyi Visitation Hospital', 'Ukpor General Hospital','','', "Okpoko St Luke's Diocesan Hospital", 'Patricia Memorial Hospital and Maternity', 'Holy Rosary Hospital','','', 'Onitsha Pieta Hospital', 'Community Hospital Oko','', 'Umunze Immaculate Heart Hospital','','', 'Umunya Comprehensive Health Centre (NAUTH)','', 'St. Joseph Hospital and Maternity, Ifite Ogwari'],
    'NMRS Use Status': ['NO', 'Yes', 'Yes', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'Yes', 'NO', 'NO', 'Yes', 'Yes', 'NO', 'NO', 'NO', 'Yes', 'Yes', 'Yes', 'NO', 'NO', 'NO', 'NO', 'Yes', 'Yes', 'Yes', 'NO', 'Yes', 'Yes', 'Yes', 'NO', 'NO', 'NO', 'Yes', 'Yes', 'NO', 'NO', 'Yes', 'Yes', 'Yes', 'Yes', 'NO', 'Yes', 'Yes', 'NO', 'Yes', 'Yes', 'Yes', 'NO', 'Yes', 'Yes', 'Yes', 'Yes', 'NO', 'NO', 'Yes', 'Yes', 'Yes', 'NO', 'NO', 'Yes', 'Yes', 'NO', 'Yes', 'NO', 'NO', 'Yes', 'NO', 'Yes']
}
dflamisnmrs = pd.DataFrame(emrData)

def cleandflamisnmrs():
    global dflamisnmrs
    dflamisnmrs = dflamisnmrs[(dflamisnmrs != '').all(axis=1)]
    dflamisnmrs.to_excel('dflamisnmrs.xlsx')
    return dflamisnmrs

def lineListToRatet():
        startDate, endDate = get_selected_date()
        start_time = time.time()  # Start time measurement
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),                                                                    
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        #df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
            df = pd.read_csv(input_file_path, dtype=object, quoting=csv.QUOTE_NONE, on_bad_lines='skip')
    
            #Clean strings
            dfstr = ['State','LGA','DatimCode','FacilityName','PatientUniqueID','PatientHospitalNo','ANCNoIdentifier','ANCNoConceptID','HTSNo','Sex','AgeAtStartOfARTYears','AgeAtStartOfARTMonths','CareEntryPoint','HIVConfirmedDate','KPType','MonthsOnART','DateTransferredIn','TransferInStatus','ARTStartDate','LastPickupDate','LastVisitDate','PillBalance','InitialRegimenLine','InitialRegimen','InitialCD4Count','InitialCD4CountDate','CurrentCD4Count','CurrentCD4CountDate','LastEACDate','CurrentRegimenLine','CurrentRegimen','PregnancyStatus','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','GestationAgeWeeks','CurrentViralLoad(c/ml)','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','ViralLoadIndication','PatientOutcome','PatientOutcomeDate','CurrentARTStatus','DispensingModality','FacilityDispensingModality','DDDDispensingModality','MMDType','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','CurrentAgeYears','CurrentAgeMonths','DateOfBirth','MarkAsDeseased','MarkAsDeseasedDeathDate','RegistrationPhoneNo','NextofKinPhoneNo','TreatmentSupporterPhoneNo','BiometricCaptured','BiometricCaptureDate','ValidCapture','CurrentWeight(Kg)','CurrentWeightDate','TBStatus','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcome','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimen','InitialFirstLineRegimenDate','InitialSecondLineRegimen','InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','DrugDurationPreviousQuarter','PatientOutcomePreviousQuarter','PatientOutcomeDatePreviousQuarter','ARTStatusPreviousQuarter','QuantityOfARVDispensedLastVisit','FrequencyOfARVDispensedLastVisit','CurrentARTStatusWithPillBalance','RecaptureDate','RecaptureCount']
            for col in dfstr:
                df[col] = df[col].str.replace('\"', '')
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['HIVConfirmedDate','DateTransferredIn','ARTStartDate','LastPickupDate','LastVisitDate','InitialCD4CountDate','CurrentCD4CountDate','LastEACDate','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','PatientOutcomeDate','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','DateOfBirth','MarkAsDeseasedDeathDate','BiometricCaptureDate','CurrentWeightDate','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimenDate', 'InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','PatientOutcomeDatePreviousQuarter','RecaptureDate']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartOfARTYears', 'MonthsOnART','DaysOfARVRefil','GestationAgeWeeks','CurrentViralLoad(c/ml)','CurrentAgeYears','CurrentWeight(Kg)','DrugDurationPreviousQuarter','QuantityOfARVDispensedLastVisit',]
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['HIVConfirmedDate','DateTransferredIn','ARTStartDate','LastPickupDate','LastVisitDate','InitialCD4CountDate','CurrentCD4CountDate','LastEACDate','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','PatientOutcomeDate','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','DateOfBirth','MarkAsDeseasedDeathDate','BiometricCaptureDate','CurrentWeightDate','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimenDate', 'InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','PatientOutcomeDatePreviousQuarter','RecaptureDate']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartOfARTYears', 'MonthsOnART','DaysOfARVRefil','GestationAgeWeeks','CurrentViralLoad(c/ml)','CurrentAgeYears','CurrentWeight(Kg)','DrugDurationPreviousQuarter','QuantityOfARVDispensedLastVisit',]
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce') 
                
        #Add state and LGA from EMR data if it's blank
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        df = df.drop(['LGA2','STATE2'], axis=1)     
            
            
        df.insert(0, 'S/N.', df.index + 1)
        #df.insert(4, 'Patient id', df.index + 1)
        df['Patient id']=''
        df['Household Unique No']=''
        df['Received OVC Service?']=''
        df.loc[df['Sex']=="M",'Sex']='Male'
        df.loc[df['Sex']=="F",'Sex']='Female'
        df['IIT DATE'] = pd.to_datetime(df['LastPickupDate'], errors='coerce') + pd.to_timedelta(pd.to_numeric(df['DaysOfARVRefil']), unit='D') + pd.Timedelta(days=29)
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
        df.loc[(df['CurrentARTStatusWithPillBalance'] == 'InActive') & (df['PatientOutcomeDate'].isna()), 'PatientOutcomeDate'] = df['IIT DATE']
        df.loc[(df['CurrentARTStatusWithPillBalance'] == 'Active') & (df['PatientOutcomeDate'].isna()), 'PatientOutcomeDate'] = df['LastPickupDate']
        df.loc[df['CurrentARTStatus']=="Death",'CurrentARTStatus']='Dead'
        df.loc[df['CurrentARTStatus']=="Discontinued Care",'CurrentARTStatus']='Stopped'
        df.loc[df['CurrentARTStatus']=="Transferred out",'CurrentARTStatus']='Transferred Out'
        df.loc[df['CurrentARTStatus']=="LTFU",'CurrentARTStatus']='IIT'
        df.loc[df['CurrentARTStatusWithPillBalance']=="Death",'CurrentARTStatusWithPillBalance']='Dead'
        df.loc[df['CurrentARTStatusWithPillBalance']=="Discontinued Care",'CurrentARTStatusWithPillBalance']='Stopped'
        df.loc[df['CurrentARTStatusWithPillBalance']=="Transferred out",'CurrentARTStatusWithPillBalance']='Transferred Out'
        df.loc[df['CurrentARTStatusWithPillBalance']=="InActive",'CurrentARTStatusWithPillBalance']='IIT'
        df.loc[df['ARTStatusPreviousQuarter']=="LTFU",'ARTStatusPreviousQuarter']='IIT'
        df.loc[df['ARTStatusPreviousQuarter']=="Discontinued Care",'ARTStatusPreviousQuarter']='Stopped'
        df.loc[df['ARTStatusPreviousQuarter']=="Transferred out",'ARTStatusPreviousQuarter']='Transferred Out'
        df.loc[df['ARTStatusPreviousQuarter']=="Death",'ARTStatusPreviousQuarter']='Dead'
        
        #Clean regimen        
        df.loc[df['CurrentRegimenLine']=="Adult 1st line ARV regimen",'CurrentRegimenLine']='Adult.1st.Line'
        df.loc[df['CurrentRegimenLine']=="Adult 2nd line ARV regimen",'CurrentRegimenLine']='Adult.2nd.Line'
        df.loc[df['CurrentRegimenLine']=="Adult 3rd line ARV regimen",'CurrentRegimenLine']='Adult.3rd.Line'
        df.loc[df['CurrentRegimenLine']=="Child 1st line ARV regimen",'CurrentRegimenLine']='Peds.1st.Line'
        df.loc[df['CurrentRegimenLine']=="Child 2nd line ARV regimen",'CurrentRegimenLine']='Peds.2nd.Line'
        df.loc[df['CurrentRegimenLine']=="Child 3rd line ARV regimen",'CurrentRegimenLine']='Peds.3rd.Line'
        df.loc[df['InitialRegimenLine']=="Adult 1st line ARV regimen",'InitialRegimenLine']='Adult.1st.Line'
        df.loc[df['InitialRegimenLine']=="Adult 2nd line ARV regimen",'InitialRegimenLine']='Adult.2nd.Line'
        df.loc[df['InitialRegimenLine']=="Adult 3rd line ARV regimen",'InitialRegimenLine']='Adult.3rd.Line'
        df.loc[df['InitialRegimenLine']=="Child 1st line ARV regimen",'InitialRegimenLine']='Peds.1st.Line'
        df.loc[df['InitialRegimenLine']=="Child 2nd line ARV regimen",'InitialRegimenLine']='Peds.2nd.Line'
        df.loc[df['InitialRegimenLine']=="Child 3rd line ARV regimen",'InitialRegimenLine']='Peds.3rd.Line'
        
        # Convert VL columns to datetime
        VLResultDatecolumns = ['ViralLoadReportedDate', 'ResultDate', 'AssayDate', 'ApprovalDate']
        VLSampleDatecolumns = ['ViralLoadSampleCollectionDate', 'LastViralLoadSampleCollectionFormDate']
        df[VLResultDatecolumns] = df[VLResultDatecolumns].apply(pd.to_datetime, errors='coerce')
        df[VLSampleDatecolumns] = df[VLSampleDatecolumns].apply(pd.to_datetime, errors='coerce')

        # Find the maximum dates across specified columns
        maxVlResultdate = df[VLResultDatecolumns].max(axis=1)
        maxVlSCdate = df[VLSampleDatecolumns].max(axis=1)

        # Assign the results back to the DataFrame
        df.loc[df['CurrentViralLoad(c/ml)'].notnull(), 'Date of Current Viral Load (yyyy-mm-dd)'] = maxVlResultdate
        df['Date of Viral Load Sample Collection (yyyy-mm-dd)'] = maxVlSCdate
        
        # Calculate the difference between 'IIT DATE' and today's date
        daystoIIT = (df['IIT DATE'] - pd.to_datetime(endDate)).dt.days
        
        # Apply the formula to get IIT chance as a percentage
        df['IITChance'] = ((29 - daystoIIT) / 29)
        
        # Apply IIT chance to active clients
        df.loc[(((df['CurrentARTStatus']== "Active") | (df['CurrentARTStatus']== "Active(A)")) & (df['IITChance'] >= 0)),'IIT Chance (%)']=df['IITChance']
        df.loc[(((df['CurrentARTStatus']== "Active") | (df['CurrentARTStatus']== "Active(A)")) & (df['IITChance'] >= 0)),'Date calculated (yyyy-mm-dd)']=df['IIT DATE']
        
        #Add restarts, transfer-in and status at registration to current ART status in Radet
        df.loc[(df['DateReturnedToCare'].notnull()) & (df['CurrentARTStatus']== "Active"),'CurrentARTStatus']='Active-Restart'
        df.loc[(df['DateReturnedToCare'].notnull()) & ((df['CurrentARTStatus']== "Active") | (df['CurrentARTStatus'] == "Active-Restart")),'PatientOutcomeDate']=df['DateReturnedToCare']
        df.loc[(df['TransferInStatus'] == 'Transfer in with records') & ((df['CurrentARTStatus']== "Active")),'CurrentARTStatus']='Active-Transfer In'
        df.loc[(df['Status at Registration']=='') & (df['TransferInStatus'] == 'Transfer in with records'),'Status at Registration']='ART Transfer In'
    
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
                    'Date of Viral Load Sample Collection (yyyy-mm-dd)',
                    'CurrentViralLoad(c/ml)',
                    'Date of Current Viral Load (yyyy-mm-dd)',
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
                    'BiometricCaptured',
                    'IIT Chance (%)',
                    'Date calculated (yyyy-mm-dd)',
                    'Case Manager']]
        
        #Convert Date Objects to Date
        dfDates = ['DateOfBirth','ARTStartDate','LastPickupDate','LastINHDispensedDate','TPT Completion date (yyyy-mm-dd)','Date of Regimen Switch/ Substitution (yyyy-mm-dd)','Date of Full Disclosure (yyyy-mm-dd)','Date of Viral Load Sample Collection (yyyy-mm-dd)','Date of Current Viral Load (yyyy-mm-dd)','Date of VL Result After VL Sample Collection (yyyy-mm-dd)','EnrollmentDate','PatientOutcomeDatePreviousQuarter','PatientOutcomeDate','Date Commenced DMOC (yyyy-mm-dd)','Date of Return of DMOC Client to Facility (yyyy-mm-dd)','Date of Commencement of EAC (yyyy-mm-dd)','Date of 3rd EAC Completion (yyyy-mm-dd)','Date of Extended EAC Completion (yyyy-mm-dd)','Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)','Date of Cervical Cancer Screening (yyyy-mm-dd)','Date of Precancerous Lesions Treatment (yyyy-mm-dd)','BiometricCaptureDate','Date calculated (yyyy-mm-dd)']
        for col in dfDates:
            df[col] = pd.to_datetime(df[col],errors='coerce').dt.date
        
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

        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

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
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f" Converted File Location: {output_file_path}")

def lineListToRatet_Baseline():
        start_time = time.time()  # Start time measurement
        startDate, endDate = get_selected_date()
        status_label.config(text=f"Attempting to read file...")
        input_Bseline = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx'),
                                                                     ('excel file','*.xlsb')))
        if not input_Bseline:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        dfbaseline = pd.read_excel(input_Bseline,sheet_name=0, dtype=object)
            
        
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function

        dfbaseline = pd.read_excel(input_Bseline,sheet_name=0, dtype=object)        
        #df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
            df = pd.read_csv(input_file_path, dtype=object, quoting=csv.QUOTE_NONE, on_bad_lines='skip')
    
            #Clean strings
            dfstr = ['State','LGA','DatimCode','FacilityName','PatientUniqueID','PatientHospitalNo','ANCNoIdentifier','ANCNoConceptID','HTSNo','Sex','AgeAtStartOfARTYears','AgeAtStartOfARTMonths','CareEntryPoint','HIVConfirmedDate','KPType','MonthsOnART','DateTransferredIn','TransferInStatus','ARTStartDate','LastPickupDate','LastVisitDate','PillBalance','InitialRegimenLine','InitialRegimen','InitialCD4Count','InitialCD4CountDate','CurrentCD4Count','CurrentCD4CountDate','LastEACDate','CurrentRegimenLine','CurrentRegimen','PregnancyStatus','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','GestationAgeWeeks','CurrentViralLoad(c/ml)','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','ViralLoadIndication','PatientOutcome','PatientOutcomeDate','CurrentARTStatus','DispensingModality','FacilityDispensingModality','DDDDispensingModality','MMDType','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','CurrentAgeYears','CurrentAgeMonths','DateOfBirth','MarkAsDeseased','MarkAsDeseasedDeathDate','RegistrationPhoneNo','NextofKinPhoneNo','TreatmentSupporterPhoneNo','BiometricCaptured','BiometricCaptureDate','ValidCapture','CurrentWeight(Kg)','CurrentWeightDate','TBStatus','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcome','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimen','InitialFirstLineRegimenDate','InitialSecondLineRegimen','InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','DrugDurationPreviousQuarter','PatientOutcomePreviousQuarter','PatientOutcomeDatePreviousQuarter','ARTStatusPreviousQuarter','QuantityOfARVDispensedLastVisit','FrequencyOfARVDispensedLastVisit','CurrentARTStatusWithPillBalance','RecaptureDate','RecaptureCount']
            for col in dfstr:
                df[col] = df[col].str.replace('\"', '')
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['HIVConfirmedDate','DateTransferredIn','ARTStartDate','LastPickupDate','LastVisitDate','InitialCD4CountDate','CurrentCD4CountDate','LastEACDate','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','PatientOutcomeDate','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','DateOfBirth','MarkAsDeseasedDeathDate','BiometricCaptureDate','CurrentWeightDate','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimenDate', 'InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','PatientOutcomeDatePreviousQuarter','RecaptureDate']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartOfARTYears', 'MonthsOnART','DaysOfARVRefil','GestationAgeWeeks','CurrentViralLoad(c/ml)','CurrentAgeYears','CurrentWeight(Kg)','DrugDurationPreviousQuarter','QuantityOfARVDispensedLastVisit',]
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['HIVConfirmedDate','DateTransferredIn','ARTStartDate','LastPickupDate','LastVisitDate','InitialCD4CountDate','CurrentCD4CountDate','LastEACDate','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','PatientOutcomeDate','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','DateOfBirth','MarkAsDeseasedDeathDate','BiometricCaptureDate','CurrentWeightDate','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimenDate', 'InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','PatientOutcomeDatePreviousQuarter','RecaptureDate']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartOfARTYears', 'MonthsOnART','DaysOfARVRefil','GestationAgeWeeks','CurrentViralLoad(c/ml)','CurrentAgeYears','CurrentWeight(Kg)','DrugDurationPreviousQuarter','QuantityOfARVDispensedLastVisit',]
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
                
        #Add state and LGA from EMR data if it's blank and also change Facility name where NMRS and Lamis does not tally
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df['Name on Lamis']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['Name on Lamis'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        df.loc[df['Name on Lamis'] != df['FacilityName'], 'FacilityName'] = df['Name on Lamis']
        df = df.drop(['LGA2','STATE2','Name on Lamis'], axis=1)

        #Remove leading zeros from Hospital Numbers and Unique IDs that are digits 
        df['PatientHospitalNo1'] = df['PatientHospitalNo'].apply(lambda x: str(x).lstrip('0') if str(x).isdigit() else x)
        df['PatientUniqueID1'] = df['PatientUniqueID'].apply(lambda x: str(x).lstrip('0') if str(x).isdigit() else x)
        dfbaseline['Hospital Number1'] = dfbaseline['Hospital Number'].apply(lambda x: str(x).lstrip('0') if str(x).isdigit() else x)
        dfbaseline['Unique ID1'] = dfbaseline['Unique ID'].apply(lambda x: str(x).lstrip('0') if str(x).isdigit() else x)
        
        #Create Unique Identifiers
        dfbaseline['unique identifiers'] = dfbaseline["LGA"].astype(str) + dfbaseline["Facility"].astype(str) + dfbaseline["Hospital Number1"].astype(str) + dfbaseline["Unique ID1"].astype(str)
        df['unique identifiers'] = df["LGA"].astype(str) + df["FacilityName"].astype(str) + df["PatientHospitalNo1"].astype(str) + df["PatientUniqueID1"].astype(str)
        
        #remove duplicates
        dfbaseline = dfbaseline.drop_duplicates(subset=['unique identifiers'], keep=False)
        d = dict(enumerate(string.ascii_uppercase))
        m = df.duplicated(['unique identifiers'], keep=False)
        df.loc[m, 'unique identifiers'] += '_' + df[m].groupby(['unique identifiers']).cumcount().map(d)
        dfbaseline.to_excel("dfbaseline.xlsx")
        
        df.insert(0, 'S/N.', df.index + 1)
        df['Patient id']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Patient ID'])
        df['Patient id'] = df['Patient id'].fillna(df['unique identifiers'])
        #df.insert(4, 'Patient id', df.index + 1)
        df['Household Unique No']=''
        df['Received OVC Service?']=''
        df.loc[df['Sex']=="M",'Sex']='Male'
        df.loc[df['Sex']=="F",'Sex']='Female'
        df['IIT DATE'] = pd.to_datetime(df['LastPickupDate']) + pd.to_timedelta(pd.to_numeric(df['DaysOfARVRefil']), unit='D') + pd.Timedelta(days=29)
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
        df.loc[(df['CurrentARTStatusWithPillBalance'] == 'InActive') & (df['PatientOutcomeDate'].isna()), 'PatientOutcomeDate'] = df['IIT DATE']
        df.loc[(df['CurrentARTStatusWithPillBalance'] == 'Active') & (df['PatientOutcomeDate'].isna()), 'PatientOutcomeDate'] = df['LastPickupDate']
        df.loc[df['CurrentARTStatus']=="Death",'CurrentARTStatus']='Dead'
        df.loc[df['CurrentARTStatus']=="Discontinued Care",'CurrentARTStatus']='Stopped'
        df.loc[df['CurrentARTStatus']=="Transferred out",'CurrentARTStatus']='Transferred Out'
        df.loc[df['CurrentARTStatus']=="LTFU",'CurrentARTStatus']='IIT'
        df.loc[df['CurrentARTStatusWithPillBalance']=="Death",'CurrentARTStatusWithPillBalance']='Dead'
        df.loc[df['CurrentARTStatusWithPillBalance']=="Discontinued Care",'CurrentARTStatusWithPillBalance']='Stopped'
        df.loc[df['CurrentARTStatusWithPillBalance']=="Transferred out",'CurrentARTStatusWithPillBalance']='Transferred Out'
        df.loc[df['CurrentARTStatusWithPillBalance']=="InActive",'CurrentARTStatusWithPillBalance']='IIT'
        df.loc[df['ARTStatusPreviousQuarter']=="LTFU",'ARTStatusPreviousQuarter']='IIT'
        df.loc[df['ARTStatusPreviousQuarter']=="Discontinued Care",'ARTStatusPreviousQuarter']='Stopped'
        df.loc[df['ARTStatusPreviousQuarter']=="Transferred out",'ARTStatusPreviousQuarter']='Transferred Out'
        df.loc[df['ARTStatusPreviousQuarter']=="Death",'ARTStatusPreviousQuarter']='Dead'
        
        df.loc[df['CurrentRegimenLine']=="Adult 1st line ARV regimen",'CurrentRegimenLine']='Adult.1st.Line'
        df.loc[df['CurrentRegimenLine']=="Adult 2nd line ARV regimen",'CurrentRegimenLine']='Adult.2nd.Line'
        df.loc[df['CurrentRegimenLine']=="Adult 3rd line ARV regimen",'CurrentRegimenLine']='Adult.3rd.Line'
        df.loc[df['CurrentRegimenLine']=="Child 1st line ARV regimen",'CurrentRegimenLine']='Peds.1st.Line'
        df.loc[df['CurrentRegimenLine']=="Child 2nd line ARV regimen",'CurrentRegimenLine']='Peds.2nd.Line'
        df.loc[df['CurrentRegimenLine']=="Child 3rd line ARV regimen",'CurrentRegimenLine']='Peds.3rd.Line'
        df.loc[df['InitialRegimenLine']=="Adult 1st line ARV regimen",'InitialRegimenLine']='Adult.1st.Line'
        df.loc[df['InitialRegimenLine']=="Adult 2nd line ARV regimen",'InitialRegimenLine']='Adult.2nd.Line'
        df.loc[df['InitialRegimenLine']=="Adult 3rd line ARV regimen",'InitialRegimenLine']='Adult.3rd.Line'
        df.loc[df['InitialRegimenLine']=="Child 1st line ARV regimen",'InitialRegimenLine']='Peds.1st.Line'
        df.loc[df['InitialRegimenLine']=="Child 2nd line ARV regimen",'InitialRegimenLine']='Peds.2nd.Line'
        df.loc[df['InitialRegimenLine']=="Child 3rd line ARV regimen",'InitialRegimenLine']='Peds.3rd.Line'
              
        # Convert VL columns to datetime
        VLResultDatecolumns = ['ViralLoadReportedDate', 'ResultDate', 'AssayDate', 'ApprovalDate']
        VLSampleDatecolumns = ['ViralLoadSampleCollectionDate', 'LastViralLoadSampleCollectionFormDate']
        df[VLResultDatecolumns] = df[VLResultDatecolumns].apply(pd.to_datetime, errors='coerce')
        df[VLSampleDatecolumns] = df[VLSampleDatecolumns].apply(pd.to_datetime, errors='coerce')

        # Find the maximum dates across specified columns
        maxVlResultdate = df[VLResultDatecolumns].max(axis=1)
        maxVlSCdate = df[VLSampleDatecolumns].max(axis=1)

        # Assign the results back to the DataFrame
        df.loc[df['CurrentViralLoad(c/ml)'].notnull(), 'Date of Current Viral Load (yyyy-mm-dd)'] = maxVlResultdate
        df['Date of Viral Load Sample Collection (yyyy-mm-dd)'] = maxVlSCdate
        
        #Check active clients in baseline Radet
        refill_days = np.where(dfbaseline['Months of ARV Refill'] == "", 15, dfbaseline['Months of ARV Refill'] * 30)
        dfbaseline['IIT DATE_baseline'] = pd.to_datetime(dfbaseline['Last Pickup Date (yyyy-mm-dd)']) + pd.to_timedelta(refill_days, unit='D') + pd.Timedelta(days=29)
        dfbaseline['Revalidation status_baseline'] = np.where((dfbaseline['IIT DATE_baseline'] >= pd.to_datetime(endDate)) &
            ~dfbaseline['Current ART Status'].isin(["Dead", "Stopped", "Transferred Out", ""]),"Active","Inactive"
        )
        
        #insert columns in df from dfbaseline
        df['Revalidation status_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Revalidation status_baseline'])
        df['Current Viral Load (c/ml)_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Current Viral Load (c/ml)'])
        df['Date of Current Viral Load (yyyy-mm-dd)_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Date of Current Viral Load (yyyy-mm-dd)'])
        df['Date of Viral Load Sample Collection (yyyy-mm-dd)_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Date of Viral Load Sample Collection (yyyy-mm-dd)'])
        df['Case Manager_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Case Manager'])
        
        df['ART Enrollment Setting_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['ART Enrollment Setting'])
        df['Date Commenced DMOC (yyyy-mm-dd)_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Date Commenced DMOC (yyyy-mm-dd)'])
        df['Type of DMOC_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Type of DMOC'])
        df['Date of Return of DMOC Client to Facility (yyyy-mm-dd)_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Date of Return of DMOC Client to Facility (yyyy-mm-dd)'])
        df['Date of TPT Start (yyyy-mm-dd)_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Date of TPT Start (yyyy-mm-dd)'])
        df['TPT Type_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['TPT Type'])
        df['TPT Completion date (yyyy-mm-dd)_baseline']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['TPT Completion date (yyyy-mm-dd)'])
        df['Patient id']=df['unique identifiers'].map(dfbaseline.set_index('unique identifiers')['Patient ID'])
        #df['Patient id'] = df['Patient id'].fillna(df['unique identifiers'])
        
        #Assign columns from baseline Radet to converted Radet
        df['Case Manager'] = df['Case Manager_baseline']
        df['ART Enrollment Setting']=df['ART Enrollment Setting_baseline']
        df['Date Commenced DMOC (yyyy-mm-dd)']=df['Date Commenced DMOC (yyyy-mm-dd)_baseline']
        df['Type of DMOC']=df['Type of DMOC_baseline']
        df['Date of Return of DMOC Client to Facility (yyyy-mm-dd)']=df['Date of Return of DMOC Client to Facility (yyyy-mm-dd)_baseline']
        df['TPT Type']=df['TPT Type_baseline']
        df['TPT Completion date (yyyy-mm-dd)']=df['TPT Completion date (yyyy-mm-dd)_baseline']
        
        #compare and update more recent data from baseline RADET to NMRS if seen
        df['CurrentViralLoad(c/ml)'] = pd.to_numeric(df['CurrentViralLoad(c/ml)'], errors='coerce')
        df.loc[(df['CurrentARTStatus'] == 'LTFU') & (df['Revalidation status_baseline'] == 'Active'), 'CurrentARTStatus'] = "Active(A)"
        df.loc[((df['Date of Current Viral Load (yyyy-mm-dd)'].isna()) & (df["Date of Current Viral Load (yyyy-mm-dd)_baseline"].notnull())) | ((df["Date of Current Viral Load (yyyy-mm-dd)_baseline"] > df['Date of Current Viral Load (yyyy-mm-dd)']) & (df["Current Viral Load (c/ml)_baseline"].notnull())), 'CurrentViralLoad(c/ml)'] = df["Current Viral Load (c/ml)_baseline"]
        df.loc[((df['Date of Current Viral Load (yyyy-mm-dd)'].isna()) & (df["Date of Current Viral Load (yyyy-mm-dd)_baseline"].notnull())) | (df["Date of Current Viral Load (yyyy-mm-dd)_baseline"] > df['Date of Current Viral Load (yyyy-mm-dd)']), 'Date of Current Viral Load (yyyy-mm-dd)'] = df["Date of Current Viral Load (yyyy-mm-dd)_baseline"]
        df.loc[((df['Date of Viral Load Sample Collection (yyyy-mm-dd)'].isna()) & (df["Date of Viral Load Sample Collection (yyyy-mm-dd)_baseline"].notnull())) | (df["Date of Viral Load Sample Collection (yyyy-mm-dd)_baseline"] > df['Date of Viral Load Sample Collection (yyyy-mm-dd)']), 'Date of Viral Load Sample Collection (yyyy-mm-dd)'] = df["Date of Viral Load Sample Collection (yyyy-mm-dd)_baseline"]
        df.loc[((df['Date of Current Viral Load (yyyy-mm-dd)'].notnull()) & (df['CurrentViralLoad(c/ml)'] <= 1000)),'ViralLoadIndication']='Routine - Routine'
        df.loc[((df['Date of Current Viral Load (yyyy-mm-dd)'].notnull()) & (df['CurrentViralLoad(c/ml)'] > 1000)),'ViralLoadIndication']='Targeted - Post EAC'
        df.loc[((df['LastINHDispensedDate'].isna()) & (df["Date of TPT Start (yyyy-mm-dd)_baseline"].notnull())) | (df["Date of TPT Start (yyyy-mm-dd)_baseline"] > df['LastINHDispensedDate']), 'LastINHDispensedDate'] = df["Date of TPT Start (yyyy-mm-dd)_baseline"]
        #df.to_excel("test.xlsx")
        #dfbaseline.to_excel("test2.xlsx")
        
        #calculate IIT Chance
    
        # Calculate the difference between 'IIT DATE' and today's date
        #today = pd.to_datetime('today')
        daystoIIT = (df['IIT DATE'] - pd.to_datetime(end_date)).dt.days
        # Apply the formula to get IIT chance as a percentage
        df['IITChance'] = ((29 - daystoIIT) / 29)
        # Apply IIT chance to active clients
        df.loc[(((df['CurrentARTStatus']== "Active") | (df['CurrentARTStatus']== "Active(A)")) & (df['IITChance'] >= 0)),'IIT Chance (%)']=df['IITChance']
        df.loc[(((df['CurrentARTStatus']== "Active") | (df['CurrentARTStatus']== "Active(A)")) & (df['IITChance'] >= 0)),'Date calculated (yyyy-mm-dd)']=df['IIT DATE']
        
        #Add restarts, transfer-in and status at registration to current ART status in Radet
        df.loc[(df['DateReturnedToCare'].notnull()) & (df['CurrentARTStatus']== "Active"),'CurrentARTStatus']='Active-Restart'
        df.loc[(df['DateReturnedToCare'].notnull()) & ((df['CurrentARTStatus']== "Active") | (df['CurrentARTStatus'] == "Active-Restart")),'PatientOutcomeDate']=df['DateReturnedToCare']
        df.loc[(df['TransferInStatus'] == 'Transfer in with records') & ((df['CurrentARTStatus']== "Active")),'CurrentARTStatus']='Active-Transfer In'
        df.loc[(df['Status at Registration']=='') & (df['TransferInStatus'] == 'Transfer in with records'),'Status at Registration']='ART Transfer In'
    
    
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
                    'Date of Viral Load Sample Collection (yyyy-mm-dd)',
                    'CurrentViralLoad(c/ml)',
                    'Date of Current Viral Load (yyyy-mm-dd)',
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
                    'BiometricCaptured',
                    'IIT Chance (%)',
                    'Date calculated (yyyy-mm-dd)',
                    'Case Manager']]
        
        #Convert Date Objects to Date
        dfDates = ['DateOfBirth','ARTStartDate','LastPickupDate','LastINHDispensedDate','TPT Completion date (yyyy-mm-dd)','Date of Regimen Switch/ Substitution (yyyy-mm-dd)','Date of Full Disclosure (yyyy-mm-dd)','Date of Viral Load Sample Collection (yyyy-mm-dd)','Date of Current Viral Load (yyyy-mm-dd)','Date of VL Result After VL Sample Collection (yyyy-mm-dd)','EnrollmentDate','PatientOutcomeDatePreviousQuarter','PatientOutcomeDate','Date Commenced DMOC (yyyy-mm-dd)','Date of Return of DMOC Client to Facility (yyyy-mm-dd)','Date of Commencement of EAC (yyyy-mm-dd)','Date of 3rd EAC Completion (yyyy-mm-dd)','Date of Extended EAC Completion (yyyy-mm-dd)','Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)','Date of Cervical Cancer Screening (yyyy-mm-dd)','Date of Precancerous Lesions Treatment (yyyy-mm-dd)','BiometricCaptureDate','Date calculated (yyyy-mm-dd)']
        for col in dfDates:
            df[col] = pd.to_datetime(df[col],errors='coerce').dt.date
        
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

        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        #status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

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
        # Update the status label upon completion
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f" Converted File Location: {output_file_path}")
        

def CleanARTLineList():

        start_time = time.time()  # Start time measurement
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
            df = pd.read_csv(input_file_path, dtype=object, quoting=csv.QUOTE_NONE, on_bad_lines='skip')
    
            #Clean strings
            dfstr = ['State','LGA','DatimCode','FacilityName','PatientUniqueID','PatientHospitalNo','ANCNoIdentifier','ANCNoConceptID','HTSNo','Sex','AgeAtStartOfARTYears','AgeAtStartOfARTMonths','CareEntryPoint','HIVConfirmedDate','KPType','MonthsOnART','DateTransferredIn','TransferInStatus','ARTStartDate','LastPickupDate','LastVisitDate','PillBalance','InitialRegimenLine','InitialRegimen','InitialCD4Count','InitialCD4CountDate','CurrentCD4Count','CurrentCD4CountDate','LastEACDate','CurrentRegimenLine','CurrentRegimen','PregnancyStatus','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','GestationAgeWeeks','CurrentViralLoad(c/ml)','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','ViralLoadIndication','PatientOutcome','PatientOutcomeDate','CurrentARTStatus','DispensingModality','FacilityDispensingModality','DDDDispensingModality','MMDType','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','CurrentAgeYears','CurrentAgeMonths','DateOfBirth','MarkAsDeseased','MarkAsDeseasedDeathDate','RegistrationPhoneNo','NextofKinPhoneNo','TreatmentSupporterPhoneNo','BiometricCaptured','BiometricCaptureDate','ValidCapture','CurrentWeight(Kg)','CurrentWeightDate','TBStatus','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcome','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimen','InitialFirstLineRegimenDate','InitialSecondLineRegimen','InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','DrugDurationPreviousQuarter','PatientOutcomePreviousQuarter','PatientOutcomeDatePreviousQuarter','ARTStatusPreviousQuarter','QuantityOfARVDispensedLastVisit','FrequencyOfARVDispensedLastVisit','CurrentARTStatusWithPillBalance','RecaptureDate','RecaptureCount']
            for col in dfstr:
                df[col] = df[col].str.replace('\"', '')
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['HIVConfirmedDate','DateTransferredIn','ARTStartDate','LastPickupDate','LastVisitDate','InitialCD4CountDate','CurrentCD4CountDate','LastEACDate','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','PatientOutcomeDate','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','DateOfBirth','MarkAsDeseasedDeathDate','BiometricCaptureDate','CurrentWeightDate','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimenDate', 'InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','PatientOutcomeDatePreviousQuarter','RecaptureDate']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartOfARTYears', 'MonthsOnART','DaysOfARVRefil','GestationAgeWeeks','CurrentViralLoad(c/ml)','CurrentAgeYears','CurrentWeight(Kg)','DrugDurationPreviousQuarter','QuantityOfARVDispensedLastVisit',]
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['HIVConfirmedDate','DateTransferredIn','ARTStartDate','LastPickupDate','LastVisitDate','InitialCD4CountDate','CurrentCD4CountDate','LastEACDate','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','PatientOutcomeDate','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','DateOfBirth','MarkAsDeseasedDeathDate','BiometricCaptureDate','CurrentWeightDate','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimenDate', 'InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','PatientOutcomeDatePreviousQuarter','RecaptureDate']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartOfARTYears', 'MonthsOnART','DaysOfARVRefil','GestationAgeWeeks','CurrentViralLoad(c/ml)','CurrentAgeYears','CurrentWeight(Kg)','DrugDurationPreviousQuarter','QuantityOfARVDispensedLastVisit',]
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        #df = pd.read_csv(input_file_path, dtype=object, quoting=csv.QUOTE_NONE, on_bad_lines='skip')
    
        #Clean strings
        #dfstr = ['State','LGA','DatimCode','FacilityName','PatientUniqueID','PatientHospitalNo','ANCNoIdentifier','ANCNoConceptID','HTSNo','Sex','AgeAtStartOfARTYears','AgeAtStartOfARTMonths','CareEntryPoint','HIVConfirmedDate','KPType','MonthsOnART','DateTransferredIn','TransferInStatus','ARTStartDate','LastPickupDate','LastVisitDate','PillBalance','InitialRegimenLine','InitialRegimen','InitialCD4Count','InitialCD4CountDate','CurrentCD4Count','CurrentCD4CountDate','LastEACDate','CurrentRegimenLine','CurrentRegimen','PregnancyStatus','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','GestationAgeWeeks','CurrentViralLoad(c/ml)','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','ViralLoadIndication','PatientOutcome','PatientOutcomeDate','CurrentARTStatus','DispensingModality','FacilityDispensingModality','DDDDispensingModality','MMDType','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','CurrentAgeYears','CurrentAgeMonths','DateOfBirth','MarkAsDeseased','MarkAsDeseasedDeathDate','RegistrationPhoneNo','NextofKinPhoneNo','TreatmentSupporterPhoneNo','BiometricCaptured','BiometricCaptureDate','ValidCapture','CurrentWeight(Kg)','CurrentWeightDate','TBStatus','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcome','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimen','InitialFirstLineRegimenDate','InitialSecondLineRegimen','InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','DrugDurationPreviousQuarter','PatientOutcomePreviousQuarter','PatientOutcomeDatePreviousQuarter','ARTStatusPreviousQuarter','QuantityOfARVDispensedLastVisit','FrequencyOfARVDispensedLastVisit','CurrentARTStatusWithPillBalance','RecaptureDate','RecaptureCount']
        #for col in dfstr:
            #df[col] = df[col].str.replace('\"', '')
        
        #Clean Dates
        dfDates = ['HIVConfirmedDate','DateTransferredIn','ARTStartDate','LastPickupDate','LastVisitDate','InitialCD4CountDate','CurrentCD4CountDate','LastEACDate','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','PatientOutcomeDate','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','DateOfBirth','MarkAsDeseasedDeathDate','BiometricCaptureDate','CurrentWeightDate','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimenDate', 'InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','PatientOutcomeDatePreviousQuarter','RecaptureDate']
        for col in dfDates:
            df[col] = pd.to_datetime(df[col],errors='coerce').dt.date
         
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        
        df = df.drop(['LGA2','STATE2'], axis=1)
        
        
        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

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
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f" Converted File Location: {output_file_path}")
        

def CleanUbuntuARTLineList():

        start_time = time.time()  # Start time measurement
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
            df = pd.read_csv(input_file_path, dtype=object, quoting=csv.QUOTE_NONE, on_bad_lines='skip')
    
            #Clean strings
            dfstr = ['IP', 'State', 'LGA', 'Datim_Code', 'FacilityName', 'patient_id', 'PEPID', 'PatientHospitalNo', 'Sex', 'CareEntryPoint', 'KPType', 'AgeAtStartofART', 'AgeinMonths', 'DateConfirmedHIV+', 'ARTStartDate', 'DaysOnART', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DaysOfARVRefill', 'RegimenLineAtARTStart', 'RegimenAtARTStart', 'CurrentRegimenLine', 'CurrentARTRegimen', 'DSD_Model', 'DSD_Model_Type', 'CurrentPregnancyStatus', 'CurrentViralLoad', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'Alphanumeric_Viral_Load_Result', 'LastDateOfSampleCollection', 'ViralLoadIndication', 'Outcomes', 'Outcomes_Date', 'Cause_of_Death', 'VA_Cause_of_Death', 'IIT_Date', 'CurrentARTStatus_Pharmacy', 'CurrentARTStatus', 'ARTStatus_PreviousQuarter', 'DOB', 'Current_Age', 'CurrentAge_Months', 'TransferredIn', 'Date_Transfered_In', 'Surname', 'Firstname', 'Educationallevel', 'MaritalStatus', 'JobStatus', 'PhoneNo', 'Address', 'State_of_Residence', 'LGA_of_Residence', 'Weight', 'Height', 'BMI', 'BP', 'Whostage', 'DateofFirstTLD_Pickup', 'CurrentCD4', 'CurrentCD4Date', 'AHD_Indication', 'Current_CD4_LFA_Result', 'Other_Test_(TB-LAM_LF-LAM_etc)', 'Serology_for_CrAg_Result', 'CSF_for_CrAg_Result', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'Days_To_Schedule', 'IPT_Screening_Date', 'Are_you_coughing_currently', 'Do_you_have_fever', 'Are_you_losing_weight', 'Are_you_having_night_sweats', 'History_of_contacts_with_TB_patients', 'Sputum_AFB', 'Sputum_AFB_Result', 'GeneXpert', 'GeneXpert_Result', 'Chest_Xray', 'Chest_Xray_Result', 'Culture', 'Culture_Result', 'Is_Patient_Eligible_For_IPT', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'TPT_Outcomes', 'Date_of_TPT_Outcome', 'Current_TB_Status', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'Positive_living', 'Treatment_Literacy', 'Adolescents_participation', 'Leadership_training', 'Peer_To_Peer_Mentoship', 'Role_of_OTZ', 'OTZ_Champion_Oreintation', 'Transitioned_Adult_Clinic', 'OTZ_Outcome', 'OTZ_Outcome_Date', 'PBS_Capturee', 'PBS_Capture_Date', 'Date_Generated', 'Unique_Id', 'PBS_Recapture', 'PBS_Recapture_Date', 'PBS_Recapture_Count', 'uuid']
            for col in dfstr:
                df[col] = df[col].str.replace('\"', '')
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        #df = pd.read_csv(input_file_path, dtype=object, quoting=csv.QUOTE_NONE, on_bad_lines='skip')
    
        #Clean strings
        #dfstr = ['State','LGA','DatimCode','FacilityName','PatientUniqueID','PatientHospitalNo','ANCNoIdentifier','ANCNoConceptID','HTSNo','Sex','AgeAtStartOfARTYears','AgeAtStartOfARTMonths','CareEntryPoint','HIVConfirmedDate','KPType','MonthsOnART','DateTransferredIn','TransferInStatus','ARTStartDate','LastPickupDate','LastVisitDate','PillBalance','InitialRegimenLine','InitialRegimen','InitialCD4Count','InitialCD4CountDate','CurrentCD4Count','CurrentCD4CountDate','LastEACDate','CurrentRegimenLine','CurrentRegimen','PregnancyStatus','PregnancyStatusDate','EDD','LastDeliveryDate','LMP','GestationAgeWeeks','CurrentViralLoad(c/ml)','ViralLoadEncounterDate','ViralLoadSampleCollectionDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','ViralLoadIndication','PatientOutcome','PatientOutcomeDate','CurrentARTStatus','DispensingModality','FacilityDispensingModality','DDDDispensingModality','MMDType','DateReturnedToCare','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','CurrentAgeYears','CurrentAgeMonths','DateOfBirth','MarkAsDeseased','MarkAsDeseasedDeathDate','RegistrationPhoneNo','NextofKinPhoneNo','TreatmentSupporterPhoneNo','BiometricCaptured','BiometricCaptureDate','ValidCapture','CurrentWeight(Kg)','CurrentWeightDate','TBStatus','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcome','CurrentINHOutcomeDate','LastINHDispensedDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','EnrollmentDate','InitialFirstLineRegimen','InitialFirstLineRegimenDate','InitialSecondLineRegimen','InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','DrugDurationPreviousQuarter','PatientOutcomePreviousQuarter','PatientOutcomeDatePreviousQuarter','ARTStatusPreviousQuarter','QuantityOfARVDispensedLastVisit','FrequencyOfARVDispensedLastVisit','CurrentARTStatusWithPillBalance','RecaptureDate','RecaptureCount']
        #for col in dfstr:
            #df[col] = df[col].str.replace('\"', '')
        
        #Clean Dates
        dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
        for col in dfDates:
            df[col] = pd.to_datetime(df[col],errors='coerce').dt.date
         
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        
        df = df.drop(['LGA2','STATE2'], axis=1)
        
        
        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

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
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f" Converted File Location: {output_file_path}")



##


#Creating A tooltip Class
class ToolTip(object):

    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 260
        y = y + cy + self.widget.winfo_rooty() +27
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(tw, text=self.text, justify=LEFT,
                      background="#ffffe0", relief=SOLID, borderwidth=1,
                      font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def CreateToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)


def open_editor():
    editor_window = tk.Toplevel(root)
    tree = ttk.Treeview(editor_window)

    # Define columns
    tree['columns'] = list(dflamisnmrs.columns)

    # Format columns
    tree.column("#0", width=0, stretch=tk.NO)
    for col in tree['columns']:
        tree.column(col, anchor=tk.W, width=100)

    # Create headings
    tree.heading("#0", text='', anchor=tk.W)
    for col in tree['columns']:
        tree.heading(col, text=col, anchor=tk.W)

    # Insert data
    for index, row in dflamisnmrs.iterrows():
        tree.insert("", tk.END, values=list(row))

    # Pack treeview
    tree.pack(fill=tk.BOTH, expand=True)

    # Create edit button
    def edit_row():
        # Get selected row
        selected = tree.focus()
        values = tree.item(selected, 'values')

        # Create edit window
        edit_window = tk.Toplevel(editor_window)
        entries = []

        for col in tree['columns']:
            label = tk.Label(edit_window, text=col)
            label.grid(row=len(entries), column=0)
            entry = tk.Entry(edit_window)
            entry.grid(row=len(entries), column=1)
            entries.append(entry)

        # Fill entries with selected row values
        for i, value in enumerate(values):
            entries[i].insert(0, value)

        # Create save button
        def save_edits():
            # Get new values
            values = [entry.get() for entry in entries]

            # Update treeview
            tree.item(selected, values=values)

            # Close edit window
            edit_window.destroy()

        save_button = tk.Button(edit_window, text="Save", command=save_edits)
        save_button.grid(row=len(entries), column=1)

    edit_button = tk.Button(editor_window, text="Edit", command=edit_row)
    edit_button.pack()

    # Create save button
    def save_changes():
        # Get updated data from treeview
        data = []
        for child in tree.get_children():
            data.append(tree.item(child, 'values'))

        # Update DataFrame
        global dflamisnmrs
        dflamisnmrs = pd.DataFrame(data, columns=dflamisnmrs.columns)

        # Print updated DataFrame
        print(dflamisnmrs)

    save_button = tk.Button(editor_window, text="Save", command=save_changes)
    save_button.pack()

# Creating Main Window
root = tk.Tk()
root.title("NETO's NMRS TO RADET CONVERTER v3.1")
root.geometry("600x450")
root.config(bg="#f0f0f0")

frame3 = tk.Frame(root)
frame3.place(relx=1.0, rely=0, anchor='ne') # Position at the top right

#date text
selectinfo = tk.Label(root, text="Please remember to adjust start date and end date if necessary", font=("Helvetica", 10), fg="red")
selectinfo.pack(pady=1)

# Create a frame to hold the DateEntry widgets

frame = tk.Frame(root)
frame.pack(padx=10, pady=2)

frame2 = tk.Frame(root)
frame2.pack(padx=10, pady=0.5)

# Get the current date
current_date = datetime.now().date()

# Create a settings button in the top right corner
editbutton = tk.Button(frame3, text="edit fac", command=open_editor)
editbutton.pack(side=tk.RIGHT, padx=20, pady=10)

# Create DateEntry widgets with date format and select mode
start_date_label = tk.Label(frame2, text="Start Date", font=("Helvetica", 9))
start_date_label.pack(side=tk.LEFT, padx=20)
cal = DateEntry(frame, width=12, background="darkblue", foreground="white", borderwidth=2, date_pattern='dd-mm-yyyy', selectmode='day', year=2000, month=1, day=1)
cal.pack(side=tk.LEFT, padx=5)

end_date_label = tk.Label(frame2, text="End Date", font=("Helvetica", 9))
end_date_label.pack(side=tk.LEFT, padx=20)
cal2 = DateEntry(frame, width=12, background="darkblue", foreground="white", borderwidth=2, date_pattern='dd-mm-yyyy', selectmode='day')
cal2.set_date(current_date)
cal2.pack(side=tk.LEFT, padx=5)

# Create a StringVar to store the selected option
selected_option = tk.StringVar(root)
selected_option.set("NMRS-ART-LINE LIST TO RADET (ONLY CONVERTS TO RADET) ▼")  # Set an initial default value

# Create a frame to hold the label and OptionMenu
select_frame = tk.Frame(root)
select_frame.pack()

# Create a label for the selection text
select_label = tk.Label(select_frame, text="Select:", font=("Helvetica", 12))
select_label.pack(side="left")

# Create the OptionMenu with invisible dropdown arrow
option_menu = tk.OptionMenu(select_frame, selected_option, "NMRS-ART-LINE LIST TO RADET (ONLY CONVERTS TO RADET) ▼", "NMRS-ART-LINE LIST TO RADET (APPENDS BASELINE INFO (SELECT RADET BEFORE NMRS LINE LIST)) ▼", "CLEAN NMRS-ART-LINE LIST DATA TYPES (FROM .csv, .xls, .xlsx TO EXCEL(.xlsx)) ▼","CLEAN Ubuntu NMRS-ART-LINE LIST DATA TYPES (FROM .csv, .xls, .xlsx TO EXCEL(.xlsx)) ▼")
option_menu.config(indicatoron=0)  # Hide the arrow
option_menu.pack(side="left", padx=5, pady=10)

# Create a button to call the function based on the selected option
def on_dropdown_click():
    selected_value = selected_option.get()
    if selected_value == "NMRS-ART-LINE LIST TO RADET (ONLY CONVERTS TO RADET) ▼":
        lineListToRatet()
    elif selected_value == "NMRS-ART-LINE LIST TO RADET (APPENDS BASELINE INFO (SELECT RADET BEFORE NMRS LINE LIST)) ▼":
        lineListToRatet_Baseline()
    elif selected_value == "CLEAN NMRS-ART-LINE LIST DATA TYPES (FROM .csv, .xls, .xlsx TO EXCEL(.xlsx)) ▼":    
        CleanARTLineList()
    elif selected_value == "CLEAN Ubuntu NMRS-ART-LINE LIST DATA TYPES (FROM .csv, .xls, .xlsx TO EXCEL(.xlsx)) ▼":    
        CleanUbuntuARTLineList()

dropdown_button = tk.Button(root, text="PROCESS SELECTION", command=on_dropdown_click, font=("Helvetica", 14), bg="green", fg="#ffffff")
dropdown_button.place(x=200, y=50)
dropdown_button.pack(pady=0.5)

convert_button_exit = tk.Button(root, text="EXIT APPLICATION.......", command=root.destroy, font=("Helvetica", 14), bg="red", fg="#ffffff")
convert_button_exit.pack(pady=0.5)

# Progress bar widget
progress_bar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress_bar.pack(pady=20)

# Labels
status_label = tk.Label(root, text="0%", font=('Helvetica', 12))
status_label.pack(pady=0.5)

status_label2 = tk.Label(root, text='Welcome to NMRS ART line list to RADET converter', bg="#D3D3D3", font=('Helvetica', 9))
status_label2.pack(pady=20)

text3 = tk.Label(root, text="you will be prompted to select required files and the location you want to save the converted file \n(Please Remember to follow the file requirement procedures)")
text3.pack(pady=1)

text2 = tk.Label(root, text="Contacts: email: chinedum.pius@gmail.com, phone: +2348134453841")
text2.pack(pady=19)



# Adding File Dialog
filedialog = tk.filedialog 

# Running the GUI
root.mainloop()

#pyinstaller nmrs-radet.py --onefile --windowed --upx-dir "C:\upx-4.2.4-win64"
#python -m nuitka --windows-console-mode=disable --onefile --enable-plugin=tk-inter nmrs-radet.py