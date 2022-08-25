import os
import pandas as pd

import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

# Calculate the path to the root directory of this script
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '.'))

# Define path to directory for reformatted PEARS module exports
# Used output path from pears_nightly_export_reformatting.py
# Otherwise, custom field labels will cause errors
# pears_export_path = r"\path\to\reformatted_pears_data"
# Script demo uses /example_inputs directory
pears_export_path = ROOT_DIR + "/example_inputs"

Program_Activities_Export = pd.ExcelFile(pears_export_path + '/' + "Program_Activities_Export.xlsx")
PA_Data = pd.read_excel(Program_Activities_Export, 'Program Activity Data')
PA_Data = PA_Data.loc[PA_Data['program_areas'] == 'SNAP-Ed']

Indirect_Activities_Export = pd.ExcelFile(pears_export_path + '/' + "Indirect_Activity_Export.xlsx")
IA_Data = pd.read_excel(Indirect_Activities_Export, 'Indirect Activity Data')
IA_Data = IA_Data.loc[IA_Data['program_area'] == 'SNAP-Ed']
IA_IC = pd.read_excel(Indirect_Activities_Export, 'Intervention Channels')
IA_IC = IA_IC.loc[~IA_IC['activity'].str.contains('(?i)TEST', regex=True)]
IA_IC_Data = pd.merge(IA_Data, IA_IC, how='inner', on='activity_id')

sites = pd.read_excel(pears_export_path + '\\' + "Site_Export.xlsx", sheet_name='Site Data')
sites = sites.loc[sites['is_active'] == 1]

Partnerships_Export = pd.ExcelFile(pears_export_path + '/' + "Partnership_Export.xlsx")
Part_Data = pd.read_excel(Partnerships_Export, 'Partnership Data')
Part_Data = Part_Data.loc[Part_Data['program_area'] == 'SNAP-Ed']

Partnerships_2021 = pd.read_excel(
    r"C:\Users\jstadni2\Box\FCS Data Analyst\Data Backups\PEARS\FY21\12.21.21 Manual Export\Partnership-Export.xlsx",
    sheet_name='Partnership Data')

FY22_INEP_Staff = pd.read_excel(r"C:\Users\jstadni2\Box\INEP Staff Lists\FY22 INEP Staff List.xlsx",
                                sheet_name='SNAP-Ed Staff List', header=1)
User_Export = pd.read_excel(pears_export_path + '/' + "User_Export.xlsx", sheet_name='User Data')
Unit_Counties = pd.read_excel(
    r"C:\Users\jstadni2\Box\FCS Data Analyst\PEARS FY22\PEARS Data Cleaning\Illinois Extension Unit Counties.xlsx",
    sheet_name='PEARS Units')

# Partnerships Data Entry Report


import numpy as np

Part_Data = Part_Data.loc[
    ~Part_Data['partnership_name'].str.contains('TEST'), ['partnership_id', 'partnership_name', 'reported_by',
                                                          'reported_by_email', 'partnership_unit', 'site_id',
                                                          'site_name', 'site_address', 'site_city', 'site_state',
                                                          'site_zip', 'created', 'modified']]

exclude_sites = ['abc placeholder', 'U of I Extension', 'University of Illinois Extension']
PA_Data = PA_Data.loc[
    (~PA_Data['name'].str.contains('TEST')) & (~PA_Data['site_name'].str.contains('|'.join(exclude_sites)))]

PA_Data['id'] = 'pa' + PA_Data['program_id'].astype('str')
PA_Data = pd.merge(PA_Data, sites[['parent_site_name', 'site_id', 'site_name', 'address', 'city', 'state', 'zip_code']],
                   how='left', left_on='site_name', right_on='parent_site_name', suffixes=['', '_child'])

PA_Data.loc[PA_Data['site_id_child'].notnull(), 'site_id'] = PA_Data['site_id_child']
PA_Data.loc[PA_Data['site_id_child'].notnull(), ['partnership_name', 'site_name']] = PA_Data['site_name_child']
PA_Data.loc[PA_Data['site_id_child'].notnull(), 'site_address'] = PA_Data['address']
PA_Data.loc[PA_Data['site_id_child'].notnull(), 'site_city'] = PA_Data['city']
PA_Data.loc[PA_Data['site_id_child'].notnull(), 'site_zip'] = PA_Data['zip_code']
PA_Data.loc[PA_Data['site_id_child'].notnull(), 'site_state'] = PA_Data['state']

IA_IC_Data['id'] = 'ia' + IA_IC_Data['activity_id'].astype('str')
IA_IC_Data = IA_IC_Data.loc[(~IA_IC_Data['title'].str.contains('TEST')) & (
    ~IA_IC_Data['site_name'].str.contains('|'.join(exclude_sites), na=True))]

# Part_Entry = PA_Data[['id', 'program_areas', 'action_plans', 'comments', 'unit', 'site_id', 'site_name', 'site_address', 'site_city', 'site_state', 'site_zip', 'parent_site_name', 'site_id_child']].rename(columns={'program_areas' : 'program_area', 'action_plans' : 'action_plan_name'}).append(IA_IC_Data[['id', 'program_area', 'action_plan_name', 'unit', 'site_id', 'site_name', 'site_address', 'site_city', 'site_state', 'site_zip']]).drop_duplicates(subset='site_id', keep='first')
Part_Entry = PA_Data[
    ['id', 'program_areas', 'comments', 'unit', 'site_id', 'site_name', 'site_address', 'site_city', 'site_state',
     'site_zip', 'parent_site_name', 'site_id_child', 'snap_ed_grant_goals', 'snap_ed_special_projects',
     'reported_by_email']].rename(columns={'program_areas': 'program_area'}).append(IA_IC_Data[
                                                                                        ['id', 'program_area', 'unit',
                                                                                         'site_id', 'site_name',
                                                                                         'site_address', 'site_city',
                                                                                         'site_state', 'site_zip',
                                                                                         'reported_by_email']]).drop_duplicates(
    subset='site_id', keep='first')

Part_Entry = Part_Entry.loc[~Part_Entry['site_id'].isin(Part_Data['site_id'])].drop_duplicates(
    subset=['site_id']).rename(columns={'unit': 'partnership_unit', 'comments': 'program_activity_comments'})
Part_Entry['partnership_name'] = Part_Entry['site_name']
Part_Entry.insert(0, 'partnership_name', Part_Entry.pop('partnership_name'))
Part_Entry['action_plan_name'] = 'Health: Chronic Disease Prevention and Management (State - 2020-2021)'
Part_Entry.insert(3, 'action_plan_name', Part_Entry.pop('action_plan_name'))
# Part_Entry['assistance_received_recruitment'] = 1
# Part_Entry['assistance_received_space'] = 1
Part_Entry[
    'assistance_received'] = 'Recruitment (includes program outreach), Space (e.g., facility or room where programs take place)'
# Part_Entry['assistance_provided_human_resources'] = 1
# Part_Entry['assistance_provided_program_implementation'] = 1
Part_Entry[
    'assistance_provided'] = 'Human resources (*staff or staff time), Program implementation (e.g. food and beverage standards)'
Part_Entry['assistance_received_funding'] = 'No'
Part_Entry.loc[Part_Entry['id'].str.contains('pa'), 'is_direct_education_intervention'] = 1
Part_Entry.loc[Part_Entry['id'].str.contains('ia'), 'is_direct_education_intervention'] = 0

Part_Entry['collaborator_unit'] = Part_Entry['partnership_unit']
Part_Entry = pd.merge(Part_Entry, Unit_Counties, how='left', left_on='partnership_unit', right_on='County')
Part_Entry.loc[Part_Entry['partnership_unit'].isin(Unit_Counties['County']), 'collaborator_unit'] = Part_Entry['Unit']
Part_Entry = Part_Entry.drop(columns={'Unit', 'County'})
staff_nulls = ('N/A', 'NEW', 'OPEN', np.nan)
Collaborators = FY22_INEP_Staff.loc[FY22_INEP_Staff['JOB CLASS'].isin(['EPC', 'UE']), ['JOB CLASS', 'E-MAIL', 'COUNTY']]
Collaborators = pd.merge(Collaborators, User_Export[['full_name', 'email', 'unit', 'viewable_units']], how='inner',
                         left_on='E-MAIL', right_on='email').drop(
    columns={'JOB CLASS', 'E-MAIL', 'COUNTY', 'email'}).rename(columns={'full_name': 'collaborators'}).drop_duplicates()
Collaborators['viewable_units'] = Collaborators['viewable_units'].str.split(", ")
Collaborators.loc[Collaborators['viewable_units'].isnull(), 'viewable_units'] = ""
Collaborators.loc[Collaborators.viewable_units.map(len) > 1, 'unit'] = Collaborators['viewable_units']
Collaborators = Collaborators.explode('unit').drop(columns=['viewable_units'])
Part_Collaborators = pd.merge(Part_Entry[['partnership_name', 'collaborator_unit']], Collaborators, how='left',
                              left_on='collaborator_unit', right_on='unit')
Part_Collaborators = Part_Collaborators.groupby('partnership_name').agg(lambda x: x.dropna().unique().tolist())
Part_Collaborators = Part_Collaborators.drop(columns={'collaborator_unit', 'unit'})
Part_Entry = pd.merge(Part_Entry, Part_Collaborators, how='left', on='partnership_name').drop(
    columns=['collaborator_unit'])
Part_Entry['collaborators'] = [', '.join(map(str, l)) for l in Part_Entry['collaborators']]

Part_Entry['relationship_depth'] = 'Cooperator'
Part_Entry['assessment_tool'] = 'None'
Part_Entry['accomplishments'] = 'N/A'
Part_Entry['lessons_learned'] = 'N/A'

c_parts_site_id = pd.merge(Part_Entry, Partnerships_2021[
    ['partnership_id', 'partnership_name', 'site_id', 'site_name', 'site_zip']], how='left', on='site_id',
                           suffixes=('', '_copy')).rename(columns={'partnership_id': 'partnership_id_copy'})
c_parts_site_id = c_parts_site_id.loc[c_parts_site_id['partnership_id_copy'].notnull()]
cols1 = ['id', 'partnership_id_copy', 'partnership_name_copy', 'program_area', 'action_plan_name', 'site_id',
         'site_name_copy', 'site_address', 'site_city', 'site_state', 'site_zip', 'partnership_unit',
         'assistance_received', 'assistance_provided', 'assistance_received_funding',
         'is_direct_education_intervention', 'collaborators', 'snap_ed_grant_goals', 'snap_ed_special_projects',
         'relationship_depth', 'parent_site_name', 'program_activity_comments', 'reported_by_email']
c_parts_site_id_out = c_parts_site_id[cols1]

new_parts_cols = ['partnership_name', 'id', 'program_area', 'action_plan_name', 'program_activity_comments',
                  'partnership_unit', 'site_id', 'site_name', 'site_address', 'site_city', 'site_state', 'site_zip',
                  'parent_site_name', 'site_id_child', 'assistance_received', 'assistance_provided',
                  'assistance_received_funding', 'is_direct_education_intervention', 'collaborators',
                  'snap_ed_grant_goals', 'snap_ed_special_projects', 'relationship_depth', 'assessment_tool',
                  'accomplishments', 'lessons_learned', 'reported_by_email']
new_parts = Part_Entry.loc[~Part_Entry['site_id'].isin(c_parts_site_id['site_id']), new_parts_cols]
new_parts = new_parts.drop(columns='site_id_child')
new_parts.insert((len(new_parts.columns) - 1), 'program_activity_comments', new_parts.pop('program_activity_comments'))

ts = (pd.to_datetime("today") - pd.DateOffset(months=1)).to_pydatetime()
prev_month = pd.to_datetime(ts).to_period('M')

out_path = ROOT_DIR + "/example_outputs"

## SNAP-Ed Workbook


SNAPED_c_parts_site_id = c_parts_site_id_out.loc[c_parts_site_id['partnership_unit'] != 'CPHP (District)']
SNAPED_new_parts = new_parts.loc[new_parts['partnership_unit'] != 'CPHP (District)']

dfs1 = {'New Partnerships': SNAPED_new_parts, 'Copy Forward - Site ID Matches': SNAPED_c_parts_site_id}

filename1 = 'SNAP-Ed Partnerships Data Entry ' + prev_month.strftime('%Y-%m') + '.xlsx'

file_path1 = out_path + '/' + filename1


def write_report(file_path, dfs_dict):
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    for sheetname, df in dfs_dict.items():  # loop through `dict` of dataframes
        df.to_excel(writer, sheet_name=sheetname, index=False, freeze_panes=(1, 0))  # send df to writer
        worksheet = writer.sheets[sheetname]  # pull worksheet object
        worksheet.autofilter(0, 0, 0, len(df.columns) - 1)
        for idx, col in enumerate(df):  # loop through all columns
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)  # set column width
    writer.save()


write_report(file_path1, dfs1)

## CPHP Workbook


CPHP_c_parts_site_id = c_parts_site_id_out.loc[c_parts_site_id['partnership_unit'] == 'CPHP (District)']
CPHP_new_parts = new_parts.loc[new_parts['partnership_unit'] == 'CPHP (District)']

dfs2 = {'New Partnerships': CPHP_new_parts, 'Copy Forward - Site ID Matches': CPHP_c_parts_site_id}

filename2 = 'CPHP Partnerships Data Entry ' + prev_month.strftime('%Y-%m') + '.xlsx'

file_path2 = out_path + '\\' + filename2

write_report(file_path2, dfs2)


# Email Data Entry Report


# Set the following variables with the appropriate credentials and recipients
admin_username = 'your_username@domain.com'
admin_password = 'your_password'
admin_send_from = 'your_username@domain.com'
report_cc = 'list@domain.com, of_recipients@domain.com'


# Send an email with or without a xlsx attachment
# send_from: string for the sender's email address
# send_to: string for the recipient's email address
# Cc: string of comma-separated cc addresses
# subject: string for the email subject line
# html: string for the email body
# username: string for the username to authenticate with
# password: string for the password to authenticate with
# isTls: boolean, True to put the SMTP connection in Transport Layer Security mode (default: True)
# wb: boolean, whether an Excel file should be attached to this email (default: False)
# file_path: string for the xlsx attachment's filepath (default: '')
# filename: string for the xlsx attachments filename (default: '')
def send_mail(send_from,
              send_to,
              cc,
              subject,
              html,
              username,
              password,
              is_tls=True,
              wb=False,
              file_path='',
              filename=''):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Cc'] = cc
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(html, 'html'))

    if wb:
        fp = open(file_path, 'rb')
        part = MIMEBase('application', 'vnd.ms-excel')
        part.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(part)

    smtp = smtplib.SMTP('smtp.office365.com', 587)
    if is_tls:
        smtp.starttls()
    try:
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to.split(',') + msg['Cc'].split(','), msg.as_string())
    except smtplib.SMTPAuthenticationError:
        print("Authentication failed. Make sure to provide a valid username and password.")
    smtp.quit()


report_recipients = 'recipient@domain.com'
report_subject = 'PEARS Sites Report ' + prev_month.strftime('%Y-%m')

html1 = """<html>
  <head></head>
<body>
            <p>
           Hello DATA ENTRY SUPPORT,<br><br>

            The attached data is for Direct/Indirect Education partners that require Partnership Module entries. Could you please enter them into PEARS? Should you need it, the Partnerships Cheat Sheet is located <a href="https://uofi.app.box.com/folder/49632670918?s=wwymjgjd48tyl0ow20vluj196ztbizlw">here</a>.          

            <ul>
              <li>New Partnerships for Direct Education contain 'pa' in the id field and whereas the id for Indirect Education Partnerships contain 'ia'.</li> 
              <li>If the Partnership Unit is set to 'Illinois - University of Illinois Extension (Implementing Agency)', please select a more appropriate unit.</li>
              <li>When copying forward Partnerships from a previous year, make sure the new entry matches the data in this spreadsheet.</li>
              <li>Copied Partnerships should only display '(Copied)' in the title once.</li>
              <li>District-level Direct Education requires an individual Partnership for each Site in attendance.</li>
              <li>If the Parent Site column is not empty, please verify that all sites listed in the Program Activity Comments have corresponding Site and Partnership entries.</li>
              <li>If the SNAP-Ed Grant Goals or SNAP-Ed Special Projects fields are empty, contact staff who created the original record (in the ID field) for the appropriate values.</li>                 			  
            </ul>

          If you have any questions, please reply to this email and I will respond at my earliest opportunity.<br>

            <br>Thanks and have a great day!<br>       
            <br> <b> FCS Evaluation Team </b> <br>
            <a href = "mailto: your_username@domain.com ">your_username@domain.com </a><br>
            </p>
  </body>
</html>
"""

send_mail(send_from, send_to1, Cc, subject1, html1, file_path1, filename1, username, password, isTls=True)

subject2 = 'CPHP Partnerships Data Entry ' + prev_month.strftime('%Y-%m')

if any(x.empty is False for x in dfs2.values()):
    send_mail(send_from, send_to1, Cc, subject2, html1, file_path2, filename2, username, password, isTls=True)
