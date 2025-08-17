

import pandas as pd
from ast import literal_eval
from pathlib import Path



# path to raw csv data
df = pd.read_csv(r"C:\Users\SMacArthur\OneDrive\Data Projects\DF Commercial.csv")



#  ******* Cleaning Steps *******

# columns to keep
keep_cols = [
    'ID',
    'Created',
    'Interaction Type',
    'Channel',
    'What device did the customer order?',
    'Were there any vulnerability markers on the customers account?',
    '1 - DP Compliance - Were the relevant data protection verification checks carried out?',
    '2 - FS Compliance - Was eligibility taken explained fuilly?',
    '5 - FS Compliance - Did agent capture all new customers data/ check existing data?',
    '8 - FS Compliance - Up front cost details explained before credit check?',
    '22 - Ofcom  Compliance - Customer consent to credit check obtained?',
    '23 - FS Compliance - E-Sign email sent & confirmation obtained?',
    '1st Line QA Outcome',
    'Is remediation Required?'
]

# normalises any non breaking spaces in headers in the raw csv
keep_cols = [c.replace('\u00A0', ' ') for c in keep_cols]
df = df[keep_cols].copy()

# rename the columns being kept for readability 
df.rename(columns={
    'What device did the customer order?': 'Device Ordered',
    'Were there any vulnerability markers on the customers account?': 'Vulnerable Customer?',
    '1 - DP Compliance - Were the relevant data protection verification checks carried out?': 'Data Protection',
    '2 - FS Compliance - Was eligibility taken explained fuilly?': 'Eligibility Explained?',
    '5 - FS Compliance - Did agent capture all new customers data/ check existing data?': 'Customer Data Captured Correctly?',
    '8 - FS Compliance - Up front cost details explained before credit check?': 'Credit Check Cost Explained?',
    '22 - Ofcom  Compliance - Customer consent to credit check obtained?': 'Credit Check Consent Gained?',
    '23 - FS Compliance - E-Sign email sent & confirmation obtained?': 'Customer Signature Obtained?',
    '1st Line QA Outcome': 'Interaction Outcome'
}, inplace=True)


# filtered data in df to 2025 only
df['Created'] = pd.to_datetime(df['Created'], dayfirst=True)
df = df[df['Created'].dt.year == 2025]


# cleaned na values in below columns with replacements to standaridise column contents
df['Vulnerable Customer?'] = df['Vulnerable Customer?'].fillna('No')
df['Data Protection'] = df['Data Protection'].fillna('Pass')
df['Eligibility Explained?'] = df['Eligibility Explained?'].fillna('Pass')
df['Customer Data Captured Correctly?'] = df['Customer Data Captured Correctly?'].fillna('Pass')
df['Credit Check Cost Explained?'] = df['Credit Check Cost Explained?'].fillna('Pass')
df['Credit Check Consent Gained?'] = df['Credit Check Consent Gained?'].fillna('Pass')
df['Customer Signature Obtained?'] = df['Customer Signature Obtained?'].fillna('Pass')


# cleaned interaction outcome column and standardised values
df['Interaction Outcome'] = df['Interaction Outcome'].str.strip()
df['Interaction Outcome']=df['Interaction Outcome'].replace({
    'Good outcome with Failures': 'Good Outcome with Failures',
    'Good outcome with failures': 'Good Outcome with Failures',
    'Good Outcome with  Failures': 'Good Outcome with Failures',
    'Good Outcome with failures': 'Good Outcome with Failures',
})


# cleaned remediation required column
# replaced some values but have removed this step for confidentiality due to employee names being visbile
df['Is remediation Required?']=df['Is remediation Required?'].str.strip()


# Replaced values in Channel to improve readability 
df['Channel'] = df['Channel'].replace('Retentions', 'Existing')
df['Channel'] = df['Channel'].replace('Acquisitions', 'New')


#  ******* Cleaning Steps *******



#  ******* KPIs *******

# total audits (2025)
total_audits = len(df)

# vulnerable customers
vuln_count = (df['Vulnerable Customer?'] == 'Yes').sum()
vuln_percent = (vuln_count / total_audits) * 100
vuln_percent_cleaned = round(vuln_percent,2)



# devices ordered
df['Device Ordered'] = df['Device Ordered'].apply(literal_eval)
total_devices_ordered = df['Device Ordered'].explode().value_counts().sum()
invidual_devices_ordered = df['Device Ordered'].explode().value_counts()
individual_device_percentage = (invidual_devices_ordered / total_devices_ordered) * 100
individual_device_percentage_cleaned = round(individual_device_percentage,2)



# Customer Types
new_customers = (df['Channel'] == 'New').sum()
existing_customers = (df['Channel'] == 'Existing').sum()
new_customer_percent = (new_customers / total_audits) * 100
new_customer_percent_cleaned = round(new_customer_percent, 2)
existing_customers_percent = (existing_customers / total_audits) * 100
existing_customers_percent_cleaned = round(existing_customers_percent, 2)



# Interaction Types
interaction_types = df['Interaction Type'].value_counts()
interaction_type_percentage = (interaction_types / total_audits) * 100
interaction_type_percentage_cleaned = round(interaction_type_percentage, 2)



# Outcomes
outcomes = df['Interaction Outcome'].value_counts()
outcomes_percentage = (outcomes / total_audits) * 100
outcomes_percentage_cleaned = round(outcomes_percentage, 2)


# Remediation
total_remediation = (df['Is remediation Required?'] == 'Yes').sum()
total_remediation_percentage = (total_remediation / total_audits) * 100
total_remediation_percentage_cleaned = round(total_remediation_percentage, 2)



# list for compliance question columns
compliance_cols = [
    'Data Protection', 'Eligibility Explained?',
    'Customer Data Captured Correctly?', 'Credit Check Cost Explained?',
    'Credit Check Consent Gained?', 'Customer Signature Obtained?'
]

# melted compliance questions to get question and pass/fail per question
# added Pass Rate % column for each compliance question column
df_unpivot = pd.melt(df,value_vars=compliance_cols, var_name='Question', value_name='Result')
compliance_question_results = df_unpivot.groupby(['Question', 'Result']).size()
compliance_results_unstacked = compliance_question_results.unstack(level=-1)
compliance_results_unstacked['Pass Rate %'] = compliance_results_unstacked['Pass'] / (compliance_results_unstacked['Fail'] + compliance_results_unstacked['Pass']) * 100
compliance_results_unstacked['Pass Rate %'] = round(compliance_results_unstacked['Pass Rate %'], 2)



# list to store single value KPIs
data = [
    ("Total Audits", total_audits),
    ("Vulnerable Customer %", vuln_percent_cleaned),
    ("New Customer %", new_customer_percent_cleaned),
    ("Existing Customer %", existing_customers_percent_cleaned),
    ("Remediation %", total_remediation_percentage_cleaned)

]

# new df created with the single KPIs above
kpi_df = pd.DataFrame(data, columns=['Metric', 'Value'])


# file path to store created excel file 
file_path = Path(r"C:\Users\SMacArthur\OneDrive\Data Projects")

# path and file name for created excel file
output_file = file_path / 'Commercial Data 2025.xlsx'

# excel created with tabs for: cleaned dataset, grouped KPI data and each individual KPI
with pd.ExcelWriter(output_file) as writer:
    df.to_excel(writer, sheet_name='Commercial Data Cleaned', index=False)
    kpi_df.to_excel(writer, sheet_name='KPI Results', index=False)
    compliance_results_unstacked.to_excel(writer, sheet_name='Compliance_Questions', index=True)
    outcomes_percentage_cleaned.to_excel(writer, sheet_name='Outcome_Percentages', index=True)
    interaction_type_percentage_cleaned.to_excel(writer, sheet_name='Interaction_Types', index=True)
    individual_device_percentage_cleaned.to_excel(writer, sheet_name='Devices_Percentages', index=True)




#  ******* KPIs *******