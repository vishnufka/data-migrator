import datetime as dt

base_path = "C:\\Users\\username\\companyname\\Documents\\Clients\\clientname\\3_final_data\\"
data_path = base_path + "3_final_data_master\\"
output_path = base_path + "3_final_data_output\\"

header_row_index = 0

data_mapping_dict = {
    "202205. Company.xlsx": "company",
    "202205. Investors.xlsx": "investorcompany",
    "202205. Investor Contacts.xlsx": "investorcontact",
    "202205. Contacts.xlsx": "contact",
    "202205. Deals.xlsx": "deal",
    "202205. Link Board member to Portfolio Company.xlsx": "boardmember",
    "202205. Link Deal team.xlsx": "dealteam",
    "dealfieldsfromcompany": "dealfieldsfromcompany",
    "202205. Financials.xlsx": "financial",
    "202205. Activity Investor.xlsx": "activity",
    "202205. Notes.xlsx": "note"
}

files_list = {
    #"202205. Company.xlsx",
    #"202205. Investors.xlsx.xlsx",
    #"202205. Financials.xlsx",
    #"202205. Activity Investor.xlsx",
    #"202205. Notes.xlsx"
}

user_mapping = {
    "A.User1": "A User1 Fullname",
    "A.User2": "A User2 Fullname",
    "A.User3": "A User3 Fullname",
    "A.User4": "A User4 Fullname",
    "A.User5": "A User5 Fullname",
    "A.User6": "A User6 Fullname",
    "A.User7": "A User7 Fullname",
    "A.User8": "A User8 Fullname",
    "A.User9": "A User9 Fullname",
    "": "",
    None: ""
}

mapping_dict = {
    "company": {
        "IQId": "Legacy Company ID",
        "Created on": None,
        "Company": "Company",
        "Short code": "Short code",
        "Status": "(REMAP)Status",
        "Legal form": None,
        "Active": None,
        "Creation date": None,
        "Web": "Website",
        "Currency": None,
        "Original country": "(REMAP)Country",
        "Transaction Process": None,
        "Legal Name": "Legal Name",
        "Fund": None,
        "Fund's role in initial investment": None,
        "clientname ownership %": None,
        "Priority": None,
        "Notes": "(REMAP)Notes",
        "Industry1": "Business Description",
        "International Transaction?": None,
        "EV": None,
        "Zipcode": "Postcode",
        "Country": None,
        "Phone": "Company Main Phone",
        "Fax": None,
        "Email": None,
        "Company info": None
    },
    "contact": {
        "IQId": "Legacy Contact ID",
        "Created on": "Legacy Created On",
        "Contact type": "Contact Type",
        "Last name": "Last Name",
        "Middle name": "(REMAP)Middle Name",
        "First name": "First Name",
        "Contact Full Name": None,
        "Suffix": "Suffix",
        "Salutation": "Salutation",
        "Contact info": None,
        "Mobile": "Mobile Phone",
        "Office": "Office Phone",
        "Home": None,
        "Primary Citizenship": None,
        "Assistant's name": "Assistant Name",
        "Assistant's phone": "Assistant Phone",
        "Description": "Description"
    },
    "deal": {
        "IQId": "Legacy Deal ID",
        "Created on": "New Deal Date",
        "Number": None,
        "Deal name": "Deal Name",
        "Deal Full name": None,
        "Deal type": None,
        "My Deal": None,
        "Target": "Company Name",
        "Assigned to": None,
        "Stage": "(REMAP)Stage",
        "Currency": None,
        "Web": None,
        "Country": "(REMAP)Deal Country",
        "Step": None,
        "Changing step date": "(REMAP)Changing Step Date",
        "Date received": None,
        "Project type": None,
        "Decline Reason": "Passed Commentary",
        "clientname Project Number": "Legacy Project Number",
        "Proprietary": "Proprietary",
        "Investment Type": "Transaction Type",
        "Transaction Process": "Process Type"
    },
    "boardmember": {
        "Company Affiliation ID": "Legacy Company Affiliation ID",
        "Company ID": "Company ID",
        "Company": None,
        "Role": "Role",
        "Category": "Category",
        "Last name": None,
        "First name": None
    },
    "dealteam": {
        "Legacy Deal Team ID": "Legacy Deal Team ID",
        "IQId": "Company ID",
        "Start date": "Start Date",
        "End date": "End Date"
    },
    "dealfieldsfromcompany": {
        "Legacy Company ID": "Legacy Company ID",
        "(REMAP)Notes": "(REMAP)Notes",
    },
    "investorcompany": {
        "IQId": "Legacy Company ID",
        "Investor": "Company Name",
        "Short Code": "Short Code",
        "Investor Type": "(REMAP)Investor Type",
        "Status": "Status",
        "Tier": "(REMAP)Tier",
        "Currency": None,
        "Web": "Website",
        "Geographical Region": None,
        "Site Name": None,
        "Address": "(REMAP)Address",
        "Address2": "(REMAP)Address2",
        "City": "City",
        "State": "State",
        "Zipcode": "Postcode",
        "Country": "(REMAP)Country",
        "Switchboard": "Company Main Phone",
    },
    "investorcontact": {
        "IQId": "Legacy Contact ID",
        "Created on": "Legacy Created On",
        "Contact type": "(REMAP)Contact Type",
        "Last name": "Last Name",
        "Middle name": "(REMAP)Middle Name",
        "First name": "First Name",
        "Contact Full Name": None,
        "Suffix": "Legacy Suffix",
        "Salutation": "Salutation",
        "Language": "Language",
        "Investor": "Company Name",
        "Investor ID": "Company ID",
        "Title / Detailed Salutation": "Legacy Detailed Salutation",
        "Position": None,
        "Detailed Position": None,
        "Department": "Department",
        "Email": "Email",
        "Contact info": None,
        "Mobile": "Mobile Phone",
        "Office": "Office Phone",
    },
    "financial": {
        "Legacy Financial ID": "Legacy Financial ID",
        "Company ID": "Company ID",
        "Company": None,
        "Date": "Legacy Created Date",
        "EBIT": "(REMAP)EBIT",
        "Total revenues": "(REMAP)Total Revenue",
    },
    "activity": {
        "Legacy Activity ID": "Legacy Activity ID",
        "Owner": "Legacy Owner",
        "Subject": "Subject",
        "Start": "Start",
        "End": "End",
        "Due Date": None,
        "Location": "Location",
        "Activity": "(REMAP)Type",
    },
    "note": {
        "Legacy Note ID": "Legacy Note ID",
        "Date": "Date",
        "Type": "(REMAP)Type",
        "Language": None,
        "Subject": "Subject",
        "Description": "Notes",
        "Status": None,
        "Include Attachment": None,
        "Created by": None,
    }
}

linking_excels = {
    "link_contact_to_company": data_path + "linking_files\\202205. Link Contact to Company.xlsx",
    "link_deal_to_company": data_path + "linking_files\\202205. Link Deal to Company.xlsx",
    "link_company_to_deal": data_path + "linking_files\\202205. Link Deal to Company.xlsx",
    "link_deal_to_advisor_company": data_path + "linking_files\\202205. Link Advisor to Deal.xlsx",
    "link_deal_to_advisor_contact": data_path + "linking_files\\202205. Link Advisor to Deal.xlsx",
    "link_deal_to_addon_deal": data_path + "linking_files\\202205. Link Add-on to Parent.xlsx",
    "link_deal_to_addon_deal_prospect_date": data_path + "linking_files\\202205. Link Add-on to Portfolio Company.xlsx",
    "link_deal_to_addon_deal_invested_date": data_path + "linking_files\\202205. Link Add-on to Portfolio Company.xlsx",
    "link_deal_to_addon_deal_rejected_date": data_path + "linking_files\\202205. Link Add-on to Portfolio Company.xlsx",
}

# both columns are from the linking_excel defined above
# left side is the index, this will be searched for in the main excel
# right side is the value, this will be added to the main excel
linking_pairs = {
    "link_contact_to_company": ["Contact ID", "Company ID"],
    "link_deal_to_company": ["Deal ID", "Company ID"],
    "link_company_to_deal": ["Company ID", "Deal ID"],
    "link_deal_to_advisor_company": ["Deal ID", "Advisor Company ID"],
    "link_deal_to_advisor_contact": ["Deal ID", "Advisor Contact ID"],
    "link_deal_to_addon_deal": ["Add-On Company ID", "Parent Company ID"],
    "link_deal_to_addon_deal_prospect_date": ["Add-On Company ID", "Date Received as Prospect"],
    "link_deal_to_addon_deal_invested_date": ["Add-On Company ID", "Date Invested"],
}

# the column to join on in the main excel, if different to index from linking_pair
linking_matching_column = {
    "link_deal_to_addon_deal": "Company ID",
    "link_company_to_deal": "Company ID",
    "link_deal_to_addon_deal_prospect_date": "Company ID",
    "link_deal_to_addon_deal_invested_date": "Company ID",
}

