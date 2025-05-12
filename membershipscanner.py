# -*- coding: utf-8 -*-
"""
Created on Mon Apr 28 15:03:45 2025

for dts


@author: izzyg
"""
import pandas as pd
from datetime import datetime, timedelta

# Load the two Excel files
previous_file = "previous_membership.xlsx"
current_file = "current_membership.xlsx"

# Read the Excel files
previous_membership = pd.read_excel(previous_file, engine='openpyxl')
current_membership = pd.read_excel(current_file, engine='openpyxl')

# Convert "End Date" columns to datetime
previous_membership['End Date'] = pd.to_datetime(previous_membership['End Date'])
current_membership['End Date'] = pd.to_datetime(current_membership['End Date'])

# Standardize "Full Name" for consistent comparison
previous_membership['Full Name'] = previous_membership['Full Name'].str.strip().str.lower()
current_membership['Full Name'] = current_membership['Full Name'].str.strip().str.lower()

# Find expired memberships in the previous membership file (End Date before today)
today = datetime.today()
expired_members = previous_membership[previous_membership['End Date'] < today]

# Filter expired memberships that are NOT present in the current membership file
unique_expired = expired_members[~expired_members['Full Name'].isin(current_membership['Full Name'])]

# Find memberships expiring soon in the current membership file (End Date within the next 30 days)
expiry_threshold = today + timedelta(days=30)
expiring_soon = current_membership[(current_membership['End Date'] > today) & (current_membership['End Date'] <= expiry_threshold)]

# Output results
print("Expired Memberships (Unique to Previous File):\n", unique_expired[['Full Name', 'End Date']])
print("\nMembers Expiring Soon:\n", expiring_soon[['Full Name', 'End Date']])

# Save results to separate Excel files
unique_expired.to_excel("unique_expired_members.xlsx", index=False, engine='openpyxl')  # Retain original format
expiring_soon.to_excel("expiring_soon.xlsx", index=False, engine='openpyxl')  # Retain original format