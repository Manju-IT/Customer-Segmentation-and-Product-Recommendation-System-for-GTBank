import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Load the dataset
df = pd.read_excel("Customer Segmentation and Product Recommendation System for GTBank.xlsx", sheet_name="Sheet1")

# Data Cleaning
# Step 1: Handle missing values and standardize "None" entries
df.replace("None", np.nan, inplace=True)
df['Loan Status'].fillna("No Loan", inplace=True)

# Step 2: Check data types and convert if needed
df = df.convert_dtypes()

# Step 3: Standardize categorical values
df['Location (City/State)'] = df['Location (City/State)'].str.title()
df['Preferred Channels'] = df['Preferred Channels'].str.title()

# Step 4: Handle outliers
numeric_cols = ['Age', 'Average Monthly Balance', 'Transaction Frequency (per month)', 'Service Feedback Score']
for col in numeric_cols:
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    df = df[~((df[col] < (Q1 - 1.5 * IQR)) | (df[col] > (Q3 + 1.5 * IQR)))]

# Exploratory Data Analysis (EDA)
plt.figure(figsize=(15, 20))

# Demographic Analysis
plt.subplot(4, 2, 1)
sns.histplot(df['Age'], bins=20, kde=True)
plt.title('Age Distribution')

plt.subplot(4, 2, 2)
df['Gender'].value_counts().plot.pie(autopct='%1.1f%%')
plt.title('Gender Distribution')

plt.subplot(4, 2, 3)
df['Location (City/State)'].value_counts().head(10).plot.bar()
plt.title('Top 10 Locations')

# Account Analysis
plt.subplot(4, 2, 4)
df['Account Type'].value_counts().plot.bar()
plt.title('Account Type Distribution')

# Transaction Behavior
plt.subplot(4, 2, 5)
sns.boxplot(x=df['Average Monthly Balance'])
plt.title('Monthly Balance Distribution')

plt.subplot(4, 2, 6)
sns.boxplot(x=df['Transaction Frequency (per month)'])
plt.title('Transaction Frequency Distribution')

# Service Analysis
plt.subplot(4, 2, 7)
df['Loan Status'].value_counts().plot.bar()
plt.title('Loan Status Distribution')

plt.subplot(4, 2, 8)
df['Preferred Channels'].value_counts().plot.bar()
plt.title('Preferred Channels Distribution')

plt.tight_layout()
plt.savefig('eda_visualizations.png', dpi=300)
plt.show()

# Key Findings Summary
print("\nKey Findings:")
print(f"1. Customer Age Range: {df['Age'].min()} - {df['Age'].max()} years")
print(f"2. Most Common Account Type: {df['Account Type'].mode()[0]}")
print(f"3. Most Active Transaction Channel: {df['Preferred Channels'].mode()[0]}")
print(f"4. Average Monthly Balance Range: ${df['Average Monthly Balance'].min():,.0f} - ${df['Average Monthly Balance'].max():,.0f}")
print(f"5. Most Frequent Loan Status: {df['Loan Status'].mode()[0]}")