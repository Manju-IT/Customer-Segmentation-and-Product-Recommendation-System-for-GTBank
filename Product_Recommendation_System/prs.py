import pandas as pd

def segment_customer(row):
    """Segment customers based on their account type, balance, transaction frequency, and loan status."""
    if row["Account Type"] == "Business":
        return "Business Owner"
    elif row["Average Monthly Balance"] > 500000 and row["Transaction Frequency (per month)"] > 15:
        return "High-Value Customer"
    elif row["Average Monthly Balance"] < 100000 and row["Transaction Frequency (per month)"] < 10:
        return "Low-Value Customer"
    elif row["Loan Status"] == "Active":
        return "Loan Seeker"
    else:
        return "Regular Customer"

def recommend_product(segment):
    """Provide tailored product recommendations based on customer segments."""
    recommendations = {
        "Business Owner": ["Business Loan", "High-Interest Savings Account", "Corporate Credit Card"],
        "High-Value Customer": ["Premium Savings Account", "Investment Plans", "Priority Banking Services"],
        "Low-Value Customer": ["Basic Savings Account", "Cashback Debit Card", "Financial Literacy Programs"],
        "Loan Seeker": ["Personal Loan Offers", "Debt Consolidation Plans", "Credit Score Improvement Services"],
        "Regular Customer": ["Mobile Banking Tools", "Standard Savings Account", "Personalized Offers"]
    }
    return ", ".join(recommendations.get(segment, ["No recommendation available"]))

def load_data(file_path):
    """Load customer data from an Excel file."""
    return pd.read_excel(file_path)

def process_data(df):
    """Apply customer segmentation and product recommendation logic."""
    df["Customer Segment"] = df.apply(segment_customer, axis=1)
    df["Product Recommendations"] = df["Customer Segment"].apply(recommend_product)
    return df

def save_data(df, output_path):
    """Save processed customer recommendations to an Excel file."""
    df.to_excel(output_path, index=False)
    print(f"Product recommendations saved successfully to {output_path}")

def main():
    """Main function to execute the product recommendation workflow."""
    input_file = "Customer Segmentation and Product Recommendation System for GTBank.xlsx"
    output_file = "Customer_Recommendations.xlsx"
    
    print("Loading customer data...")
    df = load_data(input_file)
    
    print("Processing data for segmentation and recommendations...")
    df = process_data(df)
    
    print("Saving recommendations...")
    save_data(df, output_file)
    
    print("Process completed successfully!")

if __name__ == "__main__":
    main()
