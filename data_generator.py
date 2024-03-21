import pandas as pd
from datetime import datetime
import random
import numpy as np

# Funtion to generate a random date within a given range
def random_date(start, end):
    """Generate a random date between start and end."""
    return start + (end - start) * random.random()

# Data generation process
def adjust_data(rows, start_date, end_date):
    """Generate rows of data according to the adjusted specifications."""
    data = []
    for _ in range(rows):
        product_name = random.choice(["Tennis Racket", "Tennis Balls", "Tennis Shoes", "Tennis Bag", "Tennis Apparel"])
        
        if product_name == "Tennis Racket":
            sale_price = round(random.uniform(100, 200), 2)
            quantity_sold = random.randint(1, 2)
            cost_price = round(sale_price * random.uniform(0.6, 0.8), 2)  # Cost price is 60-80% of sale price
        elif product_name == "Tennis Shoes":
            sale_price = round(random.uniform(80, 150), 2)
            quantity_sold = random.randint(1, 3)
            cost_price = round(sale_price * random.uniform(0.6, 0.8), 2)
        else:
            sale_price = round(random.uniform(20, 70), 2)
            quantity_sold = random.randint(2, 5)
            cost_price = round(sale_price * random.uniform(0.7, 0.9), 2)  # Assuming a smaller margin for accessories
        
        profit = round((sale_price - cost_price) * quantity_sold, 2)  # Profit calculation
        retail_price = round(sale_price + random.uniform(5, 20), 2)
        discount_amount = round(retail_price - sale_price, 2)
        sales_channel = random.choice(["Online", "In-store"])  # Define sales_channel
        return_status = "Yes" if random.uniform(0, 100) < 2 else "No"  # Define return_status
        shipping_cost = round(random.uniform(0, 15), 2) if sales_channel == "Online" else 0
        shipping_time = random.choice(["1 day", "2-3 days", "1 week"]) if sales_channel == "Online" else "N/A"
        reason_for_return = random.choice(["Defective", "Unwanted", "Wrong Size"]) if return_status == "Yes" else "N/A"
        customer_feedback = "Unsatisfied" if return_status == "Yes" else random.choice(["Satisfied", "Very Satisfied"])
        
        data.append([
            random.randint(1000, 9999),  # Product ID
            product_name,  # Product Name
            "Rackets" if product_name == "Tennis Racket" else "Accessories",  # Category
            random.choice(["Wilson", "Babolat", "Adidas", "Nike", "Head"]),  # Brand
            f"Model {random.randint(1, 20)}",  # Model
            f"SKU{random.randint(10000, 99999)}",  # SKU
            quantity_sold,  # Quantity Sold
            sale_price,  # Sale Price
            retail_price,  # Retail Price
            discount_amount,  # Discount Amount
            random_date(start_date, end_date).strftime("%Y-%m-%d"),  # Sale Date
            random.randint(10000, 99999),  # Customer ID
            sales_channel,  # Sales Channel
            random.choice(["Credit Card", "Cash", "Online Payment"]),  # Payment Method
            shipping_cost,  # Shipping Cost
            shipping_time,  # Shipping Time
            return_status,  # Return Status
            reason_for_return,  # Reason for Return
            customer_feedback,  # Customer Feedback
            profit, # Profit 
            random.choice(["North America", "Europe", "Asia"]),  # Region
            random.choice(["New York", "Paris", "Tokyo", "Online"]) if sales_channel == "In-store" else "Online"  # Store Location
        ])
    return data

# Seed for reproducibility
np.random.seed(42)
random.seed(42)

# Start & End Date
start_date = datetime(2022, 1, 1)
end_date = datetime(2023, 12, 31)

# Number of Rows
rows = 20000

# Generating Data
adjusted_data = adjust_data(rows, start_date, end_date)


# Column name generation
columns = [
    "Product ID", "Product Name", "Category", "Brand", "Model", "SKU", "Quantity Sold",
    "Sale Price", "Retail Price", "Discount Amount", "Sale Date", "Customer ID",
    "Sales Channel", "Payment Method", "Shipping Cost", "Shipping Time", "Return Status",
    "Reason for Return", "Customer Feedback", "Profit", "Region", "Store Location"
]


df_tennis_sales_adjusted = pd.DataFrame(adjusted_data, columns=columns)

# Save to Excel
excel_file_path = 'adjusted_tennis_equipment_sales.xlsx'
df_tennis_sales_adjusted.to_excel(excel_file_path, index=False)

print(f"Data saved successfully to {excel_file_path}")
