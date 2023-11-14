import openpyxl

def calculate_mortgage_payment(loan_amount, interest_rate, loan_term):
    monthly_interest_rate = interest_rate / 12 / 100
    num_payments = loan_term * 12
    mortgage_payment = loan_amount * monthly_interest_rate * (1 + monthly_interest_rate) ** num_payments / ((1 + monthly_interest_rate) ** num_payments - 1)
    return mortgage_payment

# Input variables
property_price = 1000000
down_payment = 200000
loan_amount_new_property = property_price - down_payment
interest_rate_new_property = 4.5  # Assuming a 4.5% interest rate
loan_term_new_property = 30  # Assuming a 30-year loan term

loan_amount_existing_property = 300000.00
interest_rate_existing_property = 4.5  # Assuming a 4.5% interest rate
loan_term_existing_property = 30  # Assuming a 30-year loan term

# Monthly income and expenses
monthly_income = 4500  # Your monthly income
monthly_rental_income = 3480  # Rental income from the investment property
monthly_expenses = 1000  # Your monthly expenses

# Calculate mortgage payments for the new property and existing investment property
mortgage_payment_new_property = calculate_mortgage_payment(loan_amount_new_property, interest_rate_new_property, loan_term_new_property)
mortgage_payment_existing_property = calculate_mortgage_payment(loan_amount_existing_property, interest_rate_existing_property, loan_term_existing_property)

total_monthly_payment = mortgage_payment_new_property + mortgage_payment_existing_property

# Calculate the surplus/shortage of income after deducting expenses and mortgage payments
monthly_surplus = monthly_income + monthly_rental_income - monthly_expenses - total_monthly_payment

# Create a new Excel workbook and select the active sheet
workbook = openpyxl.Workbook()
sheet = workbook.active

# Write the headers
sheet['A1'] = "Property"
sheet['B1'] = "Loan Amount"
sheet['C1'] = "Interest Rate"
sheet['D1'] = "Loan Term"
sheet['E1'] = "Monthly Payment"

# Write the data for the new property
sheet['A2'] = "New Property"
sheet['B2'] = loan_amount_new_property
sheet['C2'] = interest_rate_new_property
sheet['D2'] = loan_term_new_property
sheet['E2'] = mortgage_payment_new_property

# Write the data for the existing investment property
sheet['A3'] = "Existing Investment Property"
sheet['B3'] = loan_amount_existing_property
sheet['C3'] = interest_rate_existing_property
sheet['D3'] = loan_term_existing_property
sheet['E3'] = mortgage_payment_existing_property

# Write the total monthly payment for both properties
sheet['A4'] = "Total Monthly Payment"
sheet['E4'] = total_monthly_payment

# Write the income, expenses, and surplus/shortage
sheet['A6'] = "Monthly Income"
sheet['B6'] = monthly_income
sheet['A7'] = "Rental Income"
sheet['B7'] = monthly_rental_income
sheet['A8'] = "Monthly Expenses"
sheet['B8'] = monthly_expenses
sheet['A9'] = "Monthly Surplus/Shortage"
sheet['B9'] = monthly_surplus

# Save the workbook
workbook.save("mortgage_analysis.xlsx")