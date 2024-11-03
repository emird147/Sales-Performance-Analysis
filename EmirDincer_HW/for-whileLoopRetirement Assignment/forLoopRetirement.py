current_age = int(input("Enter your current age: "))
retirement_age = int(input("Enter your desired retirement age: "))
annual_contribution = float(input("Enter your annual contribution: "))
monthly_expenses = float(input("Enter your estimated monthly expenses in retirement: "))
years_in_retirement = int(input("Enter the number of years you expect to live in retirement: "))
interest_rate = float(input("Enter the annual interest rate (e.g., enter 0.05 for 5%): "))

# Total Savings
total_savings = 0

# For loop to calculate total savings by retirement
for age in range(current_age, retirement_age):
    total_savings += annual_contribution
    total_savings *= (1 + interest_rate)  # Compound interest on total savings
    print(f"At age {age + 1}, your savings would be: ${total_savings:.2f}")

print(f"\nTotal savings by retirement age {retirement_age}: ${total_savings:.2f}")

# Required monthly expenses
annual_expenses = monthly_expenses * 12
required_savings = 0

# Calculate if total savings can last for the specified retirement years
for year in range(years_in_retirement):
    required_savings += annual_expenses
    required_savings /= (1 + interest_rate) 

if total_savings >= required_savings:
    print(f"\nYou have enough savings to cover your expenses for {years_in_retirement} years in retirement.")
else:
    print(f"\nYour savings may not be enough. You would need approximately ${required_savings - total_savings:.2f} more.")