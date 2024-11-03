# Get user input for variables
current_age = int(input("Enter your current age: "))
retirement_age = int(input("Enter your desired retirement age: "))
annual_contribution = float(input("Enter your annual contribution: "))
interest_rate = float(input("Enter the annual interest rate (e.g., enter 0.05 for 5%): "))

# Total savings and age counter
total_savings = 0
age = current_age

# While loop to calculate savings with interest until retirement
while age < retirement_age:
    total_savings += annual_contribution
    total_savings *= (1 + interest_rate)
    age += 1
    print(f"At age {age}, your savings would be: ${total_savings:.2f}")

print(f"\nTotal savings by retirement age {retirement_age}: ${total_savings:.2f}")
