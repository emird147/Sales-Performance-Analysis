# Get user input for variables
current_age = int(input("Enter your current age: "))
retirement_age = int(input("Enter your desired retirement age: "))
annual_contribution = float(input("Enter your annual contribution: "))

# Total savings
total_savings = 0

# For loop to calculate total savings at retirement
for age in range(current_age, retirement_age):
    total_savings += annual_contribution
    print(f"At age {age + 1}, your savings would be: ${total_savings:.2f}")

print(f"\nTotal savings by retirement age {retirement_age}: ${total_savings:.2f}")
