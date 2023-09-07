import re

def extract_company_data(file_path):
    # Create a list to store the captured data
    captured_data = []

    # Define the regex pattern
    pattern = r'([A-Z]+\d+)\s+([\w\s,()]+)'

    # Read the .txt file
    with open(file_path, 'r', encoding='utf-8') as file:
        capturing = False
        company_data = []

        for line in file:
            # Check if the line matches the pattern for company data
            match = re.match(pattern, line.strip())
            
            if match:
                # Capture the company data and set capturing flag to True
                company_code, company_name = match.groups()
                company_data.append((company_code, company_name))
                print(company_code, company_name)
                capturing = True
            elif capturing:
                # Add the line to the captured company data
                company_data.append(line.strip())
            else:
                # If not capturing, continue to the next line
                continue

        # Append the captured company data to the list
        if company_data:
            captured_data.append(company_data)

    return captured_data

# Example usage:
data = extract_company_data('extracted_text.txt')
for i in range(4):
    print(data[i])
