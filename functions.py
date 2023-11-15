from datetime import timedelta
from datetime import datetime
import sys
import json
import random
import os
import re
import pandas as pd
import csv
from collections import defaultdict
from prettytable import PrettyTable
from vars import *

# Step 1: Read CSV Files Once
def read_csv_file(file_path):
    with open(file_path, 'r') as file:
        reader = csv.DictReader(file)
        return [row for row in reader]

def add_thousands_separator(number):
    integer_part = int(number)
    formatted_number = "{:,.0f}".format(integer_part)
    return formatted_number   

def excel_to_csv_start(excel_file):
    xls = pd.ExcelFile(excel_file)
    sheet_names = xls.sheet_names
    
    for sheet in sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet)
        csv_file = f"input/template_data - {sheet}.csv"
        df.to_csv(csv_file, index=False)
          
def get_first_name(names):
    if isinstance(names, str):
        names_list = names.split(",")
        return names_list[0].strip()
    elif isinstance(names, list):
        return names[0].strip()
    else:
        return None

def extract_tracking_number(tracking_number):
    """
    Extracts the numeric portion from a tracking number by removing non-numeric characters.
    """
    numeric_only = re.sub(r'\D', '', tracking_number)
    formatted_number = "({}) {}-{}".format(numeric_only[:3], numeric_only[3:6], numeric_only[6:])
    return formatted_number

def standardize_url(url):
    """
    Standardizes a URL to the format https://website.com.
    """
    # Remove leading/trailing whitespaces    
    url = url.strip()
    # Check if the URL starts with 'http://' or 'www.'
    if url.startswith("http://"):
        url = url.replace("http://", "https://")
    elif url.startswith("www."):
        url = url.replace("www.", "https://")

    # Add 'https://' if it's missing
    if not url.startswith("https://"):
        url = "https://" + url    
    return url

def load_postcard_rules(fileName):
    with open(fileName) as file:
        rules_data = json.load(file)
    return rules_data

def find_rule(rules_data, seller_avatar_group, sequence_step):
    for rule in rules_data:
        if (
            rule["Group Name"] == seller_avatar_group
            and str(rule["Sequence Step"]) == str(sequence_step)
        ):
            return str(rule["Template Name"]), str(rule["Template Number"]), rule["Gender"], rule["Type"]
    return None, None, None, None

def get_template_for_property(property_data):
    sequence_step = (property_data.num_dm % 4) + 1
    property_data.sequence_step = str(sequence_step)
    rules_data = load_postcard_rules(csv_to_json(INPUT_MAIL_RULES))
    template_name, template_number, gender, mail_type = find_rule(rules_data, property_data.seller_avatar_group, sequence_step)
    if template_name == "Checkletter":
        print("Checkletter")
    if template_name is not None:
        return template_name, template_number, gender, mail_type
    print("[ERROR] Template not found")

def generate_full_name(postcard_gender, property_data):
    # List of common American male first names
        
    male_first_names = [
        "James", "John", "Robert", "Michael", "William", "David", "Joseph", "Charles",
        "Thomas", "Daniel", "Matthew", "Anthony", "Donald", "Mark", "Paul", "Steven",
        "Andrew", "Kenneth", "George", "Joshua", "Kevin", "Brian", "Edward", "Ronald",
        "Timothy", "Jason", "Jeffrey", "Ryan", "Jacob", "Gary", "Nicholas", "Eric",
        "Stephen", "Jonathan", "Larry", "Justin", "Scott", "Brandon", "Benjamin", "Samuel",
        "Gregory", "Frank", "Alexander", "Raymond", "Patrick", "Jack", "Dennis", "Jerry",
        "Tyler", "Aaron", "Henry", "Douglas", "Peter", "Jose", "Adam", "Nathan", "Zachary"
    ]

    # List of common American female first names
    female_first_names = [
        "Mary", "Patricia", "Jennifer", "Linda", "Elizabeth", "Susan", "Jessica", "Sarah",
        "Karen", "Nancy", "Lisa", "Margaret", "Betty", "Dorothy", "Sandra", "Ashley",
        "Kimberly", "Donna", "Emily", "Michelle", "Carol", "Amanda", "Melissa", "Deborah",
        "Stephanie", "Rebecca", "Laura", "Sharon", "Cynthia", "Kathleen", "Amy", "Shirley",
        "Angela", "Helen", "Anna", "Brenda", "Pamela", "Nicole", "Samantha", "Katherine",
        "Emma", "Ruth", "Christine", "Catherine", "Debra", "Virginia", "Rachel", "Carolyn",
        "Janet", "Maria", "Heather", "Diane", "Julie", "Joyce", "Victoria", "Kelly"
    ]

    # List of common American last names
    last_names = [
        "Smith", "Johnson", "Williams", "Jones", "Brown", "Davis", "Miller", "Wilson",
        "Moore", "Taylor", "Anderson", "Thomas", "Jackson", "White", "Harris", "Martin",
        "Thompson", "Garcia", "Martinez", "Robinson", "Clark", "Rodriguez", "Lewis", "Lee",
        "Walker", "Hall", "Allen", "Young", "Hernandez", "King", "Wright", "Lopez", "Hill",
        "Scott", "Green", "Adams", "Baker", "Gonzalez", "Nelson", "Carter", "Mitchell",
        "Perez", "Roberts", "Turner", "Phillips", "Campbell", "Parker", "Evans", "Edwards"
    ]

    # Generate a random first name based on gender
    if postcard_gender == "Male":
        first_name = random.choice(male_first_names)
    else:
        first_name = random.choice(female_first_names)

    # Generate a random last name
    last_name = random.choice(last_names)
    area = f"({property_data.city}, {property_data.state})"
    # Return the generated full name
    return f"{first_name} {last_name[0]}. {area}"

def generate_qr_code_url(url):
    base_url = "https://api.qrserver.com/v1/create-qr-code/"
    params = {
        "data": url,
        "size": "200x200",
        "ecc": "L"
    }
    qr_code_url = base_url + "?" + "&".join(f"{key}={value}" for key, value in params.items())
    # return qr_code_url
    return url

def csv_to_json(csv_file):
    # Read CSV file and convert it to a list of dictionaries
    with open(csv_file, mode='r') as file:
        reader = csv.DictReader(file)
        data = list(reader)

    if not os.path.exists("workingFiles"):
            os.makedirs("workingFiles")

    # Write JSON data to a file
    with open("workingFiles/"+csv_file.replace(".csv", ".json").replace("input/", ""), mode='w') as file:
        json.dump(data, file, indent=4)
    
    return "workingFiles/"+csv_file.replace(".csv", ".json").replace("input/", "")

def get_drop_number(index, total_size, N):
    try:
        return (index * int(N)) // total_size + 1
    except ValueError:
        N = 4
        return (index * int(N)) // total_size + 1

def calculate_estimate_cash_offer(total_value, offer_price):    
    # Check if the offer rate is a range
    if '-' in offer_price:
        low, high = map(float, offer_price.split(' - '))
        low_offer = round(total_value * low, -2)
        high_offer = round(total_value * high, -2)
        return f"${add_thousands_separator(low_offer)} - ${add_thousands_separator(high_offer)}"
    else:
        try:
            offer = int(float(total_value)) * (float(offer_price) / 100)
            if offer < 15000 and float(offer_price) > 0:
                return "TBD"
            else:
                rounded_offer = round(offer, -2)  # Round to the nearest hundredth
                return f"${add_thousands_separator(int(rounded_offer))}"
        except:
            return "TBD"
             
def get_random_version(company_name, postcard):
    available_versions = set()
    with open(INPUT_CLIENTS_TEXTS, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Company Name'] == company_name:
                available_versions.add(row['Version'])
    postcard.version = random.choice(list(available_versions))
    return postcard.version

def create_client_folder(client):
    folder_name = client.company_name
    folder_path = os.path.join("results/", folder_name)

    try:
        # Comprueba si la carpeta ya existe
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            return folder_path
        else:
            return folder_path
    except OSError as e:
        print(f"Error al crear la carpeta '{folder_name}': {str(e)}")
        sys.exit(1)

def checking_test_percentage(client):
    # generate a random number between 0 and 100
    random_number = random.randint(1, 100)
    # if the random number is less than the given percentage, return 0
    if random_number <= client.test_percentage:
        return True
    else:
        return False

def get_text_by_postcard_name(company_name, postcard, text_number, estimate_cash_offer = 1):
    if estimate_cash_offer == "-":
        return ""
    with open(INPUT_CLIENTS_TEXTS, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            if postcard.mail_type == 'Postcard':
                if row['Company Name'] == company_name and row['Postcard Name'] == "T" + str(postcard.postcard_number) and postcard.version == row['Version']:
                    return row[f"text_{text_number}"]
            if postcard.mail_type == 'CheckLetter':
                if row['Company Name'] == company_name and row['Postcard Name'] == "CL" + str(postcard.postcard_number) and "a" == row['Version']:
                    return row[f"text_{text_number}"]
    return None

def add_30_days(drop_date: str) -> str:
    try:
        drop_date_obj = datetime.strptime(drop_date, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        try:
            drop_date_obj = datetime.strptime(drop_date, "%Y-%m-%d")
        except ValueError:
            return ""

    exp_date_obj = drop_date_obj + timedelta(days=30)
    exp_date = exp_date_obj.strftime("%Y-%m-%d")
    return exp_date

def calculate_cost(mail_pieces, default_postcard_size="8.5x5.5"):
    # Pricing dictionary
    pricing = {
        'Postcard 4x6': [(2500, 0.541), (5000, 0.501), (10000, 0.481), (20000, 0.466), (35000, 0.456), (float('inf'), 0.451)],
        'Postcard 4x6 (Google Streetview)': [(2500, 0.561), (5000, 0.521), (10000, 0.501), (20000, 0.486), (35000, 0.466), (float('inf'), 0.461)],
        'Postcard 8.5x5.5': [(2500, 0.607), (5000, 0.567), (10000, 0.547), (20000, 0.537), (35000, 0.532), (float('inf'), 0.522)],
        'Postcard 8.5x5.5 (Google Streetview)': [(2500, 0.627), (5000, 0.587), (10000, 0.567), (20000, 0.554), (35000, 0.537), (float('inf'), 0.547)],
        'Check Letter (Window Envelope)': [(2500, 0.833), (5000, 0.798), (10000, 0.766), (20000, 0.735), (35000, 0.723), (float('inf'), 0.723)],
    }
    
    # Count volume for each category
    volume_count = defaultdict(int)
    
    for piece in mail_pieces:
        if piece.mail_type == "Postcard":
            category = f"Postcard {default_postcard_size}"
            if piece.google_street_view_url:
                category += " (Google Streetview)"
        elif piece.mail_type == "CheckLetter":
            category = "Check Letter (Window Envelope)"
        else:
            continue  # Skip invalid types
        volume_count[category] += 1
    
    # Initialize table
    table = PrettyTable()
    table.field_names = ["Category", "Volume", "Cost"]
    
    # Calculate and print cost
    total_cost = 0
    total_volume = 0
    for category, volume in volume_count.items():
        adjusted_volume = volume if volume >= 2500 else 2500  # Adjust volume for pricing
        
        for volume_limit, rate in pricing[category]:
            if adjusted_volume <= volume_limit:
                cost = volume * rate  # Using the original 'volume' here for cost computation
                total_cost += cost
                total_volume += volume
                table.add_row([category, volume, f"${cost:.2f}"])  # Note: Using 'volume' here for correct print
                break
                
    print(table)
    print(f"Total Volume: {total_volume}")
    print(f"Total Cost: ${total_cost:.2f}")

