import csv
import re
import json
import random
import os
import sys

INPUT_CLIENTS_DATA      = 'input/template_data - clients_data.csv'
INPUT_COLORS_RULES      = 'input/template_data - color_rules.csv'
INPUT_POSTCARD_RULES    = 'input/template_data - postcard_sequence.csv'
INPUT_CLIENTS_LOGOS     = 'input/template_data - clients_logos.csv'
INPUT_CLIENTS_ADDRESS   = 'input/template_data - clients_address.csv'
DEBUG_MODE = False

class PostcardsList:
    def __init__(self, property_data, postcard_name=None, postcard_number=None,company_name=None, company_phone_number=None,
                 company_mailing_address=None, company_mailing_city=None, company_mailing_state=None,
                 company_mailing_zip=None, company_website=None, testimonial_name=None,
                 investor_full_name=None, targeted_message_1=None, targeted_message_2=None,
                 targeted_message_3=None, targeted_testimonial=None, owner_full_name=None,
                 owner_first_name=None, owner_property_address=None, owner_mailing_address=None,
                 owner_mailing_city=None, owner_mailing_state=None, owner_mailing_zip=None,
                 total_value=None, company_logo_url=None, google_street_view_url=None,
                 image_url=None, qr_code_url=None, credibility_logo_1_url=None,
                 credibility_logo_2_url=None, credibility_logo_3_url=None, credibility_logo_4_url=None, 
                 font_color_1 = None, font_color_2 = None, font_color_3 = None, font_color_4 = None, 
                 block_color_1 = None, block_color_2 = None
                 ):
        self.property_data = property_data
        self.postcard_name = postcard_name
        self.postcard_number = postcard_number
        self.company_name = company_name
        self.company_phone_number = company_phone_number
        self.company_mailing_address = company_mailing_address
        self.company_mailing_city = company_mailing_city
        self.company_mailing_state = company_mailing_state
        self.company_mailing_zip = company_mailing_zip
        self.company_website = company_website
        self.testimonial_name = testimonial_name
        self.investor_full_name = investor_full_name
        self.targeted_message_1 = targeted_message_1
        self.targeted_message_2 = targeted_message_2
        self.targeted_message_3 = targeted_message_3
        self.targeted_testimonial = targeted_testimonial
        self.owner_full_name = owner_full_name
        self.owner_first_name = owner_first_name
        self.owner_property_address = owner_property_address
        self.owner_mailing_address = owner_mailing_address
        self.owner_mailing_city = owner_mailing_city
        self.owner_mailing_state = owner_mailing_state
        self.owner_mailing_zip = owner_mailing_zip
        self.total_value = total_value
        self.company_logo_url = company_logo_url
        self.google_street_view_url = google_street_view_url
        self.image_url = image_url
        self.qr_code_url = qr_code_url
        self.cred_logo_1 = credibility_logo_1_url
        self.cred_logo_2 = credibility_logo_2_url
        self.cred_logo_3 = credibility_logo_3_url
        self.cred_logo_4 = credibility_logo_4_url
        self.font_color_1 = font_color_1
        self.font_color_2 = font_color_2
        self.font_color_3 = font_color_3
        self.font_color_4 = font_color_4
        self.block_color_1 = block_color_1
        self.block_color_2 = block_color_2

    def __str__(self):
        return f"Postcard Name: {self.postcard_number, self.postcard_name}\n" \
               f"Company Name: {self.company_name}\n" \
               f"Company Phone Number: {self.company_phone_number}\n" \
               f"Company Mailing Address: {self.company_mailing_address}\n" \
               f"Company Mailing City: {self.company_mailing_city}\n" \
               f"Company Mailing State: {self.company_mailing_state}\n" \
               f"Company Mailing ZIP: {self.company_mailing_zip}\n" \
               f"Company Website: {self.company_website}\n" \
               f"Testimonial Name: {self.testimonial_name}\n" \
               f"Investor Full Name: {self.investor_full_name}\n" \
               f"Targeted Message 1: {self.targeted_message_1}\n" \
               f"Targeted Message 2: {self.targeted_message_2}\n" \
               f"Targeted Message 3: {self.targeted_message_3}\n" \
               f"Targeted Testimonial: {self.targeted_testimonial}\n" \
               f"Owner Full Name: {self.owner_full_name}\n" \
               f"Owner First Name: {self.owner_first_name}\n" \
               f"Owner Property Address: {self.owner_property_address}\n" \
               f"Owner Mailing Address: {self.owner_mailing_address}\n" \
               f"Owner Mailing City: {self.owner_mailing_city}\n" \
               f"Owner Mailing State: {self.owner_mailing_state}\n" \
               f"Owner Mailing ZIP: {self.owner_mailing_zip}\n" \
               f"Total Value: {self.total_value}\n" \
               f"Company Logo URL: {self.company_logo_url}\n" \
               f"Google Street View URL: {self.google_street_view_url}\n" \
               f"Image URL: {self.image_url}\n" \
               f"QR Code URL: {self.qr_code_url}\n" \
               f"Credibility Logo 1 URL: {self.cred_logo_1}\n" \
               f"Credibility Logo 2 URL: {self.cred_logo_2}\n" \
               f"Credibility Logo 3 URL: {self.cred_logo_3}\n" \
               f"Credibility Logo 4 URL: {self.cred_logo_4}\n" \
               f"Font Color 1: {self.font_color_1}\n" \
               f"Font Color 2: {self.font_color_2}\n" \
               f"Font Color 3: {self.font_color_3}\n" \
               f"Font Color 4: {self.font_color_4}\n" \
               f"block_color_1: {self.block_color_1}\n" \
               f"block_color_2: {self.block_color_2}\n" \
     
    def assign_colors(self):
        with open(INPUT_COLORS_RULES, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if row['Company Name'] == self.company_name:
                    self.font_color_1 = row["T"+self.postcard_number] if row['Variable'] == 'font_color_1' else self.font_color_1
                    self.font_color_2 = row["T"+self.postcard_number] if row['Variable'] == 'font_color_2' else self.font_color_2
                    self.font_color_3 = row["T"+self.postcard_number] if row['Variable'] == 'font_color_3' else self.font_color_3
                    self.font_color_4 = row["T"+self.postcard_number] if row['Variable'] == 'font_color_4' else self.font_color_4
                    self.block_color_1 = row["T"+self.postcard_number] if row['Variable'] == 'block_color_1' else self.block_color_1
                    self.block_color_2 = row["T"+self.postcard_number] if row['Variable'] == 'block_color_2' else self.block_color_2

    def assign_company_information(self, client):
        self.company_name =             client.company_name
        self.company_phone_number =     client.contact_phone
        self.company_mailing_address =  client.mailing_address
        self.company_mailing_city =     ""
        self.company_mailing_state =    ""
        self.company_mailing_zip =      ""
        self.company_website =          client.website
        self.investor_full_name =       client.agent_name
        self.company_logo_url =         client.logo
        self.testimonial_name =         generate_full_name()
        self.qr_code_url =              generate_qr_code_url(client.website)

    def assign_owner_information(self, propertyData):
        self.targeted_message_1 =       propertyData.targeted_message_1
        self.targeted_message_2 =       propertyData.targeted_message_2
        self.targeted_message_3 =       propertyData.targeted_message_3
        self.targeted_testimonial =     propertyData.targeted_testimonial
        self.owner_full_name =          propertyData.owner_full_name
        self.owner_first_name =         propertyData.owner_first_name
        self.owner_property_address =   propertyData.owner_last_name
        self.owner_mailing_address =    propertyData.mailing_address
        self.owner_mailing_city =       propertyData.mailing_city
        self.owner_mailing_state =      propertyData.mailing_state
        self.owner_mailing_zip =        propertyData.mailing_zip
        self.total_value =              add_thousands_separator(propertyData.total_value)

    def assign_credibility_logos(self):
        with open(INPUT_CLIENTS_LOGOS, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Company Name'] == self.company_name and row['Type'].startswith('cred_logo'):
                    logo_number = str(row['Type'].replace("cred_logo_",""))
                    attribute_name = f"cred_logo_{logo_number}"
                    setattr(self, attribute_name, row['Logo_url'])

    def assign_tracking_number(self):
        self.company_phone_number = client.tracking_numbers[int(self.postcard_number) - 1]
    
class PropertyData:
    def __init__(self, folio, owner_full_name, owner_first_name, owner_last_name, address, city, state, zip_code, county,
                 mailing_address, mailing_city, mailing_state, mailing_zip, golden_address, golden_city, golden_state,
                 golden_zip_code, action_plans, property_status, score, distress_points, avatar, property_type,
                 link_properties, hidden_gems, tags, absentee, high_equity, downsizing, pre_foreclosure, vacant,
                 fifty_five_plus, estate, inter_family_transfer, divorce, taxes, probate, low_credit, code_violations,
                 bankruptcy, liens, eviction, thirty_sixty_days, judgment, debt_collection, total_value, num_sms, num_dm,
                 num_cold_call, seller_avatar_group, targeted_testimonial, main_distress_1, main_distress_2,
                 main_distress_3, main_distress_4, targeted_message_1, targeted_message_2, targeted_message_3,
                 targeted_message_4):
        self.folio = folio
        self.owner_full_name = owner_full_name
        self.owner_first_name = owner_first_name
        self.owner_last_name = owner_last_name
        self.address = address
        self.city = city
        self.state = state
        self.zip_code = zip_code
        self.county = county
        self.mailing_address = mailing_address
        self.mailing_city = mailing_city
        self.mailing_state = mailing_state
        self.mailing_zip = mailing_zip
        self.golden_address = golden_address
        self.golden_city = golden_city
        self.golden_state = golden_state
        self.golden_zip_code = golden_zip_code
        self.action_plans = action_plans
        self.property_status = property_status
        self.score = score
        self.distress_points = distress_points
        self.avatar = avatar
        self.property_type = property_type
        self.link_properties = link_properties
        self.hidden_gems = hidden_gems
        self.tags = tags
        self.absentee = absentee
        self.high_equity = high_equity
        self.downsizing = downsizing
        self.pre_foreclosure = pre_foreclosure
        self.vacant = vacant
        self.fifty_five_plus = fifty_five_plus
        self.estate = estate
        self.inter_family_transfer = inter_family_transfer
        self.divorce = divorce
        self.taxes = taxes
        self.probate = probate
        self.low_credit = low_credit
        self.code_violations = code_violations
        self.bankruptcy = bankruptcy
        self.liens = liens
        self.eviction = eviction
        self.thirty_sixty_days = thirty_sixty_days
        self.judgment = judgment
        self.debt_collection = debt_collection
        self.total_value = total_value
        self.num_sms = num_sms
        self.num_dm = num_dm
        self.num_cold_call = num_cold_call
        self.seller_avatar_group = seller_avatar_group
        self.targeted_testimonial = targeted_testimonial
        self.main_distress_1 = main_distress_1
        self.main_distress_2 = main_distress_2
        self.main_distress_3 = main_distress_3
        self.main_distress_4 = main_distress_4
        self.targeted_message_1 = targeted_message_1
        self.targeted_message_2 = targeted_message_2
        self.targeted_message_3 = targeted_message_3
        self.targeted_message_4 = targeted_message_4

class Client:
    def __init__(self, timestamp, company_name, contact_name, contact_email, contact_phone, mailing_address, website, logo, tracking_numbers, demographic, postcard_quantity, test_percentage, drop_date, postcard_size, featured_in_tv, bbb_accreditation, years_in_business, agent_name, website_link, response_rate, roi, deal_postcards, mail_house, postcard_designs, additional_comments, share_results):
        self.timestamp = timestamp
        self.company_name = company_name
        self.contact_name = contact_name
        self.contact_email = contact_email
        self.contact_phone = contact_phone
        self.mailing_address = mailing_address
        self.website = website
        self.logo = self.lookup_company_logo_url()
        self.tracking_numbers = tracking_numbers
        self.demographic = demographic
        self.postcard_quantity = postcard_quantity
        self.test_percentage = test_percentage
        self.drop_date = drop_date
        self.postcard_size = postcard_size
        self.featured_in_tv = featured_in_tv
        self.bbb_accreditation = bbb_accreditation
        self.years_in_business = years_in_business
        self.agent_name = agent_name
        self.website_link = website_link
        self.response_rate = response_rate
        self.roi = roi
        self.deal_postcards = deal_postcards
        self.mail_house = mail_house
        self.postcard_designs = postcard_designs
        self.additional_comments = additional_comments
        self.share_results = share_results
        self.get_client_address()
    
    def __str__(self):
        table = f"""
            ╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
            ║{self.center_text("Client Information", 118)}║
            ╟────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────╢
            ║ Timestamp: {self.timestamp:<113}
            ║ Company Name: {self.company_name:<110}
            ║ Contact Name: {self.contact_name:<110}
            ║ Contact Email: {self.contact_email:<109}
            ║ Contact Phone: {self.contact_phone:<109}
            ║ Mailing Address: {self.mailing_address:<105}
            ║ Website: {self.website:<112}
            ║ Logo: {self.logo:<115}
            ║ Tracking Numbers: {', '.join(self.tracking_numbers):<106}
            ║ Demographic: {self.demographic:<112}
            ║ Postcard Quantity: {self.postcard_quantity:<108}
            ║ Test Percentage: {self.test_percentage:<108}
            ║ Drop Date: {self.drop_date:<113}
            ║ Postcard Size: {self.postcard_size:<110}
            ║ Featured in TV: {self.featured_in_tv:<109}
            ║ BBB Accreditation: {self.bbb_accreditation:<109}
            ║ Years in Business: {self.years_in_business:<108}
            ║ Agent Name: {self.agent_name:<111}
            ║ Website Link: {self.website_link:<109}
            ║ Response Rate: {self.response_rate:<111}
            ║ ROI: {self.roi:<119}
            ║ Deal Postcards: {self.deal_postcards:<105}
            ║ Mail House: {self.mail_house:<111}
            ║ Postcard Designs: {self.postcard_designs:<106}
            ║ Additional Comments: {self.additional_comments:<103} 
            ╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
        \n"""
        return table

    def add_marketing_list(self, marketing_list):
        self.marketing_list = marketing_list

    def get_insights_mkt_list(self):
        print("=== Marketing List Insights ===")
        print(f"Client: {self.company_name}")
        print(f"Email: {self.contact_email}")
        print("")

        property_count = len(self.marketing_list)
        print(f"Total Properties: {property_count}")
        print("")

        # Calculate main distresses by count
        distress_count = {}
        for property in self.marketing_list:
            main_distress = property.main_distress_1
            if main_distress not in distress_count:
                distress_count[main_distress] = 1
            else:
                distress_count[main_distress] += 1

        # Sort distresses by count in descending order
        sorted_distresses = sorted(distress_count.items(), key=lambda x: x[1], reverse=True)

        # Print main distresses by count
        print("Main Distresses by Count:")
        for distress, count in sorted_distresses:
            print(f"- {distress}: {count}")
        print("")

        # Calculate average values
        total_value_sum = sum(property.total_value for property in self.marketing_list)
        average_total_value = total_value_sum / property_count

        # Print average values
        print("Average Values:")
        print("- Total Value: $", add_thousands_separator(average_total_value))
        print("")

        print("=== End of Insights ===")

    def lookup_company_logo_url(self):
        with open(INPUT_CLIENTS_LOGOS, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Company Name'] == self.company_name and row['Type'] == "Company":
                    return row['Logo_url']
        return ""
    
    def get_client_address(self):
        with open(INPUT_CLIENTS_ADDRESS, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Company Name'] == self.company_name:
                    self.company_mailing_address = row['Company Mailing Address']
                    self.company_mailing_city = row['Company Mailing City']
                    self.company_mailing_state = row['Company Mailing State']
                    self.company_mailing_zip = row['Company Mailing ZIP']
                    break
            else:
                # Handle the case where the company name is not found
                # Set default values or perform any necessary actions
                self.company_mailing_address = "Not found"
                self.company_mailing_city = "Not found"
                self.company_mailing_state = "Not found"
                self.company_mailing_zip = "Not found"    
    
    @staticmethod
    def center_text(text, width):
        return text.center(width)
          
def add_thousands_separator(number):
    integer_part = int(number)
    formatted_number = "{:,.0f}".format(integer_part)
    return formatted_number         
    
def read_clients_data(file_name):
    clients = []
    with open(file_name, 'r', newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            client = Client(
                row['Timestamp'],
                row['Company Name'],
                row['Name of Main Point of Contact'],
                row['Email of Main Point of Contact'],
                extract_tracking_number(row['Cell Phone of Main Point of Contact']),
                row['Company mailing Address (Street, City, State, ZIP) for return mail'],
                standardize_url(row['Company Website (seller facing)']),
                standardize_url(row['Company Logo']),
                [
                    extract_tracking_number(row['Brand New Tracking Number 1']),
                    extract_tracking_number(row['Brand New Tracking Number 2']),
                    extract_tracking_number(row['Brand New Tracking Number 3']),
                    extract_tracking_number(row['Brand New Tracking Number 4'])
                ],
                row['What is your current customer demographic?'],
                row['How many postcards would you like to send?'],
                row['What % of your list would you like to test?'],
                row['When will be your next Direct Mail drop?'],
                row['What is your postcard size preference?'],
                row['Have you been featured in TV?'],
                row['Do you have a BBB accreditation?'],
                row['How many years in business?'],
                get_first_name(row['Agent/s full name (if any)']),
                standardize_url(row['Provide a link below where you want the customers to view your website/testimonials or more about your company:']),
                row['What is your current response rate (or call rate) from Direct Mail?'],
                row['What is your current ROI from Direct Mail?'],
                row['How many postcards does it take you to get a deal?'],
                row['Who is your preferred Direct Mail House?'],
                row['Upload up to 3 of your most recent postcard designs. Only upload the ones that you have used in the past 12 months.'],
                row['Any other additional comments'],
                row['Do you agree to share the results in a monthly basis?']
            )
            clients.append(client)
    return clients

def get_first_name(names):
    if isinstance(names, str):
        names_list = names.split(",")
        return names_list[0].strip()
    elif isinstance(names, list):
        return names[0].strip()
    else:
        return None

def read_marketing_list_csv(file_name):
    marketing_list = []
    with open(file_name, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            folio = row['FOLIO']
            owner_full_name = row['OWNER FULL NAME']
            owner_first_name = row['OWNER FIRST NAME']
            owner_last_name = row['OWNER LAST NAME']
            address = row['ADDRESS']
            city = row['CITY']
            state = row['STATE']
            zip_code = row['ZIP']
            county = row['COUNTY']
            mailing_address = row['MAILING ADDRESS']
            mailing_city = row['MAILING CITY']
            mailing_state = row['MAILING STATE']
            mailing_zip = row['MAILING ZIP']
            golden_address = row['GOLDEN ADDRESS']
            golden_city = row['GOLDEN CITY']
            golden_state = row['GOLDEN STATE']
            golden_zip_code = row['GOLDEN ZIP CODE']
            action_plans = row['ACTION PLANS']
            property_status = row['PROPERTY STATUS']
            score = row['SCORE']
            distress_points = row['LIKELY DEAL SCORE']
            avatar = row['BUYBOX SCORE']
            total_value = int(row['TOTAL VALUE'].replace('$', '').replace(',', ''))
            property_type = row['PROPERTY TYPE']
            link_properties = row['LINK PROPERTIES']
            tags = row['TAGS']
            hidden_gems = row['HIDDENGEMS']
            absentee = row['ABSENTEE']
            high_equity = row['HIGH EQUITY']
            downsizing = row['DOWNSIZING']
            pre_foreclosure = row['PRE-FORECLOSURE']
            vacant = row['VACANT']
            fifty_five_plus = row['55+']
            estate = row['ESTATE']
            inter_family_transfer = row['INTER FAMILY TRANSFER']
            divorce = row['DIVORCE']
            taxes = row['TAXES']
            probate = row['PROBATE']
            low_credit = row['LOW CREDIT']
            code_violations = row['CODE VIOLATIONS']
            bankruptcy = row['BANKRUPTCY']
            liens = row['LIENS']
            eviction = row['EVICTION']
            thirty_sixty_days = row['30-60 DAYS']
            judgment = row['JUDGEMENT']
            debt_collection = row['DEBT COLLECTION']
            # num_dm = add_mkt_count(row['MARKETING DM COUNT'])
            main_distress_1 = row['MAIN DISTRESS #1']
            main_distress_2 = row['MAIN DISTRESS #2']
            main_distress_3 = row['MAIN DISTRESS #3']
            main_distress_4 = row['MAIN DISTRESS #4']
            targeted_message_1 = row['TARGETED MESSAGE #1']
            targeted_message_2 = row['TARGETED MESSAGE #2']
            targeted_message_3 = row['TARGETED MESSAGE #3']
            targeted_message_4 = row['TARGETED MESSAGE #4']
            targeted_group_name = row['TARGETED GROUP NAME']
            targeted_group_message = row['TARGETED GROUP MESSAGE']
            
            try:
                num_dm = int(row['MARKETING DM COUNT'])
            except:
                num_dm = 1
        
            property_data = PropertyData(folio, owner_full_name, owner_first_name, owner_last_name, address, city,
                                         state, zip_code, county, mailing_address, mailing_city, mailing_state,
                                         mailing_zip, golden_address, golden_city, golden_state, golden_zip_code,
                                         action_plans, property_status, score, distress_points, avatar, property_type,
                                         link_properties, hidden_gems, tags, absentee, high_equity, downsizing,
                                         pre_foreclosure, vacant, fifty_five_plus, estate, inter_family_transfer,
                                         divorce, taxes, probate, low_credit, code_violations, bankruptcy, liens,
                                         eviction, thirty_sixty_days, judgment, debt_collection, total_value, num_dm,
                                         num_dm, num_dm, targeted_group_name, targeted_group_message,
                                         main_distress_1, main_distress_2, main_distress_3, main_distress_4,
                                         targeted_message_1, targeted_message_2, targeted_message_3, targeted_message_4)
            marketing_list.append(property_data)
    return marketing_list

def extract_tracking_number(tracking_number):
    """
    Extracts the numeric portion from a tracking number by removing non-numeric characters.
    """
    numeric_only = re.sub(r'\D', '', tracking_number)
    return numeric_only

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

def get_template_for_property(property_data):
    # load_sequence_step
    sequence_step = (property_data.num_dm % 4) + 1
    rules_data = load_postcard_rules(csv_to_json(INPUT_POSTCARD_RULES))
    for rule in rules_data:
        if (
            rule["Group Name"]                  == property_data.seller_avatar_group
            and str(rule["Sequence Step"])      == str(sequence_step)
        ):
            if DEBUG_MODE:
                print("Template : "     , rule["Template Number"],  "\t\tTemplate Name: ",  rule["Template Name"] )
                print("Sequence Step: " , rule["Sequence Step"],    "\tGroup Name: ",       rule["Group Name"])
                return str(rule["Template Name"]), str(rule["Template Number"])
            else:
                return str(rule["Template Name"]), str(rule["Template Number"])
    print("[ERROR] Template not found: ",)
    return None, None

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

def generate_full_name():
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
    gender = random.choice(["male", "female"])
    if gender == "male":
        first_name = random.choice(male_first_names)
    else:
        first_name = random.choice(female_first_names)

    # Generate a random last name
    last_name = random.choice(last_names)

    # Return the generated full name
    return f"{first_name} {last_name[0]}."

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

def create_csv_file(postcards_list, client):
    fieldnames = [
        "FOLIO",
        "OWNER FULL NAME",
        "OWNER FIRST NAME",
        "OWNER LAST NAME",
        "ADDRESS",
        "CITY",
        "STATE",
        "ZIP",
        "COUNTY",
        "MAILING ADDRESS",
        "MAILING CITY",
        "MAILING STATE",
        "MAILING ZIP",
        "GOLDEN ADDRESS",
        "GOLDEN CITY",
        "GOLDEN STATE",
        "GOLDEN ZIP CODE",
        "ACTION PLANS",
        "PROPERTY STATUS",
        "SCORE",
        "LIKELY DEAL SCORE",
        "BUYBOX SCORE",
        "TOTAL VALUE",
        "PROPERTY TYPE",
        "LINK PROPERTIES",
        "TAGS",
        "HIDDENGEMS",
        "ABSENTEE",
        "HIGH EQUITY",
        "DOWNSIZING",
        "PRE-FORECLOSURE",
        "VACANT",
        "55+",
        "ESTATE",
        "INTER FAMILY TRANSFER",
        "DIVORCE",
        "TAXES",
        "PROBATE",
        "LOW CREDIT",
        "CODE VIOLATIONS",
        "BANKRUPTCY",
        "LIENS",
        "EVICTION",
        "30-60 DAYS",
        "JUDGEMENT",
        "DEBT COLLECTION",
        "MARKETING DM COUNT",
        "MAIN DISTRESS #1",
        "MAIN DISTRESS #2",
        "MAIN DISTRESS #3", 
        "MAIN DISTRESS #4",
        "targeted_message_1",
        "targeted_message_2",
        "targeted_message_3",
        "targeted_message_4",
        "TARGETED GROUP NAME",
        "targeted_test", # Targeted group message
        "Postcard Name",
        "seller_full_name",
        "seller_first_name",
        "seller_mailing_add",
        "company_name",
        "company_phone_number", # Company phone number
        "company_mailing_add",
        "Company Mailing Address",
        "Company Mailing City",
        "Company Mailing State",
        "Company Mailing ZIP",
        "company_website",
        "test_name",
        "investor_full_name",
        "company_logo",
        "qr_code",
        "cred_logo_1",
        "cred_logo_2",
        "cred_logo_3",
        "cred_logo_4",
        "font_color_1",
        "font_color_2",
        "font_color_3",
        "font_color_4",
        "block_color_1",
        "block_color_2",
        "google_street-view",
        "image"
    ]            
    with open("output/" + client.company_name + "/MktList-" +client.company_name + ".csv", "w", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for postcard in postcards_list:
            seller_mailing_add =    postcard.property_data.mailing_address + ", " + postcard.property_data.mailing_city + " " + postcard.property_data.mailing_state + ", " + postcard.property_data.mailing_zip
            writer.writerow({
                "FOLIO":                        postcard.property_data.folio,
                "OWNER FULL NAME":              postcard.property_data.owner_full_name,
                "OWNER FIRST NAME":             postcard.property_data.owner_first_name,
                "OWNER LAST NAME":              postcard.property_data.owner_last_name,
                "ADDRESS":                      postcard.property_data.address,
                "CITY":                         postcard.property_data.city,
                "STATE":                        postcard.property_data.state,
                "ZIP":                          postcard.property_data.zip_code,
                "COUNTY":                       postcard.property_data.county,
                "MAILING ADDRESS":              postcard.property_data.mailing_address,
                "MAILING CITY":                 postcard.property_data.mailing_city,
                "MAILING STATE":                postcard.property_data.mailing_state,
                "MAILING ZIP":                  postcard.property_data.mailing_zip,
                "GOLDEN ADDRESS":               postcard.property_data.golden_address,
                "GOLDEN CITY":                  postcard.property_data.golden_city,
                "GOLDEN STATE":                 postcard.property_data.golden_state,
                "GOLDEN ZIP CODE":              postcard.property_data.golden_zip_code,
                "ACTION PLANS":                 postcard.property_data.action_plans,
                "PROPERTY STATUS":              postcard.property_data.property_status,
                "SCORE":                        postcard.property_data.score,
                "LIKELY DEAL SCORE":            postcard.property_data.distress_points,
                "BUYBOX SCORE":                 postcard.property_data.avatar,
                "TOTAL VALUE":                  postcard.property_data.total_value,
                "PROPERTY TYPE":                postcard.property_data.property_type,
                "LINK PROPERTIES":              postcard.property_data.link_properties,
                "TAGS":                         postcard.property_data.tags,
                "HIDDENGEMS":                   postcard.property_data.hidden_gems,
                "ABSENTEE":                     postcard.property_data.absentee,
                "HIGH EQUITY":                  postcard.property_data.high_equity,
                "DOWNSIZING":                   postcard.property_data.downsizing,
                "PRE-FORECLOSURE":              postcard.property_data.pre_foreclosure,
                "VACANT":                       postcard.property_data.vacant,
                "55+":                          postcard.property_data.fifty_five_plus,
                "ESTATE":                       postcard.property_data.estate,
                "INTER FAMILY TRANSFER":        postcard.property_data.inter_family_transfer,
                "DIVORCE":                      postcard.property_data.divorce,
                "TAXES":                        postcard.property_data.taxes,
                "PROBATE":                      postcard.property_data.probate,
                "LOW CREDIT":                   postcard.property_data.low_credit,
                "CODE VIOLATIONS":              postcard.property_data.code_violations,
                "BANKRUPTCY":                   postcard.property_data.bankruptcy,
                "LIENS":                        postcard.property_data.liens,
                "EVICTION":                     postcard.property_data.eviction,
                "30-60 DAYS":                   postcard.property_data.thirty_sixty_days,
                "JUDGEMENT":                    postcard.property_data.judgment,
                "DEBT COLLECTION":              postcard.property_data.debt_collection,
                "MARKETING DM COUNT":           postcard.property_data.num_dm,
                "MAIN DISTRESS #1":             postcard.property_data.main_distress_1,
                "MAIN DISTRESS #2":             postcard.property_data.main_distress_2,
                "MAIN DISTRESS #3":             postcard.property_data.main_distress_3,
                "MAIN DISTRESS #4":             postcard.property_data.main_distress_4,
                "targeted_message_1":          postcard.property_data.targeted_message_1,
                "targeted_message_2":          postcard.property_data.targeted_message_2,
                "targeted_message_3":          postcard.property_data.targeted_message_3,
                "targeted_message_4":          postcard.property_data.targeted_message_4,
                "TARGETED GROUP NAME":          postcard.property_data.seller_avatar_group,
                "targeted_test":       postcard.property_data.targeted_testimonial,
                "Postcard Name":                "T" + postcard.postcard_number,
                "seller_full_name":             postcard.property_data.owner_full_name,
                "seller_first_name":            postcard.property_data.owner_first_name,
                "seller_mailing_add":           seller_mailing_add,
                "company_name":                 postcard.company_name,
                "company_phone_number":         postcard.company_phone_number,
                "company_mailing_add":          client.mailing_address,
                "Company Mailing Address":      client.company_mailing_address,
                "Company Mailing City":         client.company_mailing_city,
                "Company Mailing State":        client.company_mailing_state,
                "Company Mailing ZIP":          client.company_mailing_zip,
                "company_website":              postcard.company_website,
                "test_name":             postcard.testimonial_name,
                "investor_full_name":           postcard.investor_full_name,
                "company_logo":                 postcard.company_logo_url,
                "qr_code":                  postcard.qr_code_url,
                "cred_logo_1":                  postcard.cred_logo_1,
                "cred_logo_2":                  postcard.cred_logo_2,
                "cred_logo_3":                  postcard.cred_logo_3,
                "cred_logo_4":                 postcard.cred_logo_4,
                "font_color_1":                 postcard.font_color_1,
                "font_color_2":                 postcard.font_color_2,
                "font_color_3":                 postcard.font_color_3,
                "font_color_4":                 postcard.font_color_4,
                "block_color_1":                postcard.block_color_1,
                "block_color_2":                postcard.block_color_2,
                "google_street-view":           "",
                "image":                        ""
            })

def create_client_folder(client):
    folder_name = client.company_name
    folder_path = os.path.join("output/", folder_name)

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

def find_marketingList(company_name):
    marketing_list_folder = "input/marketingLists"
    for file_name in os.listdir(marketing_list_folder):
        if file_name.endswith(".csv") and file_name[:-4] == company_name:
            file_path = os.path.join(marketing_list_folder, file_name)
            return file_path
    return False

if __name__ == "__main__":
    # Read clients' data from the specified CSV file
    clients = read_clients_data(INPUT_CLIENTS_DATA)
    # Iterate over each client
    for client in clients:
        if find_marketingList(client.company_name):
            create_client_folder(client)
            postcards_list = list()
            # Add the marketing list data to the client instance by reading the CSV file
            client.add_marketing_list(read_marketing_list_csv(find_marketingList(client.company_name)))
            # Retrieve and print the insights of the marketing list for the client
            if DEBUG_MODE:
                pass
                # client.get_insights_mkt_list()
    
            # Iterate over each property data in the client's marketing list
            i = 0
            for property_data in client.marketing_list:
                # Determine the template to use for each property data
                postcard_name, postcard_number = get_template_for_property(property_data)
                newPostcardTemplate = PostcardsList(property_data, postcard_name, postcard_number)
                newPostcardTemplate.assign_company_information(client)
                newPostcardTemplate.assign_owner_information(property_data) 
                newPostcardTemplate.assign_colors()
                newPostcardTemplate.assign_credibility_logos()   
                newPostcardTemplate.assign_tracking_number()   
                postcards_list.append(newPostcardTemplate) 
                print("Progress: " + str(i) + "/" + str(len(client.marketing_list)))
                i += 1
                

            create_csv_file(postcards_list, client)          
               
