import csv
import re
import sys
import json
import random
import os
import sys
import pandas as pd
from datetime import datetime, timedelta

INPUT_FILE                  = 'input/template_data.xlsx'
INPUT_CLIENTS_DATA          = 'input/template_data - clients_data.csv'
INPUT_COLORS_RULES          = 'input/template_data - color_rules.csv'
INPUT_MAIL_RULES        = 'input/template_data - mail_sequence.csv'
INPUT_CLIENTS_LOGOS         = 'input/template_data - clients_logos.csv'
INPUT_CLIENTS_ADDRESS       = 'input/template_data - clients_address.csv'
INPUT_CLIENTS_PHONES        = 'input/template_data - clients_phones.csv'
INPUT_CLIENTS_BG_IMG        = "input/template_data - clients_bg_img.csv"
INPUT_CLIENTS_OFFER_PRICE   = "input/template_data - clients_offer_price.csv"
INPUT_CLIENTS_TEXTS         = "input/template_data - clients_texts.csv"
INPUT_CLIENTS_CAMPAIGN      = "input/template_data - clients_campaign_data.csv"
DEBUG_MODE = True

class PostcardsList:
    def __init__(self, 
                 property_data, 
                 postcard_name=None, 
                 postcard_number=None,
                 postcard_gender=None,
                 company_name=None, 
                 company_phone_number=None,
                 company_mailing_address=None, 
                 company_mailing_city=None, 
                 company_mailing_state=None,
                 company_mailing_zip=None, 
                 company_website=None, 
                 testimonial_name=None,
                 investor_full_name=None, 
                 targeted_message_1=None, 
                 targeted_message_2=None,
                 targeted_message_3=None, 
                 targeted_testimonial=None, 
                 owner_full_name=None,
                 owner_first_name=None, 
                 owner_property_address=None, 
                 owner_mailing_address=None,
                 owner_mailing_city=None, 
                 owner_mailing_state=None, 
                 owner_mailing_zip=None,
                 total_value=None, 
                 company_logo_url=None, 
                 google_street_view_url=None,
                 image_url=None, 
                 qr_code_url=None, 
                 credibility_logo_1_url=None,
                 credibility_logo_2_url=None, 
                 credibility_logo_3_url=None, 
                 credibility_logo_4_url=None, 
                 font_color_1 = None, 
                 font_color_2 = None, 
                 font_color_3 = None, 
                 font_color_4 = None, 
                 block_color_1 = None, 
                 block_color_2 = None
                 ):
        self.property_data = property_data
        self.postcard_name = postcard_name
        self.postcard_number = postcard_number
        self.postcard_gender = postcard_gender
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
               f"Seller Avatar Group: {self.property_data.seller_avatar_group}\n" \
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
     
    def assign_colors(self, client):
        """
        The function assigns colors to different variables based on the company name and postcard number.
        """
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
        self.testimonial_name =         generate_full_name(self.postcard_gender, self.property_data)
        self.qr_code_url =              generate_qr_code_url(client.website_link)

    def assign_owner_information(self, propertyData, client):
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

    def assign_logos(self, client):
        with open(INPUT_CLIENTS_LOGOS, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Company Name'] == self.company_name and row['Type'].startswith('cred_logo'):
                    logo_number = str(row['Type'].replace("cred_logo_",""))
                    attribute_name = f"cred_logo_{logo_number}"
                    setattr(self, attribute_name, row['Logo_url_T' + self.postcard_number])
                if row['Company Name'] == self.company_name and row['Type'] == 'Company':
                    company_logo = row['Logo_url_T' + str(self.postcard_number)]
                    self.company_logo_url = company_logo

    def assign_tracking_number(self, client):
        self.company_phone_number = client.tracking_numbers[int(self.postcard_number) - 1]
    
    def assign_bg_image(self, client):
        with open(INPUT_CLIENTS_BG_IMG, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Company Name'] == self.company_name and str(row["Template Number"]) == str(self.postcard_number):             
                    self.bg_img = row[self.property_data.seller_avatar_group]
                    return True
            print("Could not find a matching background for this postcard number\n\n")
            print(self)
            sys.exit(1)
            return False
    
    def assign_google_street_view(self):
        if str(self.postcard_number) == "1":
            self.google_street_view_url = True
        else:
            self.google_street_view_url = False
    
class PropertyData:
    def __init__(self, folio, owner_full_name, owner_first_name, owner_last_name, address, city, state, zip_code, county,
                 mailing_address, mailing_city, mailing_state, mailing_zip, golden_address, golden_city, golden_state,
                 golden_zip_code, action_plans, property_status, score, distress_points, avatar, property_type,
                 link_properties, hidden_gems, tags, absentee, high_equity, downsizing, pre_foreclosure, vacant,
                 fifty_five_plus, estate, inter_family_transfer, divorce, taxes, probate, low_credit, code_violations,
                 bankruptcy, liens_city, liens_other, liens_utility, liens_hoa,  liens_mechanic, eviction, thirty_sixty_days, judgment, debt_collection, total_value, num_dm,
                 seller_avatar_group, targeted_testimonial, main_distress_1, main_distress_2,
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
        self.liens_city = liens_city
        self.liens_other = liens_other
        self.liens_utility = liens_utility
        self.liens_hoa = liens_hoa
        self.liens_mechanic = liens_mechanic
        self.eviction = eviction
        self.thirty_sixty_days = thirty_sixty_days
        self.judgment = judgment
        self.debt_collection = debt_collection
        self.total_value = total_value
        # self.num_sms = num_sms
        self.num_dm = num_dm
        # self.num_cold_call = num_cold_call
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
    def __init__(self, client_id, company_name, contact_name, contact_email, contact_phone, mailing_address, website, logo, tracking_numbers, demographic, postcard_quantity, test_percentage, range_offer, drop_date, postcard_size, featured_in_tv, bbb_accreditation, years_in_business, agent_name, website_link, response_rate, roi, deal_postcards, mail_house, postcard_designs, additional_comments, share_results):
        self.client_id = client_id
        self.company_name = company_name
        self.contact_name = contact_name
        self.contact_email = contact_email
        self.contact_phone = contact_phone
        self.mailing_address = mailing_address
        self.website = website
        self.tracking_numbers = tracking_numbers
        self.demographic = demographic
        self.postcard_quantity = postcard_quantity
        self.test_percentage = test_percentage
        self.range_offer = range_offer
        self.drop_date = ""
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
        self.get_client_campaign_data()
        self.get_offer_price()
        self.text_1 = []
        self.text_2 = []
        self.text_3 = []
        self.text_4 = []
        self.text_5 = []
        self.text_6 = []
        self.text_7 = []
        self.text_8 = []
        self.text_9 = []
        self.text_10 = []

    def __str__(self):
        table = f"""
            ╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
            ║{self.center_text("Client Information", 118)}║
            ╟────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────╢
            ║ Timestamp: {self.client_id:<113}
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
    
    def get_client_campaign_data(self):
        with open(INPUT_CLIENTS_CAMPAIGN, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['client_name'] == self.company_name:
                    self.campaign_name = row['campaign_name']
                    self.drop_nums = row['drop_nums']
                    self.drop_date = row['drop_date']
                    break
            else:
                self.campaign_name = "Not found"
                self.drop_nums = "Not found"
    
    def get_offer_price(self):
        with open(INPUT_CLIENTS_OFFER_PRICE, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Company Name'] == self.company_name:
                    self.offer_price = int(row['Price Offer Rate'])/100
                    break
            else:
                self.offer_price = .85 
    
    @staticmethod
    def center_text(text, width):
        return text.center(width)

def excel_to_csv_start(excel_file):
    xls = pd.ExcelFile(excel_file)
    sheet_names = xls.sheet_names
    
    for sheet in sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet)
        csv_file = f"input/template_data - {sheet}.csv"
        df.to_csv(csv_file, index=False)
          
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
                row['client_id'],
                row['Company Name'],
                row['Name of Main Point of Contact'],
                row['Email of Main Point of Contact'],
                row['Cell Phone of Main Point of Contact'],
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
                int(float(row['What % of your list would you like to test?'])),
                row['range_offer'],
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
            liens_city = row['LIENS CITY/COUNTY']
            liens_other = row['LIENS OTHER']
            liens_utility = row['LIENS UTILITY']
            liens_hoa = row['LIENS HOA']
            liens_mechanic = row['LIENS MECHANIC']
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
            
            if row['TARGETED GROUP NAME'] == "No distresses":
                targeted_group_name = "Absentee"
                
            try:
                num_dm = int(row['MARKETING DM COUNT'])
            except:
                num_dm = 0
                
            try:
                total_value = int(row['TOTAL VALUE'].replace('$', '').replace(',', ''))
            except:
                if row['TOTAL VALUE'] == 'N/A' or row['TOTAL VALUE'] == 'n/a' or row['TOTAL VALUE'] == "":
                    total_value = 0
        
            property_data = PropertyData(folio, owner_full_name, owner_first_name, owner_last_name, address, city,
                                         state, zip_code, county, mailing_address, mailing_city, mailing_state,
                                         mailing_zip, golden_address, golden_city, golden_state, golden_zip_code,
                                         action_plans, property_status, score, distress_points, avatar, property_type,
                                         link_properties, hidden_gems, tags, absentee, high_equity, downsizing,
                                         pre_foreclosure, vacant, fifty_five_plus, estate, inter_family_transfer,
                                         divorce, taxes, probate, low_credit, code_violations, bankruptcy, liens_city,
                                         liens_other, liens_utility, liens_hoa,  liens_mechanic,
                                         eviction, thirty_sixty_days, judgment, debt_collection, total_value, num_dm,
                                         targeted_group_name, targeted_group_message,
                                         main_distress_1, main_distress_2, main_distress_3, main_distress_4,
                                         targeted_message_1, targeted_message_2, targeted_message_3, targeted_message_4)
            marketing_list.append(property_data)
    return marketing_list

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

def get_template_for_property(property_data):
    
    if property_data.total_value < 5000 or property_data.total_value == "":
        sequence_step = ((property_data.num_dm + 1) % 4) + 1
    else:
        sequence_step = (property_data.num_dm % 4) + 1
    
    property_data.sequence_step = str(sequence_step)
    rules_data = load_postcard_rules(csv_to_json(INPUT_MAIL_RULES))
    for rule in rules_data:       
        if (
            rule["Group Name"]                  == property_data.seller_avatar_group
            and str(rule["Sequence Step"])      == str(sequence_step)
        ):              
            return str(rule["Template Name"]), str(rule["Template Number"]), rule["Gender"]
        
    
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

def get_drop_number(index, total_size, N):
    return (index * int(N)) // total_size + 1

def calculate_estimate_cash_offer(client, total_value, offer_price):
    if client.range_offer == "True":
        low_estimated_offer     = round(total_value * 0.8, -2) 
        high_estimated_offer    = round(total_value * 1.2, -2) 
        return str("$" + str(add_thousands_separator(low_estimated_offer)) + " - " + "$" + str(add_thousands_separator(high_estimated_offer)))

    else:
        offer = int(total_value) * offer_price
        if offer < 15000 and offer_price > 0:
            return "TBD"
        else:
            rounded_offer = round(offer, -2)  # Round to the nearest hundredth
            return str("$" + str(int(rounded_offer))) 
                
def get_random_version(company_name, postcard):
    available_versions = set()
    with open(INPUT_CLIENTS_TEXTS, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Company Name'] == company_name:
                available_versions.add(row['Version'])
    postcard.version = random.choice(list(available_versions))
    return postcard.version

def create_csv_files(postcards_list, client):
    fieldnames = [
        "client_id",
        "client_name",
        "campaign_name",
        "drop_date",
        "exp_date",
        "DM CASE STUDY",
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
        "LIENS CITY/COUNTY",
        "LIENS OTHER",
        "LIENS UTILITY",
        "LIENS HOA",
        "LIENS MECHANIC",
        "EVICTION",
        "30-60 DAYS",
        "JUDGEMENT",
        "DEBT COLLECTION",
        "MARKETING DM COUNT",
        "sequence_step",
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
        "Version",
        "seller_full_name",
        "seller_first_name",
        "seller_mailing_add",
        "estimated_offer_price", # Estimated offer price
        "company_name",
        "company_phone_number", # Company phone number
        "company_mailing_add",
        # "Company Mailing Address",
        "Company Mailing City",
        "Company Mailing State",
        "Company Mailing ZIP",
        "company_website",
        "test_name",
        "investor_full_name",
        "company_logo",
        "qr_code_url",
        "cred_logo_1",
        "cred_logo_2",
        "cred_logo_3",
        "cred_logo_4",
        "text_1",
        "text_2",
        "text_3",
        "text_4",
        "text_5",
        "text_6",
        "text_7",
        "text_8",
        "text_9",
        "text_10",
        "text_11",
        "font_color_1",
        "font_color_2",
        "font_color_3",
        "font_color_4",
        "block_color_1",
        "block_color_2",
        "google_street_view",
        "image",
        "postcard_size",
        "Drop #"
    ]            
    with open("results/" + client.company_name + "/FULL-" +client.company_name + ".csv", "w", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        count_drop_size = 0
        for i, postcard in enumerate(postcards_list):
            seller_mailing_add  = postcard.property_data.mailing_address + ", " + postcard.property_data.mailing_city + " " + postcard.property_data.mailing_state + ", " + postcard.property_data.mailing_zip
            company_mailing_add = client.company_mailing_address + ", " + client.company_mailing_city + " " + client.company_mailing_state + ", " + client.company_mailing_zip
            estimate_cash_offer = calculate_estimate_cash_offer(client, postcard.property_data.total_value, client.offer_price)
            if checking_test_percentage(client):  
                count_drop_size += 1 
                writer.writerow({
                    "client_id":                    client.client_id,
                    "client_name":                  client.company_name,
                    "campaign_name":                client.campaign_name,
                    "drop_date":                    client.drop_date,
                    "exp_date":                     add_30_days(client.drop_date),
                    "DM CASE STUDY":                True,
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
                    "LIENS CITY/COUNTY":            postcard.property_data.liens_city,
                    "LIENS OTHER":                  postcard.property_data.liens_other,
                    "LIENS UTILITY":                postcard.property_data.liens_utility,
                    "LIENS HOA":                    postcard.property_data.liens_hoa,
                    "LIENS MECHANIC":               postcard.property_data.liens_mechanic,
                    "EVICTION":                     postcard.property_data.eviction,
                    "30-60 DAYS":                   postcard.property_data.thirty_sixty_days,
                    "JUDGEMENT":                    postcard.property_data.judgment,
                    "DEBT COLLECTION":              postcard.property_data.debt_collection,
                    "MARKETING DM COUNT":           postcard.property_data.num_dm,
                    "sequence_step":                postcard.property_data.sequence_step,
                    "MAIN DISTRESS #1":             postcard.property_data.main_distress_1,
                    "MAIN DISTRESS #2":             postcard.property_data.main_distress_2,
                    "MAIN DISTRESS #3":             postcard.property_data.main_distress_3,
                    "MAIN DISTRESS #4":             postcard.property_data.main_distress_4,
                    "targeted_message_1":           postcard.property_data.targeted_message_1,
                    "targeted_message_2":           postcard.property_data.targeted_message_2,
                    "targeted_message_3":           postcard.property_data.targeted_message_3,
                    "targeted_message_4":           postcard.property_data.targeted_message_4,
                    "TARGETED GROUP NAME":          postcard.property_data.seller_avatar_group,
                    "targeted_test":                postcard.property_data.targeted_testimonial,
                    "Postcard Name":                "T" + postcard.postcard_number,
                    "Version":                      get_random_version(client.company_name, postcard),
                    "seller_full_name":             postcard.property_data.owner_full_name,
                    "seller_first_name":            postcard.property_data.owner_first_name,
                    "seller_mailing_add":           seller_mailing_add,
                    "estimated_offer_price":        estimate_cash_offer,
                    "company_name":                 postcard.company_name,
                    "company_phone_number":         postcard.company_phone_number,
                    "company_mailing_add":          company_mailing_add,
                    # "Company Mailing Address":      client.mailing_address, 
                    "Company Mailing City":         client.company_mailing_city,
                    "Company Mailing State":        client.company_mailing_state,
                    "Company Mailing ZIP":          client.company_mailing_zip,
                    "company_website":              postcard.company_website,
                    "test_name":                    postcard.testimonial_name,
                    "investor_full_name":           postcard.investor_full_name,
                    "company_logo":                 postcard.company_logo_url,
                    "qr_code_url":                  postcard.qr_code_url,
                    "cred_logo_1":                  postcard.cred_logo_1,
                    "cred_logo_2":                  postcard.cred_logo_2,
                    "cred_logo_3":                  postcard.cred_logo_3,
                    "cred_logo_4":                  postcard.cred_logo_4,
                    "text_1":                       get_text_by_postcard_name(client.company_name, postcard, 1, estimate_cash_offer),
                    "text_2":                       get_text_by_postcard_name(client.company_name, postcard, 2, estimate_cash_offer),
                    "text_3":                       get_text_by_postcard_name(client.company_name, postcard, 3),
                    "text_4":                       get_text_by_postcard_name(client.company_name, postcard, 4),
                    "text_5":                       get_text_by_postcard_name(client.company_name, postcard, 5),
                    "text_6":                       get_text_by_postcard_name(client.company_name, postcard, 6),
                    "text_7":                       get_text_by_postcard_name(client.company_name, postcard, 7),
                    "text_8":                       get_text_by_postcard_name(client.company_name, postcard, 8),
                    "text_9":                       get_text_by_postcard_name(client.company_name, postcard, 9),
                    "text_10":                      get_text_by_postcard_name(client.company_name, postcard, 10),
                    "text_11":                      get_text_by_postcard_name(client.company_name, postcard, 11),
                    "font_color_1":                 postcard.font_color_1,
                    "font_color_2":                 postcard.font_color_2,
                    "font_color_3":                 postcard.font_color_3,
                    "font_color_4":                 postcard.font_color_4,
                    "block_color_1":                postcard.block_color_1,
                    "block_color_2":                postcard.block_color_2,
                    "google_street_view":           postcard.google_street_view_url,
                    "image":                        postcard.bg_img,
                    "postcard_size":                client.postcard_size,
                    "Drop #":                       get_drop_number(i, len(postcards_list),client.drop_nums)
                })
            else:
                writer.writerow({
                    "client_id":                    client.client_id,
                    "client_name":                  client.company_name,
                    "campaign_name":                False,
                    "drop_date":                    False,
                    "exp_date":                     False, 
                    "DM CASE STUDY":                False,
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
                    "LIENS CITY/COUNTY":            postcard.property_data.liens_city,
                    "LIENS OTHER":                  postcard.property_data.liens_other,
                    "LIENS UTILITY":                postcard.property_data.liens_utility,
                    "LIENS HOA":                    postcard.property_data.liens_hoa,
                    "LIENS MECHANIC":               postcard.property_data.liens_mechanic,
                    "EVICTION":                     postcard.property_data.eviction,
                    "30-60 DAYS":                   postcard.property_data.thirty_sixty_days,
                    "JUDGEMENT":                    postcard.property_data.judgment,
                    "DEBT COLLECTION":              postcard.property_data.debt_collection,
                    "MARKETING DM COUNT":           postcard.property_data.num_dm,
                    "sequence_step":                postcard.property_data.sequence_step,
                    "MAIN DISTRESS #1":             postcard.property_data.main_distress_1,
                    "MAIN DISTRESS #2":             postcard.property_data.main_distress_2,
                    "MAIN DISTRESS #3":             postcard.property_data.main_distress_3,
                    "MAIN DISTRESS #4":             postcard.property_data.main_distress_4,
                    "targeted_message_1":           postcard.property_data.targeted_message_1,
                    "targeted_message_2":           postcard.property_data.targeted_message_2,
                    "targeted_message_3":           postcard.property_data.targeted_message_3,
                    "targeted_message_4":           postcard.property_data.targeted_message_4,
                    "TARGETED GROUP NAME":          postcard.property_data.seller_avatar_group                })
        print("\tDropSize: ", count_drop_size)
        
    if client.test_percentage != 100:
        df = pd.read_csv("results/" + client.company_name + "/FULL-" +client.company_name + ".csv")
            # Create two dataframes based on the "DM CASE STUDY" field
        df_true = df[df['DM CASE STUDY'] == True]
        df_false = df[df['DM CASE STUDY'] == False]
        
        # Save the dataframes to new csv files
        df_true.to_csv("results/" + client.company_name + "/CaseStudy-" +client.company_name + ".csv", index=False)
        df_false.to_csv("results/" + client.company_name + "/NoCaseStudy-" +client.company_name + ".csv", index=False)

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

def find_marketingList(company_name):
    marketing_list_folder = "input/marketingLists"
    for file_name in os.listdir(marketing_list_folder):
        if file_name.endswith(".csv") and file_name[:-4] == company_name:
            file_path = os.path.join(marketing_list_folder, file_name)
            return file_path
        elif file_name.endswith(".xlsx") and file_name[:-5] == company_name:
            file_path = os.path.join(marketing_list_folder, file_name)
            # Read Excel file
            df = pd.read_excel(file_path)
            # Save as CSV
            csv_file_path = os.path.join(marketing_list_folder, f"{company_name}.csv")
            df.to_csv(csv_file_path, index=False)
            return csv_file_path
    return False

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
            if row['Company Name'] == company_name and row['Postcard Name'] == "T" + postcard.postcard_number and postcard.version == row['Version']:
                return row[f"text_{text_number}"]
    return None

def delete_csv_files(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".csv"):
            file_path = os.path.join(folder_path, file_name)
            os.remove(file_path)

def add_30_days(drop_date: str) -> str:
    try:
        drop_date_obj = datetime.strptime(drop_date, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        try:
            drop_date_obj = datetime.strptime(drop_date, "%Y-%m-%d")
        except ValueError:
            return "Invalid date format"

    exp_date_obj = drop_date_obj + timedelta(days=30)
    exp_date = exp_date_obj.strftime("%Y-%m-%d %H:%M:%S")
    return exp_date

if __name__ == "__main__":
    excel_to_csv_start(INPUT_FILE)
    # Read clients' data from the specified CSV file
    clients = read_clients_data(INPUT_CLIENTS_DATA)
    # Iterate over each client
    for client in clients:
        if find_marketingList(client.company_name):
            print("Client:",client.company_name)
            print("\tChecking test amount:\t", client.test_percentage, "%" )
            print("\tChecking offer price:\t", int(client.offer_price*100), "%" )  
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
                # print(property_data.action_plans, property_data.score)
                # we want postcards:
                postcard_name, postcard_number, postcard_gender = get_template_for_property(property_data)
                newPostcardTemplate = PostcardsList(property_data, postcard_name, postcard_number, postcard_gender)
                newPostcardTemplate.assign_company_information(client)
                newPostcardTemplate.assign_owner_information(property_data, client) 
                newPostcardTemplate.assign_colors(client)
                newPostcardTemplate.assign_logos(client)   
                newPostcardTemplate.assign_tracking_number(client)   
                newPostcardTemplate.assign_bg_image(client)
                newPostcardTemplate.assign_google_street_view()
                postcards_list.append(newPostcardTemplate) 
            create_csv_files(postcards_list, client) 
            print("\tCompleted\n")  
        else:
            pass
    # delete_csv_files("input/marketingLists")
    delete_csv_files("input")
            # print("\tPass\n")         
  