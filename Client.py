import csv
from vars import *

class Client:
    def __init__(self, client_id, company_name, contact_name, contact_email, contact_phone, mailing_address, website, logo, tracking_numbers, demographic, postcard_quantity, test_percentage, drop_date, postcard_size, featured_in_tv, bbb_accreditation, years_in_business, agent_name, website_link, response_rate, roi, deal_postcards, mail_house, postcard_designs, additional_comments, share_results):
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
                self.campaign_name = ""
                self.drop_nums = "Not found"
    
    def get_offer_price(self):
        with open(INPUT_CLIENTS_OFFER_PRICE, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Company Name'] == self.company_name:
                    self.offer_price = row['Price Offer Rate']
                    break
            else:
                self.offer_price = .85 
    
    @staticmethod
    def center_text(text, width):
        return text.center(width)
