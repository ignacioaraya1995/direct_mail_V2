import csv
from vars import *
from functions import *

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

    def assign_logos(self):
        with open(INPUT_CLIENTS_LOGOS, 'r') as file:
            reader = csv.DictReader(file)
            rows = [row for row in reader]
            
        for row in rows:
            if row['Company Name'] == self.company_name:
                if row['Type'].startswith('cred_logo'):
                    logo_number = str(row['Type'].replace("cred_logo_", ""))
                    attribute_name = f"cred_logo_{logo_number}"
                    setattr(self, attribute_name, row[f'Logo_url_T{self.postcard_number}'])
                elif row['Type'] == 'Company':
                    self.company_logo_url = row[f'Logo_url_T{self.postcard_number}']


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
    
    def assign_google_street_view(self):
        if str(self.postcard_number) == "1":
            self.google_street_view_url = True
        else:
            self.google_street_view_url = False
    