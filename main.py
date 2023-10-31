from vars import *
from functions import *
import random
from Client import Client
from MailPiece import MailPiece
from PropertyData import PropertyData

def read_clients_data(file_name):
    clients_address_data = read_csv_file(INPUT_CLIENTS_ADDRESS)
    clients_campaign_data = read_csv_file(INPUT_CLIENTS_CAMPAIGN)
    clients_offer_price_data = read_csv_file(INPUT_CLIENTS_OFFER_PRICE)
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
                row['Do you agree to share the results in a monthly basis?'],
                clients_address_data,
                clients_campaign_data,
                clients_offer_price_data
            )
            clients.append(client)
    return clients
 
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
    marketing_list = sorted(marketing_list, key=lambda x: int(x.score), reverse=True)
    return marketing_list

def create_mail_piece(property_data, client, colors_rules_data, clients_logos_data, clients_bg_img_data, force_mail_strategy):
    if force_mail_strategy:
        sequence_step = (property_data.num_dm % 4) + 1
        property_data.sequence_step = str(sequence_step)
        postcard_name = ""
        postcard_number = force_mail_strategy[1]
        postcard_gender = random.choice(["Male", "Female"])
        mail_type = force_mail_strategy[0]
    if force_mail_strategy == False:
        postcard_name, postcard_number, postcard_gender, mail_type = get_template_for_property(property_data)
        
        
    postcard = MailPiece(mail_type, property_data, postcard_name, postcard_number, postcard_gender)
    postcard.assign_company_information(client)
    postcard.assign_owner_information(property_data)
    postcard.assign_colors(colors_rules_data)
    postcard.assign_logos(clients_logos_data)
    postcard.assign_tracking_number(client)
    postcard.assign_bg_image(clients_bg_img_data)
    postcard.assign_google_street_view()
    return postcard

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
            estimate_cash_offer = calculate_estimate_cash_offer(postcard.property_data.total_value, client.offer_price)          
            if postcard.mail_type == "Postcard":
                mail_type_id = "T" + str(postcard.postcard_number)
                mail_version = get_random_version(client.company_name, postcard)            

            
            if postcard.mail_type == "CheckLetter":
                mail_type_id = "CL" + str(postcard.postcard_number)
                mail_version = "a"          

                
            
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
                    "Postcard Name":                mail_type_id,
                    "Version":                      mail_version,
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
        print("DropSize: ", count_drop_size)
        
    if client.test_percentage != 100:
        df = pd.read_csv("results/" + client.company_name + "/FULL-" +client.company_name + ".csv")
            # Create two dataframes based on the "DM CASE STUDY" field
        df_true = df[df['DM CASE STUDY'] == True]
        df_false = df[df['DM CASE STUDY'] == False]
        
        # Save the dataframes to new csv files
        df_true.to_csv("results/" + client.company_name + "/CaseStudy-" +client.company_name + ".csv", index=False)
        df_false.to_csv("results/" + client.company_name + "/NoCaseStudy-" +client.company_name + ".csv", index=False)

def delete_csv_files(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".csv"):
            file_path = os.path.join(folder_path, file_name)
            os.remove(file_path)

if __name__ == "__main__":
    excel_to_csv_start(INPUT_FILE)
    clients = read_clients_data(INPUT_CLIENTS_DATA)
    colors_rules_data = read_csv_file(INPUT_COLORS_RULES)
    clients_logos_data = read_csv_file(INPUT_CLIENTS_LOGOS)
    clients_bg_img_data = read_csv_file(INPUT_CLIENTS_BG_IMG)
        
    for client in clients:
        if not find_marketingList(client.company_name):
            continue
        
        print(f"Client: {client.company_name}\tTest Amount: {client.test_percentage}%\tOffer Price/Range: {client.offer_price}")     
        create_client_folder(client)
        client.add_marketing_list(read_marketing_list_csv(find_marketingList(client.company_name)))
        mail_list = []
        FORCE_STRATEGY = True
        
        if FORCE_STRATEGY:
            counters = {"CheckLetter": 0, "Postcard (Google Streetview)": 0, "Postcard": 0}
            limits = [("CheckLetter", 9999999), ("Postcard (Google Streetview)", 0), ("Postcard", 0)]
            for mail_type, limit in limits:
                for property_data in client.marketing_list[counters[mail_type]:]:
                    if counters[mail_type] >= limit:
                        break
                    if mail_type == "CheckLetter":
                        force_mail_strategy = ["CheckLetter", random.choice([1, 2])]
                        mail_list.append(create_mail_piece(property_data, client, colors_rules_data, clients_logos_data, clients_bg_img_data, force_mail_strategy))
                    elif mail_type == "Postcard (Google Streetview)":
                        force_mail_strategy = ["Postcard", 1]
                        mail_list.append(create_mail_piece(property_data, client, colors_rules_data, clients_logos_data, clients_bg_img_data, force_mail_strategy))
                    elif mail_type == "Postcard":
                        force_mail_strategy = ["Postcard", random.choice([3,4])]
                        mail_list.append(create_mail_piece(property_data, client, colors_rules_data, clients_logos_data, clients_bg_img_data, force_mail_strategy))
                    counters[mail_type] += 1
                    
        if FORCE_STRATEGY == False:
            for property_data in client.marketing_list:
                mail_list.append(create_mail_piece(property_data, client, colors_rules_data, clients_logos_data, clients_bg_img_data, FORCE_STRATEGY))
            
        
            
        calculate_cost(mail_list)
        create_csv_files(mail_list, client)
         
    print("\tCompleted\n")
    delete_csv_files("input")

