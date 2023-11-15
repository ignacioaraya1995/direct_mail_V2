import re
from vars import *

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
        self.owner_full_name = owner_full_name.lower().title().replace('Llc', 'LLC')
        self.owner_first_name = owner_first_name.lower().title().replace('Llc', 'LLC')
        self.owner_last_name = owner_last_name.lower().title().replace('Llc', 'LLC')
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
        self.clean_owner_info()
        self.check_golden_addresses()

    def check_golden_addresses(self):
        if self.golden_address != "" and self.golden_address != self.mailing_address:
            self.mailing_address    = self.golden_address
            self.mailing_city       = self.golden_city
            self.mailing_state      = self.golden_state
            self.mailing_zip        = self.golden_zip_code

    def _remove_single_letters_and_titles(self, full_name):
        return ' '.join(word for word in full_name.split() if len(word) > 1 and word.lower() not in ['jr', 'sr'])

    def _split_name(self, full_name):
        cleaned_name = re.sub(r'[^\w\s]', '', full_name)
        name_parts = cleaned_name.split()
        return ' '.join(name_parts[:-1]), name_parts[-1] if len(name_parts) > 1 else ''

    def _is_valid_name(self, name):
        return name and bool(re.match('^[a-zA-Z\s]+$', name)) and len(name) > 1

    def _correct_name_order(self, first_name, last_name):
        # If either name part is more common as a first name, swap them
        if first_name.title() in common_first_names and last_name.title() not in common_first_names:
            return first_name, last_name
        elif last_name.title() in common_first_names and first_name.title() not in common_first_names:
            return last_name, first_name
        return first_name, last_name

    def clean_owner_info(self):
        self.owner_full_name = str(self.owner_full_name).strip(" ,-.")
        
        # If "llc" is in the name, convert it to "LLC"
        if 'llc' or "trust" or "investment" or "properties" or "prop" or "capital" or "acquisitions" in self.owner_full_name.lower():
            self.owner_full_name = self.owner_full_name.lower().title().replace('Llc', 'LLC')
            self.owner_first_name = ""
            self.owner_last_name = ""
            return

        # Remove single-letter words (like middle initials) and 'Jr' and 'Sr' from full_name
        cleaned_full_name = self._remove_single_letters_and_titles(self.owner_full_name)
        first_name_candidate, last_name_candidate = self._split_name(cleaned_full_name)
        
        # If the pre-existing first name is not valid, use the split first name candidate
        if not self._is_valid_name(self.owner_first_name):
            if self._is_valid_name(first_name_candidate):
                self.owner_first_name = first_name_candidate
            else:
                self.owner_first_name = self.owner_full_name
        
        # Reset last_name to avoid repetition when first_name gets updated
        self.owner_last_name = None

        # If the pre-existing last name is not valid, use the split last name candidate
        if not self._is_valid_name(self.owner_last_name):
            if self._is_valid_name(last_name_candidate):
                self.owner_last_name = last_name_candidate
            else:
                self.owner_last_name = self.owner_full_name

        # Correct the name order using common first names list
        self.owner_first_name, self.owner_last_name = self._correct_name_order(self.owner_first_name, self.owner_last_name)

