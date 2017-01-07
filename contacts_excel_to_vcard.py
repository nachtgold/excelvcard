#!/usr/bin/python

import vobject
from openpyxl import load_workbook
import re
import os
from datetime import datetime

wb = load_workbook(filename='Kontakte.xlsx', read_only=True)
sheet = wb.active

for row in sheet.iter_rows(min_row=2):
    # assume that the row could be empty
    valid_row = False

    # check all cells of the row if a single value exists
    for cell in row:
        if cell.value is not None:
            valid_row = True

    # only Rows with a single value are interesting
    if valid_row:
        full_name = row[0].value
        full_address = row[1].value
        mobile_phone = row[2].value
        home_phone = row[3].value
        work_phone = row[4].value
        email = row[5].value
        birthday = row[6].value
        notes = row[7].value

        given_name = ''
        family_name = ''
        city_name = ''
        street_name = ''
        zip_code = ''

        # split complex values
        if full_name is not None:
            if ',' in full_name:
                parts = full_name.split(',')
                family_name = parts[0].strip()
                given_name = parts[1].strip()
            else:
                given_name = full_name
        else:
            full_name = ''

        if full_address is not None:
            if ',' in full_address:
                parts = full_address.split(',')
                street_name = parts[0].strip()
                city_name = parts[1].strip()

                if ' ' in city_name:
                    parts = city_name.split(' ')
                    zip_code = parts[0].strip()
                    city_name = parts[1].strip()
            else:
                city_name = full_address

        # fix phone numbers
        if mobile_phone is not None:
            mobile_phone = re.sub("[^0-9]", "", mobile_phone)
            if mobile_phone.startswith('00'):
                mobile_phone = '+' + mobile_phone[2:]
            if mobile_phone.startswith('0'):
                mobile_phone = '+49' + mobile_phone[1:]
        if home_phone is not None:
            home_phone = re.sub("[^0-9]", "", home_phone)
            if home_phone.startswith('00'):
                home_phone = '+' + home_phone[2:]
            if home_phone.startswith('0'):
                home_phone = '+49' + home_phone[1:]
        if work_phone is not None:
            work_phone = re.sub("[^0-9]", "", work_phone)
            if work_phone.startswith('00'):
                work_phone = '+' + work_phone[2:]
            if work_phone.startswith('0'):
                work_phone = '+49' + work_phone[1:]

        # Create new contact
        new_contact = vobject.vCard()

        new_contact.add('n').value = vobject.vcard.Name(family=family_name, given=given_name)
        new_contact.add('fn').value = full_name

        if email is not None:
            new_contact.add('email').value = email
            new_contact.email.type_param = 'INTERNET'

        new_contact.add('adr').value = vobject.vcard.Address(street=street_name, city=city_name, code=zip_code)
        new_contact.adr.type_param = 'HOME'

        if mobile_phone is not None:
            mobile = new_contact.add('tel')
            mobile.value = mobile_phone
            mobile.type_param = 'CELL'

        if home_phone is not None:
            home = new_contact.add('tel')
            home.value = home_phone
            home.type_param = 'HOME'

        if work_phone is not None:
            work = new_contact.add('tel')
            work.value = work_phone
            work.type_param = 'WORK'

        if birthday is not None:
            if isinstance(birthday, datetime):
                new_contact.add('bday').value = birthday.strftime('%Y-%m-%d')
            else:
                new_contact.add('bday').value = datetime.strptime(birthday, '%d.%m.').strftime('--%m-%d')

        if notes is not None:
            new_contact.add('note').value = notes

        # write the card
        if not os.path.exists('contacts'):
            os.makedirs('contacts')
        f = open('contacts/' + re.sub("[^a-zA-Z]", "", full_name) + '.vcf', 'w')
        f.write(new_contact.serialize())
        f.close()
