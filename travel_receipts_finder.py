from __future__ import print_function

import base64
import email
import quopri
import re
import os
import subprocess

from csv import DictWriter
from datetime import timedelta
from dateutil import parser as date_parser
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

INR_SYM = '\xe2\x82\xb9'


class OlaEmailParser(object):
    def __init__(self, parsed_mail_dict):
        self.parsed_mail_dict = parsed_mail_dict

    @staticmethod
    def get_provider():
        return "Ola"

    def get_source_dest_address(self):
        html_content = self.parsed_mail_dict["text"][1]
        addr_marker = html_content.find('class="left-space-address"',
                                        html_content.find('class="left-space-address"') + 10)
        addr_start = html_content.find('<span', html_content.find('<td', html_content.find('<td', addr_marker)))
        addr_end = html_content.find('</span>', addr_start)
        addr_sub = html_content[addr_start:addr_end]
        src_addr = addr_sub[addr_sub.rfind(">") + 1:].strip()
        addr_marker = html_content.find('class="left-space-address"', addr_end)
        addr_start = html_content.find('<span', html_content.find('<td', html_content.find('<td', addr_marker)))
        addr_end = html_content.find('</span>', addr_start)
        addr_sub = html_content[addr_start:addr_end]
        dest_addr = addr_sub[addr_sub.rfind(">") + 1:].strip()
        return src_addr, dest_addr

    def get_fare(self):
        fare_start = self.parsed_mail_dict['text'][0].find(INR_SYM)
        fare_end = self.parsed_mail_dict['text'][0].find(' ', fare_start)
        fare = float(self.parsed_mail_dict['text'][0][fare_start + len(INR_SYM):fare_end].strip())
        return fare

    def get_trip_time(self):
        trip_date = self.parsed_mail_dict['snippet'][:self.parsed_mail_dict['snippet'].find(INR_SYM)].strip()
        time_regex = '(?P<trip_time>\d{2}:\d{2}\s?(AM|PM))'
        m = re.search(time_regex, self.parsed_mail_dict['snippet'])
        if m is None:
            return None
        trip_date = '%s %s IST' % (trip_date, m.groupdict()['trip_time'])
        return date_parser.parse(trip_date)

    def save_receipt(self, file_name, save_path):
        filepath = os.path.join(save_path, file_name + '.pdf')
        ola_receipt_b64 = self.parsed_mail_dict['application'].items()[0][1]['data']
        open(filepath, 'w').write(base64.b64decode(ola_receipt_b64))


class UberMailParser(object):
    URL_TO_PDF_CMD_FORMAT = '/Users/msachdev/my_projects/email_scraper/urltopdf --url=file://{html_file} --autosave-path={save_path} --autosave-name=URL'

    def __init__(self, parsed_mail_dict):
        self.parsed_mail_dict = parsed_mail_dict
        self.address_elem_pattern = re.compile('class="address\s[\w\s]*"')
        self.html_content = self.__get_html_text_with_embedded_images(parsed_mail_dict=self.parsed_mail_dict)

    @staticmethod
    def get_provider():
        return "Uber"

    @staticmethod
    def __get_html_text_with_embedded_images(parsed_mail_dict):
        html_text = parsed_mail_dict["text"][0]
        content_map = parsed_mail_dict["images"]
        html_content = quopri.decodestring(html_text).decode('utf-8')
        pattern = "data:image/{imgtype};base64, {data}"
        for cid, img_dict in content_map.iteritems():
            html_content = html_content.replace("cid:%s" % cid,
                                                pattern.format(imgtype=img_dict["img_type"],
                                                               data=img_dict["data"]))
        return html_content

    def get_source_dest_address(self):
        # First comes the src address
        m = self.address_elem_pattern.search(self.html_content)
        address_marker = "</span>"
        address_marker_loc = self.html_content.find(address_marker, m.end()) + len(address_marker)
        src_addr = self.html_content[address_marker_loc:self.html_content.find("</td>", address_marker_loc)].strip()
        # Next comes dest address
        m = self.address_elem_pattern.search(self.html_content, m.end())
        address_marker_loc = self.html_content.find(address_marker, m.end()) + len(address_marker)
        dest_addr = self.html_content[address_marker_loc:self.html_content.find("</td>", address_marker_loc)].strip()
        return src_addr, dest_addr

    def get_fare(self):
        html_content = self.html_content.encode('utf-8')
        fare_start = html_content.find(INR_SYM)
        fare_end = html_content.find(' ', fare_start)
        fare = float(html_content[fare_start + len(INR_SYM):fare_end].strip())
        return fare

    def get_trip_time(self):
        months = ['January', 'Feburary', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
                  'November', 'December']
        uber_date_regex = '(%s)\s?\d{1,2},\s?\d{4}' % '|'.join(months)
        m = re.search('(?P<trip_date>%s)' % uber_date_regex, self.parsed_mail_dict['snippet'])
        if m is None:
            return None
        trip_date = m.groupdict()['trip_date']
        time_start = self.parsed_mail_dict['snippet'].find('|')
        time_end = self.parsed_mail_dict['snippet'].find('|', time_start + 1)
        time_regex = '(?P<trip_time>\d{2}:\d{2}\s?(AM|PM))'
        m = re.search(time_regex, self.parsed_mail_dict['snippet'][time_start + 1:time_end], re.IGNORECASE)
        if m is None:
            return None
        trip_date = '%s %s IST' % (trip_date, m.groupdict()['trip_time'])
        return date_parser.parse(trip_date)

    def save_receipt(self, file_name, save_path):
        tmp_html_file_path = os.path.join('/tmp', file_name + '.html')
        open(tmp_html_file_path, 'w').write(self.html_content.encode('utf-8'))
        cmd = self.URL_TO_PDF_CMD_FORMAT.format(html_file=tmp_html_file_path, save_path=save_path).split(' ')
        return subprocess.call(cmd) != 0


class EmailFinder(object):
    # If modifying these scopes, delete the file token.json.
    SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
    UBER_SEARCH_FORMAT = 'from:("Uber Receipts") has:attachment after:{after_date} before:{before_date}'
    OLA_SEARCH_FORMAT = 'from:("Ola") subject:ride has:attachment after:{after_date} before:{before_date}'

    def __init__(self):
        store = file.Storage('token.json')
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('credentials.json', self.SCOPES)
            creds = tools.run_flow(flow, store)
        self.service = build('gmail', 'v1', http=creds.authorize(Http()))
        self.address_elem_pattern_uber = re.compile('class="address\s[\w\s]*"')
        self.parsed_emails = list()
        self.commute_emails = list()

    @staticmethod
    def __check_address_match(trip_address, verification_keywords):
        for kwd in verification_keywords:
            if not re.search(kwd, trip_address, re.IGNORECASE):
                return False
        return True

    @staticmethod
    def __is_commute(parsed_email, home_addr_keywords_list, office_addr_keywords_list):
        trip_date = parsed_email.get_trip_time()
        if trip_date is None or trip_date.weekday() >= 5:
            return False
        start, end = parsed_email.get_source_dest_address()

        start_is_home = False
        end_is_home = False
        for home_kwds in home_addr_keywords_list:
            start_is_home = start_is_home or EmailFinder.__check_address_match(trip_address=start,
                                                                               verification_keywords=home_kwds)
            end_is_home = end_is_home or EmailFinder.__check_address_match(trip_address=end,
                                                                           verification_keywords=home_kwds)
        start_is_work = False
        end_is_work = False
        for office_kwds in office_addr_keywords_list:
            start_is_work = start_is_work or EmailFinder.__check_address_match(trip_address=start,
                                                                               verification_keywords=office_kwds)
            end_is_work = end_is_work or EmailFinder.__check_address_match(trip_address=end,
                                                                           verification_keywords=office_kwds)

        home_to_work = start_is_home and end_is_work
        work_to_home = start_is_work and end_is_home
        return home_to_work or work_to_home

    def __add_content_dictionary_to_list(self, content_dict, mime_msg):
        message_main_type = mime_msg.get_content_maintype()
        if message_main_type == 'multipart':
            for part in mime_msg.get_payload():
                self.__add_content_dictionary_to_list(content_dict=content_dict, mime_msg=part)
        elif message_main_type == 'text':
            if "text" not in content_dict:
                content_dict["text"] = list()
            content_dict["text"].append(mime_msg.get_payload())
        elif message_main_type == 'image':
            if "images" not in content_dict:
                content_dict["images"] = dict()
            cid = mime_msg["Content-Id"]
            cid = cid[1:-1]
            content_dict["images"][cid] = dict(data=mime_msg.get_payload(), img_type=mime_msg.get_content_subtype())
        elif message_main_type == 'application':
            if "application" not in content_dict:
                content_dict["application"] = dict()
            ctype = mime_msg["Content-Type"]
            fname = re.search('name="(?P<fname>[\w]+\.[\w]+)"', ctype).groupdict()['fname']
            content_dict["application"][fname] = dict(data=mime_msg.get_payload(),
                                                      img_type=mime_msg.get_content_subtype())

    def __get_parsed_content_diction(self, message):
        msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
        mime_msg = email.message_from_string(msg_str)
        content_dict = dict(snippet=message['snippet'].encode('utf-8'))
        self.__add_content_dictionary_to_list(content_dict=content_dict, mime_msg=mime_msg)
        return content_dict

    def get_parsed_message_uber(self, message):
        return UberMailParser(parsed_mail_dict=self.__get_parsed_content_diction(message=message))

    def get_parsed_message_ola(self, message):
        return OlaEmailParser(parsed_mail_dict=self.__get_parsed_content_diction(message=message))

    def process_uber_msg(self, response, exception):
        if exception is not None:
            print("Error fetching Uber email: " % str(exception))
            return
        self.parsed_emails.append((response, self.get_parsed_message_uber(message=response)))

    def process_ola_msg(self, response, exception):
        if exception is not None:
            print("Error fetching Ola email: " % str(exception))
            return
        self.parsed_emails.append((response, self.get_parsed_message_ola(message=response)))

    def fetch_all_receipts(self, start_date, end_date):
        after_date = start_date + timedelta(days=-1)
        before_date = end_date + timedelta(days=1)
        uber_query = self.UBER_SEARCH_FORMAT.format(after_date=after_date.date(), before_date=before_date.date())
        uber_mails = self.service.users().messages().list(userId='me', q=uber_query).execute()
        ola_query = self.OLA_SEARCH_FORMAT.format(after_date=after_date.date(), before_date=before_date.date())
        ola_mails = self.service.users().messages().list(userId='me', q=ola_query).execute()
        batch = self.service.new_batch_http_request()
        if len(uber_mails['messages']) + len(ola_mails['messages']) > 500:
            raise ValueError('Too many receipts to process. Please select shorter time period')
        for m in uber_mails['messages']:
            batch.add(self.service.users().messages().get(userId='me', id=m['id'], format='raw'),
                      callback=lambda _, res, exc: self.process_uber_msg(response=res, exception=exc))
        for m in ola_mails['messages']:
            batch.add(self.service.users().messages().get(userId='me', id=m['id'], format='raw'),
                      callback=lambda _, res, exc: self.process_ola_msg(response=res, exception=exc))
        batch.execute(http=Http())

    def fetch_commute_receipts(self, start_date, end_date, home_addr_keywords_list, office_addr_keywords_list):
        self.fetch_all_receipts(start_date=start_date, end_date=end_date)
        self.commute_emails = [e[1] for e in self.parsed_emails
                               if self.__is_commute(parsed_email=e[1], home_addr_keywords_list=home_addr_keywords_list,
                                                    office_addr_keywords_list=office_addr_keywords_list)]

    def save_receipts(self, dest_dir, csv_filename):
        report_data = list()
        for e in self.commute_emails:
            trip_time = e.get_trip_time()
            receipt_filename = "{trip_date}_{trip_time}hrs_{provider}".format(
                trip_date=trip_time.strftime('%Y%m%d'), trip_time=trip_time.time().strftime('%H%M'),
                provider=e.get_provider())
            e.save_receipt(file_name=receipt_filename, save_path=dest_dir)
            src, dest = e.get_source_dest_address()
            invoice_line = dict(trip_date=trip_time.strftime('%Y%m%d'),
                                trip_time="%shrs" % trip_time.time().strftime('%H%M'),
                                start_addr=src, end_addr=dest, fare=e.get_fare())
            report_data.append(invoice_line)
        with open(os.path.join(dest_dir, csv_filename), 'w') as csvfile:
            fieldnames = report_data[0].keys()
            writer = DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(report_data)


def fetch_commute_receipts(start_date_str, end_date_str, home_addr_keywords_list, office_addr_keywords_list, save_path):
    email_finder = EmailFinder()
    start_date = date_parser.parse(start_date_str)
    end_date = date_parser.parse(end_date_str)
    email_finder.fetch_commute_receipts(start_date=start_date, end_date=end_date,
                                        home_addr_keywords_list=home_addr_keywords_list,
                                        office_addr_keywords_list=office_addr_keywords_list)
    suffix = '%s_to_%s' % (start_date.strftime('%Y%m%d'), end_date.strftime('%Y%m%d'))
    dest_dir = os.path.join(save_path, suffix)
    os.makedirs(dest_dir)
    email_finder.save_receipts(dest_dir=dest_dir, csv_filename="%s.csv" % suffix)


# fetch_commute_receipts(start_date_str='2018-08-01', end_date_str='2018-09-01',
#                        home_addr_keywords=['Diamond District', 'Old Airport'],
#                        office_addr_keywords=['11th Main', 'Indiranagar'],
#                        save_path='/Users/msachdev/tmp/receipts')
