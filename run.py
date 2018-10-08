import argparse
import os
from travel_receipts_finder import fetch_commute_receipts

_usage = """Tool to download commute receipts from GMAIL"""


class StoreDataFromFile(argparse.Action):
    def __call__(self, parser, namespace, values, option_string=None):
        if not os.path.exists(values):
            raise ValueError("File not foud: %s" % values)
        curr_val = getattr(namespace, self.dest)
        if curr_val is None:
          curr_val = list()
        with open(values, 'r') as f:
            for line in f:
                line = line.strip()
                kwds = line.split(',')
                curr_val.append(kwds)
        setattr(namespace, self.dest, curr_val)


def main(args):
    fetch_commute_receipts(start_date_str=args.from_date, end_date_str=args.to_date,
                           home_addr_keywords_list=args.home_addr_kwds_list,
                           office_addr_keywords_list=args.office_addr_kwds_list, save_path=args.output_dir)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=_usage)
    parser.add_argument('--from_date', help="""Starting date for fetch in format YYYY-MM-DD (e.g. 2018-09-01)""")
    parser.add_argument('--to_date', help="""Last date for fetch in format YYYY-MM-DD (e.g. 2018-10-01)""")
    parser.add_argument('--home_address_file', dest='home_addr_kwds_list', action=StoreDataFromFile,
                        help="""File containing list of keywords. The format is such that each line is an alternate
                        to be considered (i.e. OR condition) and from each line, comma-separated strings will be 
                        looked up as-in as an AND operator""")
    parser.add_argument('--office_address_file', dest='office_addr_kwds_list', action=StoreDataFromFile,
                        help="""File containing list of keywords. The format is such that each line is an alternate
                        to be considered (i.e. OR condition) and from each line, comma-separated strings will be 
                        looked up as-in as an AND operator""")
    parser.add_argument('--output_dir', dest='output_dir',
                        help="""Target directory to store the PDF receipts and the CSV file""")
    main(args=parser.parse_args())


