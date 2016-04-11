#!/usr/bin/python

from xlrd import open_workbook, XLRDError
import os
# import csv
from twisted.web.client import downloadPage
from twisted.internet import reactor


def main():
    doc_num = 1103
    doc_max = 1264
    for no in range(doc_num, doc_max + 1):
        xls_url = "http://web3.moeaboe.gov.tw/ECW/populace/content/\
wHandStatistics_File.ashx?statistics_id={}&serial_no=2".format(no)
        dl_path = "/tmp/{}.xls".format(no)
        downloadPage(xls_url, dl_path).addBoth(stop)
        reactor.run()
        excel_to_csv(dl_path)


def stop():
    print("tttest!!!")
    reactor.stop()


def excel_to_csv(xls):
    # csv_path = xls.replace("xls", "csv")
    try:
        wb = open_workbook(xls)
    except XLRDError:
        os.remove(xls)
        return None
    sh = wb.sheet_by_index(0)
    # csv_h = open(csv_path, 'wb')
    # wr = csv.writer(csv_h, quoting=csv.QUOTE_ALL)

    for rownum in xrange(sh.nrows):
        # wr.writerow(sh.row_values(rownum))
        print [x.encode('utf-8') if isinstance(x, type(u'')) else x
               for x in sh.row_values(rownum)]

    # csv_h.close()


def handle_error(error):
    print("an error occurred: {}".format(error))
    reactor.stop()

if __name__ == "__main__":
    excel_to_csv("/tmp/data.xls")
    # main()
