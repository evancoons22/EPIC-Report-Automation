# import email
# from pythonautomation.report import build_report
import report
import pandas as pd

portfolio = pd.read_excel("./data/Epic Portfolio.xlsx")
reporting_list = pd.read_excel("./data/reportinglist1.xlsx")
commitments_and_transactions = pd.read_excel("./data/commitmentstransactions1.xlsx")

email_addresses= ['aaaaaaaaaaa', 'bbbbbbbbbbbbb', 'ccccccccccccc', 'dddddddddddd']

count = 1
for address in email_addresses: 
    report.build_report(address, portfolio, reporting_list, commitments_and_transactions, count)
    count +=1
