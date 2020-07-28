#!/usr/bin/env python3

from xlrd import open_workbook, XLRDError, xldate_as_tuple
import datetime
import bgtunnel
import MySQLdb
import re
from emailer import *
from dotenv import load_dotenv

load_dotenv()
cutoff_year = '14'

# setup data validationhttp://imgur.com/bDcs1dd
dealerships = ('3RIVERS', 
    'ALASKA FRONTIER', 
    'AVATAA',
    'BOAT COUNTRY', 
    'CLEMENS EUGENE', 
    'CLEMENS MARINA', 
    'ELEPHANT BOYS',
    'IDAHO MARINE', 
    'MAXXUM MARINE', 
    'PRINCE GEORGE MOTORSPORTS',
    'PORT BOAT HOUSE', 
    'RIVERFRONT MARINA', 
    'THE BAY COMPANY',
    'VALLEY MARINE',
    'VOLGA BOAT LLC', 
    'Y-MARINA'
)

boat_models = ("18' SEAHAWK",
    "20' CASCADE", "20' OSPREY", "20' SCOUT",  "20' SEAHAWK",
    "21' OSPREY", "21' SCOUT", "21' SEAHAWK", "21'6 COMMANDER",
    "22' COMMANDER", "22' OSPREY", "22' SEAHAWK", "22' SEAHAWK HT", "22' SEAHAWK INBOARD", "22' SCOUT",
    "23' COMMANDER", "23' OSPREY", "23' SEAHAWK", "23' SEAHAWK HT", "23' SEAHAWK INBOARD", "23' SEAHAWK OS", "23' SCOUT",
    "24' COMMANDER", "24' OSPREY", "24' SEAHAWK", "24' SEAHAWK HT", "24' SEAHAWK INBOARD", "24' SCOUT",
    "25' COMMANDER","25' OSPREY", "25' SEAHAWK", "25' SEAHAWK HT", "25' SEAHAWK INBOARD", "25' SCOUT", "25' SEAHAWK OS", "25'6 SEAHAWK CUDDY",
    "27' SEAHAWK OS",
    "29' SEAHAWK OS", "29' SEAHAWK OSWA",
    "31' SEAHAWK OS", "31' SEAHAWK OSWA",
    "33' VOYAGER",
    "35' VOYAGER",
)

pattern = "^NRB (18|19|20|21|22|23|24|25|26|27|28|29|30|31|32|33|35)\d{3} [A-L](212|213|313|314|414|415|515|516|616|617|717|718)$";

xlsfile = os.getenv('XLSFILE')

output = ""


#### HEAR BE DRAGONS
def readsheet(xlsfile):
    # Read boat/dealer/model from spreadsheet
    book = open_workbook(xlsfile)
    sh = book.sheet_by_index(0)

    duplicate_guard = []
    xlshulls = []
    errors_dealer = []
    errors_boat_model = []
    errors_hull = []

    for rx in range(sh.nrows):
        hull, last_name, first_name, phone, \
            mailing_address, mailing_city, mailing_state, mailing_zip, \
            street_address, street_city, street_state, street_zip, \
            date_purchased, dealer, boat_model, p = [x.value for x in sh.row_slice(rx,0 , 16)]

        # print(rx, hull,  p)
        if (hull in duplicate_guard):
            mail_results(
                 'Registrations and Dealer Inventory Sheet Duplictate', 
                 '<p>Hull ' + hull + ' is a duplicate\n</p>'
            )
            print('Hull ' + hull + 'is a duplicate')
            continue
        else:
            duplicate_guard.append(hull)
        if (hull[:3] == 'NRB'):
            if p:
                p = "%4d-%02d-%02d" % xldate_as_tuple(p, book.datemode)[:3]
            else:
                p = None 
            if date_purchased:
                date_purchased = "%4d-%02d-%02d" % xldate_as_tuple(date_purchased, book.datemode)[:3]
            else:
                date_purchased = None 
            flag  = (not(dealer in dealerships)) * 1 + (not(boat_model in boat_models)) * 2
            if (flag & 1 and (hull[-2:] > cutoff_year)):
                errors_dealer.append([hull, dealer, boat_model])
            if (flag & 2 and (hull[-2:] > cutoff_year)): # do not verify model on older boats
                errors_boat_model.append([hull, dealer, boat_model])
            if (re.match(pattern,hull)):
               flag = 1
               errors_hull.append([hull, dealer, boat_model])
            if (flag == 0):
                boat_model = boat_model.replace("CASCADE","Cascade")
                boat_model = boat_model.replace("COMMANDER","Commander")
                boat_model = boat_model.replace("OSPREY","Osprey")
                boat_model = boat_model.replace("SCOUT","Scout")
                boat_model = boat_model.replace("SEAHAWK CUDDY", "Seahawk Cuddy")
                boat_model = boat_model.replace("SEAHAWK HT", "Seahawk Hardtop")
                boat_model = boat_model.replace("SEAHAWK INBOARD", "Seahawk Inboard")
                boat_model = boat_model.replace("SEAHAWK","Seahawk")
                boat_model = boat_model.replace("VOYAGER PILOT HOUSE","Voyager Pilot House")
                boat_model = boat_model.replace("VOYAGER WALK AROUND","Voyager Walk Around")
                xlshulls.append([hull[:3] + ' ' + hull[3:8] + ' ' + hull[8:], dealer.title(), boat_model, 
                    last_name, first_name, phone,
                    mailing_address, mailing_city, mailing_state, mailing_zip,
                    street_address, street_city, street_state, street_zip,
                    date_purchased, p])
    del sh
    del book
    return xlshulls, errors_dealer, errors_boat_model, errors_hull

def format_hull_errors(errors_hull):
    output = ""
    if (len(errors_hull)):
        output += """<span style="font-size:2em;">Hull Errors</span>
    <hr style="margin-left: 0; width: 40em;">
    <table style="border-collapse: collapse;width: 40em;">
      <tr>
        <th style="text-align: left; padding: 8px;">Serial Number</th>
        <th style="text-align: left; padding: 8px;">Model</th>
        <th style="text-align: left; padding: 8px;">Dealership</th>
      </tr>\n"""
        row = True
        td = '<td style="text-align: lefty padding: 8px;">'
        for item in sorted(errors_hull):
            # if int(item[0][12:14]) > cutoff_year:
            if row :
               tr = '<tr style="background-color: #e2e2e2;">'
            else:
               tr = '<tr>'
            row = not row
            output += """  %s
        %s%s</td>
        %s%s</td>
        %s%s</td>
      </tr>\n""" %  (tr,td,item[0],td,item[1],td,item[2]) 
        output += """</table>\n<p>&nbsp;</p>\n<p>&nbsp;</p>"""
    return output

def format_dealer_errors(errors_dealer):
    output = ""
    if (len(errors_dealer)):
        output += """<span style="font-size:2em;">Dealer Errors</span>
    <hr style="margin-left: 0; width: 40em;">
    <table style="border-collapse: collapse;width: 40em;">
      <tr>
        <th style="text-align: left; padding: 8px;">Serial Number</th>
        <th style="text-align: left; padding: 8px;">Model</th>
        <th style="text-align: left; padding: 8px;">Dealership</th>
      </tr>\n"""
        row = True
        td = '<td style="text-align: lefty padding: 8px;">'
        for item in sorted(errors_dealer):
            # if int(item[0][12:14]) > cutoff_year:
            if row :
               tr = '<tr style="background-color: #e2e2e2;">'
            else:
               tr = '<tr>'
            row = not row
            output += """  %s
        %s%s</td>
        %s%s</td>
        %s%s</td>
      </tr>\n""" %  (tr,td,item[0],td,item[1],td,item[2]) 
        output += """</table>\n<p>&nbsp;</p>\n<p>&nbsp;</p>"""
    return output

def format_boat_model_errors(errors_boat_model):
    output = ""
    if (len(errors_boat_model)>0):
        output += """<span style="font-size:2em;">Model Errors</span>
    <hr style="margin-left: 0; width: 40em;">
    <table style="border-collapse: collapse;width: 40em;">
      <tr>
        <th style="text-align: left; padding: 8px;">Serial Number</th>
        <th style="text-align: left; padding: 8px;">Model</th>
        <th style="text-align: left; padding: 8px;">Dealership</th>
      </tr>\n"""
        row = True
        td = '<td style="text-align: lefty padding: 8px;">'
        for item in sorted(errors_boat_model):
            # if int(item[0][12:14]) > cutoff_year:
            if row :
               tr = '<tr style="background-color: #e2e2e2;">'
            else:
               tr = '<tr>'
            row = not row
            output += """  %s
        %s%s</td>
        %s%s</td>
        %s%s</td>
      </tr>\n""" %  (tr,td,item[0],td,item[1],td,item[2]) 
        output += """</table>\n<p>&nbsp;</p>\n<p>&nbsp;</p>"""
    return output

def format_errors(errors_dealer, errors_boat_model, errors_hull):
    return format_hull_errors(errors_hull) + format_dealer_errors(errors_dealer) + format_boat_model_errors(errors_boat_model)

def push_sheet(xlshulls):
    forwarder = bgtunnel.open(ssh_user=os.getenv('SSH_USER'), ssh_address=os.getenv('SSH_HOST'), host_port=3306, bind_port=3307)
    #forwarder = bgtunnel.open(ssh_user=os.getenv('SSH_USER'), ssh_address='10.10.200.93', host_port=3306, bind_port=3306)

    conn= MySQLdb.connect(host='127.0.0.1', port=3307, user=os.getenv('DB_USER'), passwd=os.getenv('DB_PASS'), db=os.getenv('DB_NAME'))

    conn.query("""TRUNCATE TABLE wp_nrb_hulls""")
    r=conn.use_result()

    cursor = conn.cursor()
    sql = """
    INSERT INTO wp_nrb_hulls (
        hull_serial_number, dealership, model,
        last_name, first_name, phone,
        mailing_address, mailing_city, mailing_state, mailing_zip,
        street_address, street_city, street_state, street_zip,
        date_purchased, p
    ) VALUES (
        %s, %s, %s, 
        %s, %s, %s, 
        %s, %s, %s, %s, 
        %s, %s, %s, %s, 
        %s, %s
    )"""
    cursor.executemany(sql,sorted(xlshulls))
    conn.commit()

    cursor.close()
    conn.close()
    forwarder.close()


def mail_results(subject, body):
    mFrom = os.getenv('MAIL_FROM')
    mTo = os.getenv('MAIL_TO')
    m = Email(os.getenv('MAIL_SERVER'))
    m.setFrom(mFrom)
    for email in mTo.split(','):
        m.addRecipient(email)
    m.addCC(os.getenv('MAIL_FROM'))
   
    m.setSubject(subject)
    m.setTextBody("You should not see this text in a MIME aware reader")
    m.setHtmlBody(body)
    #  m.send()

def main():
    try:
        xlshulls, errors_dealer, errors_boat_model, errors_hull = readsheet(xlsfile)
        push_sheet(xlshulls)
        if (errors_dealer or errors_boat_model or errors_hull):
            body = format_errors(errors_dealer, errors_boat_model, errors_hull)
            mail_results('Dealer Inventory Data Entry Errors', body)
    except OSError:
        mail_results(
            'Registrations and Dealer Inventory Sheet is Open',
            'Registrations and Dealer Inventory Sheet is Open, website can not be updated'
        )
    except Exception as e:
        mail_results(
            'Registrations and Dealer Inventory Sheet Processing Error',
            '<p>Website can not be updated due to error on sheet:<br />\n' + e + '</p>'
        )

if __name__ == "__main__":
    main()
