#!/usr/bin/env python3

import bgtunnel
import click
import datetime
from dotenv import load_dotenv
from emailer import Email
from hashlib import md5
from mysql_tunnel import TunnelSQL
import MySQLdb
import os
import re
import sys
import traceback
from xlrd import open_workbook, XLRDError, xldate_as_tuple
import xlwt
from xlutils.copy import copy

cutoff_year = '14'

# setup data validationhttp://imgur.com/bDcs1dd
dealerships = (
    '3RIVERS',
    'ALASKA FRONTIER',
    'AVATAA',
    'BOAT COUNTRY',
    'CLEMENS EUGENE',
    'CLEMENS MARINA',
    'DRUMMONDVILLE MARINE',
    'ELEPHANT BOYS',
    'ENNS BROTHERS',
    'ERIE MARINE',
    'IDAHO MARINE',
    'HAMPTON MARINE',
    'MAXXUM MARINE',
    'PRINCE GEORGE MOTORSPORTS',
    'PORT BOAT HOUSE',
    'RIVERFRONT MARINA',
    'TEST DEALERSHIP',
    'THE BAY COMPANY',
    'VALLEY MARINE',
    'VOLGA BOAT LLC',
    'Y-MARINA'
)

boat_models = (
    # 18
    "18' SEAHAWK",
    # 20
    "20' CASCADE",
    "20' OSPREY",
    "20' PURSUIT",
    "20' SCOUT",
    "20' SEAHAWK",
    # 21
    "21' OSPREY",
    "21' PURSUIT",
    "21' SCOUT",
    "21' SEAHAWK FB",
    "21' SEAHAWK",
    "21'6 COMMANDER",
    # 22
    "22' COMMANDER",
    "22' OSPREY",
    "22' PURSUIT",
    "22' SCOUT",
    "22' SEAHAWK FB",
    "22' SEAHAWK HT",
    "22' SEAHAWK INBOARD",
    "22' SEAHAWK",
    # 23
    "23' COMMANDER",
    "23' OSPREY",
    "23' PURSUIT",
    "23' SCOUT",
    "23' SEAHAWK FB",
    "23' SEAHAWK HT",
    "23' SEAHAWK INBOARD",
    "23' SEAHAWK OS",
    "23' SEAHAWK",
    # 24
    "24' COMMANDER",
    "24' OSPREY",
    "24' SCOUT",
    "24' SEAHAWK FB",
    "24' SEAHAWK HT",
    "24' SEAHAWK INBOARD",
    "24' SEAHAWK",
    # 25
    "25' COMMANDER",
    "25' OSPREY",
    "25' SCOUT",
    "25' SEAHAWK FB",
    "25' SEAHAWK HT",
    "25' SEAHAWK INBOARD",
    "25' SEAHAWK OS",
    "25' SEAHAWK",
    "25'6 SEAHAWK CUDDY",
    # 26
    "26' OSPREY",
    # 27
    "27' SEAHAWK OS",
    # 29
    "29' SEAHAWK OS",
    "29' SEAHAWK OSWA",
    # 31
    "31' SEAHAWK OS",
    "31' SEAHAWK OSWA",
    # 33
    "33' SEAHAWK OS",
    "33' SEAHAWK OSWA",
    "33' VOYAGER",
    # 35
    "35' SEAHAWK OS",
    "35' SEAHAWK OSWA",
    "35' VOYAGER",
    # 37
    "37' VOYAGER",
)

pattern = "^NRB (18|19|20|21|22|23|24|25|26|27|28|29|30|31|32|33|35)\d{3} [A-L]"
pattern += "(212|213|313|314|414|415|515|516|616|617|717|718|818|819|919|920|"
pattern += "020|021|121|122|222|223|323|324|424|425|525|526|626|627|727|728|828|829|929|930"
pattern += "030|031|131|132|232|233|333|334|434|435|535|536|636|637|737|738|838|839|939|940)$"
xlsfile = ""
verbosity = ""

"""
Levels
0 = no verbosity
1 = minimal verbosity
2 = verbose outupt
3 = very verbose outupt
"""
dbgs = 0
verbosity = 0
def dbg(level, text):
    if verbosity > (level -1):
        print(text)

#### HEAR BE DRAGONS
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def readsheet(xlsfile):
    # Read boat/dealer/model from spreadsheet
    print(xlsfile)
    book = open_workbook(xlsfile, formatting_info=True)
    sh = book.sheet_by_index(0)
    wb = copy(book)             # to write to file  wb.save('filename')
    ws = wb.get_sheet(0)        # write-only copy   ws.write(row,col,'value'

    font_size_style = xlwt.easyxf('font: name Garmond, bold off, height 240;')
    nulls = 0
    duplicate_guard = []
    xlshulls = []
    errors_dealer = []
    errors_boat_model = []
    errors_hull = []
    changed = False  # write file back out if PIN is written

    dbg(2, f"Rows: {sh.nrows}")
    for rx in range(sh.nrows):
        hull, last_name, first_name, phone, \
            mailing_address, mailing_city, mailing_state, mailing_zip, \
            street_address, street_city, street_state, street_zip, \
            email_address, \
            date_purchased, dealer, boat_model, date_delivered, \
            date_finished, pin, opr, css = [x.value for x in sh.row_slice(rx,0, 21)]

        dbg(3, "{}\t{}\t{}\t{}\t{}".format(rx, hull, pin, date_finished, date_delivered))
        dbg(3, "Date Finished: {}".format(date_finished))
        # bail after 6 non hull rows, header row counts as non hull
        if (hull[:3] != 'NRB'):
            nulls += 1
            if nulls > 6:
                break
            else:
                continue

        # deal with duplicate hull numbers
        if (hull in duplicate_guard):
            mail_results(
                 'Registrations and Dealer Inventory Sheet Duplictate',
                 '<p>Hull ' + hull + ' is a duplicate\n</p>'
            )
            dbg(1, "Hull {} is a duplicate".format(hull))
            continue
        else:
            duplicate_guard.append(hull)

        # pin
        if not pin:
            changed = True
            pin = "{:04.0f}".format(int(md5(hull.encode()).hexdigest()[:9],16)%10000)
            ws.write(rx, 18, pin, font_size_style)

        #clean up dates
        if date_delivered:
            date_delivered = "%4d-%02d-%02d" % xldate_as_tuple(date_delivered, book.datemode)[:3]
        else:
            date_delivered = None
        if date_finished:
            date_finished = "%4d-%02d-%02d" % xldate_as_tuple(date_finished, book.datemode)[:3]
        else:
            date_finished = None
        if date_purchased:
            date_purchased = "%4d-%02d-%02d" % xldate_as_tuple(date_purchased, book.datemode)[:3]
        else:
            date_purchased = None

        # deal with invalid dealer, boat model, or hull number
        # flags 1=invalid dealer 2=invalid boat model 4=invalid hull number
        flag  = (not(dealer in dealerships)) * 1 + (not(boat_model in boat_models)) * 2
        if (flag & 1 and (hull[-2:] > cutoff_year)):
            dbg(1,f"Dealer Error: {hull[:3]} {hull[3:8]} {hull[8:]}  {dealer:25.25}  {boat_model}")
            errors_dealer.append([hull, dealer, boat_model])
        if (flag & 2 and (hull[-2:] > cutoff_year)): # do not verify model on older boats
            dbg(1,f" Model Error: {hull[:3]} {hull[3:8]} {hull[8:]}  {dealer:25.25}  {boat_model}")
            errors_boat_model.append([hull, dealer, boat_model])
        if (re.match(pattern,hull)):
            dbg(1,f"  Hull Error: {hull[:3]} {hull[3:8]} {hull[8:]}  {dealer:25.25}  {boat_model}")
            flag = 4
            errors_hull.append([hull, dealer, boat_model])
        # if we had any errors, loop
        if flag:
            continue
        # process hull and adujust model names 
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
            email_address,
            date_purchased, date_delivered, date_finished, pin, opr, css])
    if (changed and not dbgs):
        try:
            wb.save(xlsfile)
            dbg(1, 'Registrations and Dealer Inventory Sheet Saved')
        except OSError:
            pass
    del sh
    del book
    del wb
    return xlshulls, errors_dealer, errors_boat_model, errors_hull


def format_hull_errors(errors_hull):
    verbosity = ""
    if (len(errors_hull)):
        verbosity += """<span style="font-size:2em;">Hull Errors</span>
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
            verbosity += """  %s
        %s%s</td>
        %s%s</td>
        %s%s</td>
      </tr>\n""" %  (tr,td,item[0],td,item[1],td,item[2])
        verbosity += """</table>\n<p>&nbsp;</p>\n<p>&nbsp;</p>"""
    return verbosity


def format_dealer_errors(errors_dealer):
    verbosity = ""
    if (len(errors_dealer)):
        verbosity += """<span style="font-size:2em;">Dealer Errors</span>
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
            verbosity += """  %s
        %s%s</td>
        %s%s</td>
        %s%s</td>
      </tr>\n""" %  (tr,td,item[0],td,item[2],td,item[1])
        verbosity += """</table>\n<p>&nbsp;</p>\n<p>&nbsp;</p>"""
    return verbosity


def format_boat_model_errors(errors_boat_model):
    verbosity = ""
    if (len(errors_boat_model)>0):
        verbosity += """<span style="font-size:2em;">Model Errors</span>
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
            verbosity += """  %s
        %s%s</td>
        %s%s</td>
        %s%s</td>
      </tr>\n""" %  (tr,td,item[0],td,item[1],td,item[2])
        verbosity += """</table>\n<p>&nbsp;</p>\n<p>&nbsp;</p>"""
    return verbosity


def format_errors(errors_dealer, errors_boat_model, errors_hull):
    return format_hull_errors(errors_hull) + format_dealer_errors(errors_dealer) + format_boat_model_errors(errors_boat_model)

def push_sheet(xlshulls):
    if dbgs:
        dbg(1, "skipping pushing to server")
        return
    silent = dbgs < 1
    db = TunnelSQL(silent, cursor='DictCursor')
    sql = """TRUNCATE TABLE wp_nrb_hulls"""
    _ = db.execute(sql)
    dbg(2,f"Hulls to insert: {len(xlshulls)}")

    sql = """
    INSERT INTO wp_nrb_hulls (
        hull_serial_number, dealership, model,
        last_name, first_name, phone,
        mailing_address, mailing_city, mailing_state, mailing_zip,
        street_address, street_city, street_state, street_zip,
        email_address,
        date_purchased, date_delivered, date_finished, pin,
        opr, css
    ) VALUES (
        %s, %s, %s,
        %s, %s, %s,
        %s, %s, %s, %s,
        %s, %s, %s, %s,
        %s,
        %s, %s, %s, %s,
        %s, %s
    )"""
    db.executemany(sql, sorted(xlshulls))
    dbg(2,"Hulls Inserted")
    db.close()


def mail_results(subject, body):
    if dbgs:
        return
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
    m.send()

@click.command()
@click.option('--debug', '-d', is_flag=True, help='show debug verbosity/do not'
              'save verbosity')
@click.option('--verbose', '-v', default=0, type=int, help='verbosity level 0-3')
def main(debug, verbose):
    global xlsfile
    global dbgs
    global verbosity
    verbosity = verbose
    dbgs = debug

    # load environmental variables
    load_dotenv(dotenv_path=resource_path(".env"))

    if os.getenv('HELP'):
      with click.get_current_context() as ctx:
        click.echo(ctx.get_help())
        ctx.exit()

    xlsfile = os.getenv('XLSFILE')

    if verbosity > 0:
        try:
            print(f"{xlsfile} is {os.path.getsize(xlsfile)} bytes in size")
        except OSError as e:
            print(f"{xlsfile} is not found")

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
            "<br /><br /><pre>" + traceback.format_exc() + "</pre>")
    except Exception as e:
        mail_results(
            'Registrations and Dealer Inventory Sheet Processing Error',
            '<p>Website can not be updated due to error on sheet:<br />\n' + e + '</p>'
            "<br /><br /><pre>" + traceback.format_exc() + "</pre>")
    sys.exit(0)

if __name__ == "__main__":
    main()
