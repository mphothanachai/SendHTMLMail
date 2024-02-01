import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime as dt
from pathlib import Path

#variable for keep
cc_mail = ['person1@gmail.com', 'person2@gmail.com']
html_header = """
<html>
    <body>
        <h1 style="font-size: 1.8rem;">Last 14 days activity</h1>
        <table style="border-collapse: collapse;">
            <tr>
                <th style="border: 1px solid black; padding: 0.2rem 1rem; text-align:center; font-weight:bold;">date</th>
                <th style="border: 1px solid black; padding: 0.2rem 1rem; text-align:center; font-weight:bold;">timestamp</th>
                <th style="border: 1px solid black; padding: 0.2rem 1rem; text-align:center; font-weight:bold;">invoice</th>
                <th style="border: 1px solid black; padding: 0.2rem 1rem; text-align:center; font-weight:bold;">status1</th>
                <th style="border: 1px solid black; padding: 0.2rem 1rem; text-align:center; font-weight:bold;">status2</th>
                <th style="border: 1px solid black; padding: 0.2rem 1rem; text-align:center; font-weight:bold;">status3</th>
                <th style="border: 1px solid black; padding: 0.2rem 1rem; text-align:center; font-weight:bold;">status4</th>
            </tr>
"""
html_footer = """
        </table>
    </body>
</html>
"""

# Read CSV files
today = dt.date.today()
REPORT_LOOK_BACK = 14
REPORT_PATH = "/PATH_CSV/"

html = html_header

# Loop read csv and keep to list in variable
for _d in range(1, REPORT_LOOK_BACK):
    
    day_back = (today - dt.timedelta(days=_d)).strftime("%Y-%m-%d")
    _f = f"{REPORT_PATH}/{day_back}.csv"
    lines = []
    #HTML nodata
    try:
        with open(_f, 'r', newline='', encoding="UTF8") as file:
            lines = file.readlines()
    except Exception as e:
        print(f"{_f} no data")
        html += f"""
            <tr>
                <td style="border: 1px solid black; padding: 0.2rem 1rem;">{day_back}</td>
                <td colspan="6" style="border: 1px solid black; padding: 0.2rem 1rem; text-align:center; ">-- No Data --</t>
            </tr>
            """
        continue

    # Process each line and skip header line
    for line in lines[1:]:
        # Read column
        data = line.strip().split(',')
        timestamp = data[0]
        try:
            # Clean Timestamp, Remove millisecond
            timestamp = dt.datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S.%f").strftime("%Y-%m-%d %H:%M:%S")
        except Exception as e:
            pass  # skip, do nothing
        #Format data
        col_status1 = f'{int(data[2]):,}' if int(data[2]) != 0 else "-"
        col_status2 = f'{int(data[3]):,}' if int(data[3]) != 0 else "-"
        col_status3 = f'{int(data[4]):,}' if int(data[4]) != 0 else "-"
        col_status4 = f'{int(data[5]):,}' if int(data[5]) != 0 else "-"


        html += f"""
        <tr>
            <td style="border: 1px solid black; padding: 0.2rem 1rem;">{day_back}</td>
            <td style="border: 1px solid black; padding: 0.2rem 1rem;">{timestamp}</td>
            <td style="border: 1px solid black; padding: 0.2rem 1rem;">{data[1]}</td>
            <td style="border: 1px solid black; padding: 0.2rem 1rem; text-align: right;">{col_status1}</td>
            <td style="border: 1px solid black; padding: 0.2rem 1rem; text-align: right;">{col_status2}</td>
            <td style="border: 1px solid black; padding: 0.2rem 1rem; text-align: right;">{col_status3}</td>
            <td style="border: 1px solid black; padding: 0.2rem 1rem; text-align: right;">{col_status4}</td>
        </tr>
        """
html += html_footer
#Create email message
msg = MIMEMultipart()
msg['Subject'] = 'Topic mail'
msg['From'] = 'From'
msg['To'] = 'To'
msg['Cc'] = ', '.join(cc_mail)
msg.attach(MIMEText(html, 'html'))
with smtplib.SMTP('Ip_address', port=15, local_hostname="Localname") as smtpobj:
    smtpobj.sendmail(msg['From'], [msg['To']] + cc_mail, msg.as_string())


