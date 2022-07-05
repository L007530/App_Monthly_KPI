import win32com.client
from win32com.client import Dispatch, constants
import os

# const = win32com.client.constants
# olMailItem = 0x0
# obj = win32com.client.Dispatch("Outlook.Application")
# newMail = obj.CreateItem(olMailItem)
# newMail.Subject = "I AM SUBJECT!!"
# # newMail.Body = "I AM\nTHE BODY MESSAGE!"
# newMail.BodyFormat = 2  # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
# newMail.HTMLBody = "<HTML><BODY>Enter the <span style='color:red'>message</span> text here.</BODY></HTML>"
# newMail.To = "email@demo.com"
# attachmentFile = "1.pdf"
# attachment1 = os.getcwd() + "\\" + attachmentFile
# newMail.Attachments.Add(Source=attachment1)
# newMail.save()
# newMail.display()

# Hard coded email subject
MAIL_SUBJECT = 'AUTOMATED Text Python Email without attachments'

# Hard coded email HTML text
# MAIL_BODY = \
#     '<html> ' \
#     ' <body>' \
#     ' <p><b>Dear</b> Receipient,<br><br>' \
#     ' This is an automatically generated email by <font size="5" color="blue">Python.</font><br>' \
#     ' It is so <del>amazing and</del> fantastic<br>' \
#     ' <strong>Wish you</strong> all the <font size="5" color="green">best</font><br>' \
#     ' </body>' \
#     '</html>'
arrowUp = \
    '''
    <span style='font-size:12.0pt;font-family:
      "Wingdings 3";color:#00B050'>Ç</span><span style='color:black'>
    '''
arrowDown = \
    '''
    <span style='font-size:12.0pt;font-family:
      "Wingdings 3";color:#ED7D31'>È</span><span style='color:black'>
    '''
MAIL_BODY = \
    f'''
        <p class=MsoNormal><span style='font-family:Wingdings;color:black'>Ø</span><span
style='font-size:7.0pt;font-family:"Times New Roman",serif;color:black'>&nbsp;</span><b><u><span
style='color:#1F497D'>LCCP System</span></u></b> <span style='color:#1F497D'><o:p></o:p></span></p>

<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=1397
 style='width:1047.95pt;margin-left:-.05pt;border-collapse:collapse;mso-yfti-tbllook:
 1184;mso-padding-alt:0in 0in 0in 0in'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:.2in'>
  <td width=127 nowrap rowspan=2 style='width:95.0pt;border:solid windowtext 1.0pt;
  background:#9CC2E5;padding:0in 5.4pt 0in 5.4pt;height:.2in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>Application</span></b><span style='color:black'> <b><o:p></o:p></b></span></p>
  </td>
  <td width=119 nowrap rowspan=2 style='width:89.0pt;border:solid windowtext 1.0pt;
  border-left:none;background:#9CC2E5;padding:0in 5.4pt 0in 5.4pt;height:.2in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>Target Audience</span></b><span style='color:black'> </span><b><span
  style='font-size:10.5pt;color:black'><o:p></o:p></span></b></p>
  </td>
  <td width=241 rowspan=2 style='width:181.05pt;border-top:solid windowtext 1.0pt;
  border-left:none;border-bottom:solid black 1.0pt;border-right:solid windowtext 1.0pt;
  background:#9CC2E5;padding:0in 5.4pt 0in 5.4pt;height:.2in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>{'2020 Jan'}</span></b><span style='color:black'> </span><b><span
  lang=ZH-CN style='font-family:DengXian;color:#1F497D'>– </span><span
  style='color:black'>{'2020 Dec'}<br>
  Patient testing record Trend</span></b><span style='color:black'> <b><o:p></o:p></b></span></p>
  </td>
  <td width=250 nowrap colspan=3 style='width:187.6pt;border:solid windowtext 1.0pt;
  border-left:none;background:#9CC2E5;padding:0in 5.4pt 0in 5.4pt;height:.2in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>YTD 2020</span></b><span style='color:black'> <b><o:p></o:p></b></span></p>
  </td>
  <td width=255 nowrap colspan=3 style='width:191.2pt;border:solid windowtext 1.0pt;
  border-left:none;background:#9CC2E5;padding:0in 5.4pt 0in 5.4pt;height:.2in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>MTD {'2020(Dec)'} </span><o:p></o:p></b></p>
  </td>
  <td width=247 colspan=3 style='width:185.1pt;border:solid windowtext 1.0pt;
  border-left:none;background:#9CC2E5;padding:0in 5.4pt 0in 5.4pt;height:.2in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='font-size:10.0pt;color:black'>Up to Date(Since launch from {'Apr 2017'})</span></b><span
  style='color:black'> <b><o:p></o:p></b></span></p>
  </td>
  <td width=159 rowspan=2 style='width:119.0pt;border:solid windowtext 1.0pt;
  border-left:none;background:#9CC2E5;padding:0in 5.4pt 0in 5.4pt;height:.2in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>Remarks / Comments <o:p></o:p></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1;height:.8in'>
  <td width=85 style='width:63.8pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#9CC2E5;
  padding:0in 5.4pt 0in 5.4pt;height:.8in'>
  <p class=MsoNormal align=center style='text-align:center'><b><o:p>&nbsp;</o:p></b></p>
  <p class=MsoNormal align=center style='margin-bottom:12.0pt;text-align:center'><b><span
  style='color:black'>New HCP enrollment <o:p></o:p></span></b></p>
  </td>
  <td width=85 style='width:63.8pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#9CC2E5;
  padding:0in 5.4pt 0in 5.4pt;height:.8in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>New Patient enrollment <o:p></o:p></span></b></p>
  </td>
  <td width=80 style='width:60.0pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#9CC2E5;
  padding:0in 5.4pt 0in 5.4pt;height:.8in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>Patient testing record <o:p></o:p></span></b></p>
  </td>
  <td width=85 style='width:63.8pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#9CC2E5;
  padding:0in 5.4pt 0in 5.4pt;height:.8in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>New HCP enrollment <o:p></o:p></span></b></p>
  </td>
  <td width=85 style='width:63.8pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#9CC2E5;
  padding:0in 5.4pt 0in 5.4pt;height:.8in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>New Patient enrollment <o:p></o:p></span></b></p>
  </td>
  <td width=85 style='width:63.6pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#9CC2E5;
  padding:0in 5.4pt 0in 5.4pt;height:.8in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>Patient testing record <o:p></o:p></span></b></p>
  </td>
  <td width=80 style='width:60.2pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#9CC2E5;
  padding:0in 5.4pt 0in 5.4pt;height:.8in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>Total HCP enrollment <o:p></o:p></span></b></p>
  </td>
  <td width=85 style='width:63.8pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#9CC2E5;
  padding:0in 5.4pt 0in 5.4pt;height:.8in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>Total Patient enrollment <o:p></o:p></span></b></p>
  </td>
  <td width=81 style='width:61.1pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#9CC2E5;
  padding:0in 5.4pt 0in 5.4pt;height:.8in'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='color:black'>Total Patient testing record <o:p></o:p></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:2;mso-yfti-lastrow:yes;height:28.35pt'>
  <td width=127 nowrap style='width:95.0pt;border:solid windowtext 1.0pt;
  border-top:none;padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='color:black'>LCCP System</span> <span style='color:black'><o:p></o:p></span></p>
  </td>
  <td width=119 nowrap style='width:89.0pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='color:black'>Lilly Sales Reps, HCPs, Patients <o:p></o:p></span></p>
  </td>
  <td width=241 nowrap style='width:181.05pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:10.5pt'><!--[if gte vml 1]><v:shape id="图片_x0020_4" o:spid="_x0000_i1029"
   type="#_x0000_t75" alt="" style='width:162.6pt;height:24pt'>
   <v:imagedata src="file:///C:/Users/L007530/OneDrive%20-%20Eli%20Lilly%20and%20Company/Desktop/FW%20China%20Application%20Service%20KPI%20Report%20Dec_files/image009.png"
    o:href="cid:image005.png@01D6EDCC.66A44040"/>
  </v:shape><![endif]--><![if !vml]><img border=0 width=217 height=32
  src="file:///C:/Users/L007530/OneDrive%20-%20Eli%20Lilly%20and%20Company/Desktop/FW%20China%20Application%20Service%20KPI%20Report%20Dec_files/image010.jpg"
  style='height:.333in;width:2.258in' v:shapes="图片_x0020_4"><![endif]></span></p>
  </td>
  <td width=85 nowrap style='width:63.8pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'>{'3,682'} <span
  style='color:black'><o:p></o:p></span></p>
  </td>
  <td width=85 nowrap style='width:63.8pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'>{'57, 380'} <span
  style='color:black'><o:p></o:p></span></p>
  </td>
  <td width=80 nowrap style='width:60.0pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'>{'7, 910, 976'} <span
  style='color:black'><o:p></o:p></span></p>
  </td>
  <td width=85 nowrap style='width:63.8pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'>{'142'} <span
  style='color:black'><o:p></o:p></span></p>
  </td>
  <td width=85 nowrap style='width:63.8pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'>{'9, 070'}<span
  style='color:black'><o:p></o:p></span></p>
  </td>
  <td width=85 nowrap style='width:63.6pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='color:black'>{'690, 388'} </span>{arrowUp} <o:p></o:p></span></p>
  </td>
  <td width=80 nowrap style='width:60.2pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'>{'25, 287'}<span
  style='color:black'><o:p></o:p></span></p>
  </td>
  <td width=85 nowrap style='width:63.8pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'>{'246, 436'}<span
  style='color:black'><o:p></o:p></span></p>
  </td>
  <td width=81 nowrap style='width:61.1pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'>
  <p class=MsoNormal align=center style='text-align:center'>{'21, 407, 015'}<span
  style='color:black'><o:p></o:p></span></p>
  </td>
  <td width=159 nowrap style='width:119.0pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:28.35pt'></td>
 </tr>
</table>
        '''


def send_outlook_html_mail(recipients=None, subject='No Subject', body='Blank', send_or_display='save', copies=None,
                           attachments=[]):
    """
    Send an Outlook HTML email
    :param recipients: list of recipients' email addresses (list object)
    :param subject: subject of the email
    :param body: HTML body of the email
    :param send_or_display: Send - send email automatically | Display - email gets created user have to click Send
    :param copies: list of CCs' email addresses
    :return: None
    """
    if len(recipients) > 0 and isinstance(recipient_list, list):
        outlook = win32com.client.Dispatch("Outlook.Application")

        ol_msg = outlook.CreateItem(0)

        str_to = ""
        for recipient in recipients:
            str_to += recipient + ";"

        ol_msg.To = str_to

        if copies is not None:
            str_cc = ""
            for cc in copies:
                str_cc += cc + ";"

            ol_msg.CC = str_cc

        ol_msg.Subject = subject
        ol_msg.HTMLBody = body

        if len(attachments) > 0:
            for attachment in attachments:
                # attachment1 = os.getcwd() + "\\" + attachmentFile
                ol_msg.Attachments.Add(Source=attachment)

        if send_or_display == 'SEND':
            ol_msg.Send()
        elif send_or_display == 'display':
            ol_msg.Display()
        else:
            ol_msg.Save()
    else:
        print('Recipient email address - NOT FOUND')


if __name__ == '__main__':
    recipient_list = ['recipient1@someemaildomain.com',
                      'recipient2@someemaildomain.com',
                      'recipient3@someemaildomain.com']

    copies_list = ['cc1@someemaildomain.com']
    attachment1 = os.getcwd() + "\\1.pdf"
    attachment2 = os.getcwd() + "\\2.pdf"
    attachments = [attachment1, attachment2]

    send_outlook_html_mail(recipients=recipient_list, subject=MAIL_SUBJECT, body=MAIL_BODY, send_or_display='display',
                           copies=copies_list, attachments=[])
