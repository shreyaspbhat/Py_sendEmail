import openpyxl
import time
import subprocess
import itertools

wb = openpyxl.load_workbook('sheet.xlsx')
#print(wb.sheetnames)
sheet = wb['Data']


def remove(sheet):                                              #function to remove empty rows in the middle

    for row in sheet.iter_rows():

        # all() return False if all of the row value is None
        if not all(cell.value for cell in row):
            # detele the empty row
            sheet.delete_rows(row[0].row, 1)

            # recursively call the remove() with modified sheet data
            remove(sheet)

            return

#print('max row is', sheet.max_row)

#print evetything
#for i in range(1, sheet.max_row+1, 1):
#        print(sheet.cell(row=i, column=1).value, sheet.cell(row=i, column=2).value, sheet.cell(row=i, column=3).value, sheet.cell(row=i, column=4).value, sheet.cell(row=i, column=5).value)

#count of email IDs in dictionary

email_ids = list()
counts = dict()

for i in range(2, sheet.max_row+1, 1):
    a = sheet.cell(row=i, column=5).value
    email_ids.append(a)

#print(email_ids)

for email_id in email_ids:
    counts[email_id] = counts.get(email_id, 0) + 1
#print(counts)

unique_mail_ids = list(counts)
#print(unique_mail_ids)
total_number_of_unique_mail_ids = len(unique_mail_ids)
#print(total_number_of_unique_mail_ids)
#print(unique_mail_ids)
total = total_number_of_unique_mail_ids

for i in range(0, total ):                          #to slice workbook into multiple sheets per owner name
    wb.create_sheet(index=i+1, title=str(i))
    #time.sleep(1)
    #print(i, unique_mail_ids[i])

wb.save('sheet.xlsx')

z = 0

#current_name = unique_mail_ids[0]

for j in unique_mail_ids:                              #write data to all the sheets in the workbook


    for i in range(2, sheet.max_row+1, 1):
        #a = sheet.cell(row=i, column=5).value
        if j == sheet.cell(row=i, column=5).value:
            #print(sheet.cell(row=i, column= 1).value, sheet.cell(row=i, column= 2).value, sheet.cell(row=i, column= 3).value, sheet.cell(row=i, column= 4).value, sheet.cell(row=i, column= 5).value)
            a = sheet.cell(row=i, column=1).value
            b = sheet.cell(row=i, column=2).value
            c = sheet.cell(row=i, column=3).value
            d = sheet.cell(row=i, column=4).value
            e = sheet.cell(row=i, column=5).value
           # f = sheet.cell(row=i, column=6).value
            x = str(z)
            sheet1 = wb[x]

            sheet1['A1'] = 'Region'
            sheet1['B1'] = 'Server Name'
            sheet1['C1'] = 'GXP'
            sheet1['D1'] = 'Usage as per CMDB'
            sheet1['E1'] = 'Owner1'
           # sheet1['F1'] = 'Owner2'



            sheet1.cell(row=i, column=1).value = a
            sheet1.cell(row=i, column=2).value = b
            sheet1.cell(row=i, column=3).value = c
            sheet1.cell(row=i, column=4).value = d
            sheet1.cell(row=i, column=5).value = e
           # sheet1.cell(row=i, column=6).value = f

        sheet = wb['Data']
    z = z + 1
            #print('got the desired value at row', i, sheet.cell(row=i, column=5).value)


wb.save('sheet.xlsx')


for j in range(0, total, 1):           # calling remove() to remove empty rows in the middle
    h = str(j)
    sheet1 = wb[h]
    for row in sheet1:
        remove(sheet1)

wb.save('sheet.xlsx')


wb = openpyxl.load_workbook('sheet.xlsx')

for i in range(0, total, 1):
    k = str(i)
    sheet2 = wb[k]
    owner1_name = sheet2.cell(2, 5).value

    row = sheet2.max_row
    # second_owner = sheet2.cell(2, 6).value
    column = sheet2.max_column
    server_list = list()
    usage_list = list()
    for j in range(2, row + 1):
        servers = sheet2.cell(j, 2).value
        usage = sheet2.cell(j, 4).value
        # actual_list_with_region = str(servers) + "             " + str(region)
        server_list.append(servers)
        usage_list.append(usage)

    #print(server_list)
    #print(usage_list)

    # print(server_list)

    file1 = open("mailcontent.txt", "w")

    L2 = ["""<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">

<head>
    <meta http-equiv=Content-Type content="text/html; charset=us-ascii">
    <meta name=Generator content="Microsoft Word 15 (filtered medium)">
    <style>
        <!--
        /* Font Definitions */
        @font-face {
            font-family: "Cambria Math";
            panose-1: 2 4 5 3 5 4 6 3 2 4;
        }

        @font-face {
            font-family: Calibri;
            panose-1: 2 15 5 2 2 2 4 3 2 4;
        }

        /* Style Definitions */
        p.MsoNormal,
        li.MsoNormal,
        div.MsoNormal {
            margin: 0cm;
            margin-bottom: .0001pt;
            font-size: 11.0pt;
            font-family: "Calibri", sans-serif;
            mso-fareast-language: EN-US;
        }

        h2 {
            mso-style-priority: 9;
            mso-style-link: "Heading 2 Char";
            mso-margin-top-alt: auto;
            margin-right: 0cm;
            mso-margin-bottom-alt: auto;
            margin-left: 0cm;
            font-size: 18.0pt;
            font-family: "Calibri", sans-serif;
            font-weight: bold;
        }

        a:link,
        span.MsoHyperlink {
            mso-style-priority: 99;
            color: #0563C1;
            text-decoration: underline;
        }

        a:visited,
        span.MsoHyperlinkFollowed {
            mso-style-priority: 99;
            color: #954F72;
            text-decoration: underline;
        }

        p.MsoPlainText,
        li.MsoPlainText,
        div.MsoPlainText {
            mso-style-priority: 99;
            mso-style-link: "Plain Text Char";
            margin: 0cm;
            margin-bottom: .0001pt;
            font-size: 11.0pt;
            font-family: "Calibri", sans-serif;
            mso-fareast-language: EN-US;
        }

        p {
            mso-style-priority: 99;
            mso-margin-top-alt: auto;
            margin-right: 0cm;
            mso-margin-bottom-alt: auto;
            margin-left: 0cm;
            font-size: 12.0pt;
            font-family: "Times New Roman", serif;
        }

        span.EmailStyle17 {
            mso-style-type: personal-compose;
            font-family: "Calibri", sans-serif;
            color: windowtext;
        }

        span.Heading2Char {
            mso-style-name: "Heading 2 Char";
            mso-style-priority: 9;
            mso-style-link: "Heading 2";
            font-family: "Calibri", sans-serif;
            mso-fareast-language: EN-IN;
            font-weight: bold;
        }

        span.PlainTextChar {
            mso-style-name: "Plain Text Char";
            mso-style-priority: 99;
            mso-style-link: "Plain Text";
            font-family: "Calibri", sans-serif;
        }

        .MsoChpDefault {
            mso-style-type: export-only;
            font-family: "Calibri", sans-serif;
            mso-fareast-language: EN-US;
        }

        @page WordSection1 {
            size: 612.0pt 792.0pt;
            margin: 72.0pt 72.0pt 72.0pt 72.0pt;
        }

        div.WordSection1 {
            page: WordSection1;
        }
        -->
    </style>
    <!--[if gte mso 9]><xml>
    <o:shapedefaults v:ext="edit" spidmax="1026" />
    </xml><![endif]-->
    <!--[if gte mso 9]><xml>
    <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1" />
    </o:shapelayout></xml><![endif]-->
</head>

<body lang=EN-IN link="#0563C1" vlink="#954F72">
    <div class=WordSection1>
        <h2>Dear Linux customer, <o:p></o:p>
        </h2>
        <p>You are receiving this email because you are registered in our database as a responsible person for a
            system(s) that has been identified as requiring patching during the second Linux patching cycle of
            2021.&nbsp; Below you will find details of this patching cycle.&nbsp; We appreciate your cooperation on this
            requirement. <o:p></o:p>
        </p>
        <p class=MsoNormal><b>Affected Systems</b>:<o:p></o:p>
        </p>
        <p class=MsoNormal>
            <o:p>&nbsp;</o:p>
        </p>
        <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=0
            style='width:208.0pt;border-collapse:collapse'>
            <tr style='height:14.5pt'>
                <td width=156 nowrap
                    style='width:117.0pt;border:solid windowtext 1.0pt;background:#002060;padding:0cm 5.4pt 0cm 5.4pt;height:14.5pt'>
                    <p class=MsoNormal><b><span style='font-size:10.0pt;color:white;mso-fareast-language:EN-IN'>Server
                                Name<o:p></o:p></span></b></p>
                </td>
                <td width=121 nowrap
                    style='width:91.0pt;border:solid windowtext 1.0pt;border-left:none;background:#002060;padding:0cm 5.4pt 0cm 5.4pt;height:14.5pt'>
                    <p class=MsoNormal><b><span style='font-size:10.0pt;color:white;mso-fareast-language:EN-IN'>Usage as
                                per CMDB<o:p></o:p></span></b></p>
                </td>
            </tr>"""]






    L3 = ["""</table>
        <p class=xmsonormal style='margin-bottom:12.0pt'><span lang=EN-US>
                <o:p>&nbsp;</o:p>
            </span></p>
        <p class=xmsonormal style='margin-bottom:12.0pt'><span
                style='font-size:12.0pt;font-family:"Times New Roman",serif'>&nbsp;</span><span lang=EN-US>
                <o:p></o:p>
            </span></p>
        <p class=xmsonormal><span style='font-size:12.0pt;font-family:"Times New Roman",serif'>The maximum planned
                interruption of your system is 2 hours, but we expect that your system will be available earlier.
            </span><span lang=EN-US>
                <o:p></o:p>
            </span></p>
        <p class=xmsonormal><span style='font-size:12.0pt;font-family:"Times New Roman",serif'>&nbsp;</span><span
                lang=EN-US>
                <o:p></o:p>
            </span></p>
        <p class=xmsonormal><span style='font-size:12.0pt;font-family:"Times New Roman",serif'>&nbsp;</span><span
                lang=EN-US>
                <o:p></o:p>
            </span></p>
        <p class=xmsoplaintext><b><span style='font-size:12.0pt;font-family:"Times New Roman",serif'>Thanks and
                    Regards</span></b><span style='font-size:12.0pt;font-family:"Times New Roman",serif'>, </span><span
                lang=EN-US>
                <o:p></o:p>
            </span></p>
        <p class=xmsoplaintext><span style='font-size:12.0pt;font-family:"Times New Roman",serif'>CCP Linux Patching
                Team </span><span lang=EN-US>
                <o:p></o:p>
            </span></p>
        <p class=xmsonormal><span style='font-size:12.0pt;font-family:"Times New Roman",serif'>&nbsp;</span><span
                lang=EN-US>
                <o:p></o:p>
            </span></p>
        <p class=MsoNormal>
            <o:p>&nbsp;</o:p>
        </p>
    </div>
</body>

</html>"""]

    file1.writelines(L2)
    # file1.write("Hello \n")

    file1.close()

    file1 = open("mailcontent.txt", "a+")

    for (h, n) in zip(server_list, usage_list):
        L11 = ["""<tr style='height:14.5pt'>
                <td width=128 nowrap valign=bottom
                    style='width:96.0pt;border:solid windowtext 1.0pt;border-top:none;padding:0cm 5.4pt 0cm 5.4pt;height:14.5pt'>
                    <p class=MsoNormal><span style='color:black'>"""]
        file1.writelines(L11)
        file1.writelines(h)

        L12 = ["""<o:p></o:p></span></p>
                </td>
                <td width=117 nowrap valign=bottom
                    style='width:88.0pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt;height:14.5pt'>
                    <p class=MsoNormal><span style='color:black'>"""]
        file1.writelines(L12)
        file1.writelines(n)



        #print(h, n)

        L13 = ["""<o:p></o:p></span></p>
                </td>
            </tr>"""]
        file1.writelines(L13)


#    with open('mailcontent.txt', 'a') as f:
#        f.writelines('\n'.join(server_list))

    file1.close()

    file1 = open("mailcontent.txt", "a+")

    file1.writelines(L3)

    file1.close()

    l = """mailx -s "DOWNTIME REQUIRED : Monsanto Linux Server Patching" -a "From: CCP Linux Patching Team<ccpbayerops.in@capgemini.com>" -a "Content-type: text/html" -c shreyas.p@capgemini.com. """
    m = str(owner1_name)
    # o = ","
    # p = second_owner
    n = """ < "mailcontent.txt" """

    # cmd = l + m + o + p + n
    cmd = l + m + n
    # print(cmd)

    # cmd = """mailx -s "Linux OS Patching and Downtime Request - Announcemen" -a "From: CCP Linux Team<ccpbayerops.in@capgemini.com>" -c shreyas.p@capgemini.com. shreyas.p@capgemini.com < "mailcontent.txt" """

    subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    print("mail sent to:")
    print(i)
    print("Delay of 15 seconds")
    time.sleep(15)


