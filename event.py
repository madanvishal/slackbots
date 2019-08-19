import sys
import xlrd
from datetime import datetime


class Event:
    def __init__(self, bot):
        self.bot = bot
        self.workbook1 = xlrd.open_workbook(r"antalmagarpattacity-candidatetracker-archive.xlsx", on_demand=True)
        self.sheet1 = self.workbook1.sheet_by_name("Candidates")
        self.workbook2 = xlrd.open_workbook(r"antalmagarpattacity-candidatetracker-master.xlsx", on_demand=True)
        self.sheet2 = self.workbook2.sheet_by_name("Candidates")


    def wait_for_event(self):
        events = self.bot.slack_client.rtm_read()
        if events:
            for event in events:
                # print(evnet)
                self.parse_event(event)

    def parse_event(self, event):
        # if our bot get's mentioned, in a text message
        if event and "text" in event and self.bot.bot_id in event["text"]:
            try:
                data = event["text"].split()
                print(len(data))
                if (len(data) == 3 and data[1].upper() == "EMAIL"):
                    case = "Email"
                    emaildata1 = data[2].split(":")
                    emaildata2 = emaildata1[1].split("|")
                    emaildata = emaildata2[0]
                    print(case)
                    print(emaildata)
                else:
                    if (len(data) == 3 and data[1].upper() == "PHONE"):
                        case = "Phone"
                        print(case)
                    else:
                        if(len(data) == 5 and data[1].upper() == "EMAIL" and data[3].upper() == "PHONE"):
                            case = "Both"
                            emaildata1 = data[2].split(":")
                            emaildata2 = emaildata1[1].split("|")
                            emaildata = emaildata2[0]
                            print(case)
                            print(emaildata)
                            print(case)
                        else:
                            case = "Invalid"
                            response = "Usage : @duplicitycheckbot [Email/Phone] emailaddress/phone"
                   

                if(case == "Phone"):
                    print(data[2])
                    phonefound = 0
                    for row_num in range(1,self.sheet1.nrows):
                        row_value = self.sheet1.row_values(row_num)
                        #print(row_value[32])
                        if (data[2] in row_value[32]):
                            response = row_value
                            dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(response[0]) - 2)
                            response = """Candidate {} was submitted by {} on {}""".format(response[2], response[1],dt)
                            phonefound = 1
                            print("Found in Archive")
                            break
                    if (phonefound == 0):
                        print("Checking Master..")
                        for row_num in range(1,self.sheet2.nrows):
                            row_value = self.sheet2.row_values(row_num)
                            #print(row_value[43])
                            if (data[2] in row_value[43]):
                                response = row_value
                                dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(response[0]) - 2)
                                response = """Candidate {} was submitted by {} on {}""".format(response[2], response[1],dt)
                                print("Found in Master")
                                break
                            else:
                                response = "Phone not found"

                if(case == "Email"):
                    print(data[2])
                    print(emaildata)
                    found = 0
                    for row_num in range(1,self.sheet1.nrows):
                        row_value = self.sheet1.row_values(row_num)
                        #print(row_value[32])
                        if (emaildata in row_value[3]):
                            response = row_value
                            dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(response[0]) - 2)
                            response = """Candidate {} was submitted by {} on {}""".format(response[2], response[1],dt)
                            found = 1
                            print("Found in Archive")
                            break
                    if (found == 0):
                        print("Checking Master..")
                        for row_num in range(1,self.sheet2.nrows):
                            row_value = self.sheet2.row_values(row_num)
                            #print(row_value[43])
                            if (emaildata in row_value[3]):
                                response = row_value
                                dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(response[0]) - 2)
                                response = """Candidate {} was submitted by {} on {}""".format(response[2], response[1],dt)
                                print("Found in Master")
                                break
                            else:
                                response = "Email address not found"

                if(case == "Both"):
                    print(data[2])
                    print(emaildata)
                    emailfound = 0
                    for row_num in range(1,self.sheet1.nrows):
                        row_value = self.sheet1.row_values(row_num)
                        #print(row_value[32])
                        if (emaildata in row_value[3]):
                            response = row_value
                            dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(response[0]) - 2)
                            emailresponse = """Email Check: Candidate {} was submitted by {} on {}""".format(response[2], response[1],dt)
                            emailfound = 1
                            print("Found in Archive")
                            break
                    if (emailfound == 0):
                        print("Checking Master..")
                        for row_num in range(1,self.sheet2.nrows):
                            row_value = self.sheet2.row_values(row_num)
                            #print(row_value[43])
                            if (emaildata in row_value[3]):
                                rowvalue = row_value
                                dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(rowvalue[0]) - 2)
                                emailresponse = """Email Check: Candidate {} was submitted by {} on {}""".format(rowvalue[2], rowvalue[1],dt)
                                print("Found in Master")
                                emailfound = 1
                                break
                    phonefound = 0
                    for row_num in range(1,self.sheet1.nrows):
                        row_value = self.sheet1.row_values(row_num)
                        #print(row_value[32])
                        if (data[4] in row_value[32]):
                            rowvalue = row_value
                            dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(rowvalue[0]) - 2)
                            phoneresponse = """& Phone Check: Candidate {} was submitted by {} on {}""".format(rowvalue[2], rowvalue[1],dt)
                            phonefound = 1
                            print("Found in Archive")
                            break
                    if (phonefound == 0):
                        print("Checking Master..")
                        for row_num in range(1,self.sheet2.nrows):
                            row_value = self.sheet2.row_values(row_num)
                            #print(row_value[43])
                            if (data[4] in row_value[43]):
                                rowvalue = row_value
                                dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(rowvalue[0]) - 2)
                                phoneresponse = """ & Phone Check: Candidate {} was submitted by {} on {}""".format(rowvalue[2], rowvalue[1],dt)
                                print("Found in Master")
                                phonefound = 1
                                break
                    if (phonefound == 0):
                        phoneresponse = " Phone not found "
                    if (emailfound == 0):
                        emailresponse = " Email not found "
                    response = emailresponse + phoneresponse


                        
                channel = event["channel"]
                self.send_message(channel, response)
            except:
                print("Unexpected error:", sys.exc_info()[0])
                raise

    def send_message(self, channel, response):
        self.bot.slack_client.api_call("chat.postMessage", channel=channel, text=response, as_user=True)
