import sys
import xlrd
from datetime import datetime


class Event:
    def __init__(self, bot):
        self.bot = bot
        self.workbook1 = xlrd.open_workbook(r"How did we do as a Recruiter_.xlsx", on_demand=True)
        self.sheet1 = self.workbook1.sheet_by_name("Sheet1")
        self.sheet2 = self.workbook1.sheet_by_name("Form1")


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
                if (len(data) == 2):
                    consultant = data[1]
                    print(consultant)
                    for row_num in range(1,self.sheet1.nrows):
                        row_value = self.sheet1.row_values(row_num)
                        if (data[1] in row_value[0]):
                            response = row_value
                            response = """{} has recieved {:0.0f} feedbacks and has an average rating of {:0.2f}""".format(consultant,response[2], response[1])
                            response += "\nShowing the ones with more than 100 characters\n"
                            for row_num in range(1,self.sheet2.nrows):
                                row_value = self.sheet2.row_values(row_num)
                                if (data[1] in row_value[5] and int(row_value[7]) >= 1 and len(row_value[9]) > 100 ):
                                    print(row_num)
                                    response += "\n----------------------------------------\n"
                                    dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(row_value[1]) - 2)
                                    response += """Date: {} ,  Rating: {} \n{} """.format(dt,str(row_value[7]),row_value[9] )
                            break
                else:
                    consultant = "Total" 
                    print(consultant)
                    for row_num in range(1,self.sheet1.nrows):
                        row_value = self.sheet1.row_values(row_num)
                        if (consultant in row_value[0]):
                            response = row_value
                            response = """As a team we have recieved {:0.0f} feedbacks and have an average rating of {:0.2f}.""".format(response[2], response[1])
                            break 
                   
                channel = event["channel"]
                self.send_message(channel, response)
            except:
                print("Unexpected error:", sys.exc_info()[0])
                raise

    def send_message(self, channel, response):
        self.bot.slack_client.api_call("chat.postMessage", channel=channel, text=response, as_user=True)
