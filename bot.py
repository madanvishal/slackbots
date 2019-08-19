import time
from slackclient import SlackClient

import event

class Bot(object):

    def __init__(self):
        self.slack_client = SlackClient("xoxb-538983125493-541180241446-EdYBhuOqZmvxYR217h1TglR0")
        self.bot_name = "duplicitycheckbot" # NAME OF YOUR BOT HERE
        self.bot_id = self.get_bot_id()

        if self.bot_id is None:
            print("Error, could not find " + self.bot_name)

        self.event = event.Event(self)
        self.listen()

    def get_bot_id(self):
        api_call = self.slack_client.api_call("users.list")
        if api_call.get("ok"):
            # retrieving all users so that we can find out bot
            users = api_call.get("members")
            for user in users:
                if "name" in user and user.get("name") == self.bot_name:
                    return "<@" + user.get("id") + ">"
        return None

    def listen(self):
        if self.slack_client.rtm_connect(with_team_state=False):
            print("Successfully connected, listening for commands")
            while True:
                # to keep it running always
                self.event.wait_for_event()
                time.sleep(1)
        else:
            print("Error, Connection failed!")
