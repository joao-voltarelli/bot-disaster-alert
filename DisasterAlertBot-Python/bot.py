from botcity.core import DesktopBot
from datetime import datetime
import pandas as pd
from pathlib import Path
from PIL import ImageGrab

class Bot(DesktopBot):
    def action(self, execution):
        self.load()

        path_dir_alerts = "Alerts/"
        Path(path_dir_alerts).mkdir(exist_ok=True)

        #READING THE DISASTER ALERTS WORKSHEET AND FILTERING THE DATA
        data = pd.read_excel('disaster_alert.xlsx')
        data.dropna(axis='columns', how='all')
        print(data)

        #OPENING GLOBAL DISASTER ALERT PAGE ON BROWSER
        self.execute("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe")
        self.wait(5000)
        self.paste("https://gdacs.org/Alerts/default.aspx", 2000)
        self.enter(5000)

        #FOR EACH LINE OF THE WORKSHEET, SEARCH THE EVENT DESCRIPTION AND COLLECT THE INFORMATION
        for index, row in data.iterrows():
            for col in data.columns:
                if str(col) == 'Event':
                    event = str(row[col])
                elif str(col).__contains__("From"):
                    from_date = datetime.strftime(row[col], '%Y-%m-%d')
                elif str(col).__contains__("To"):
                    to_date = datetime.strftime(row[col], '%Y-%m-%d')
                elif str(col).__contains__("Level"):
                    level = str(row[col])
                elif str(col).__contains__("Severity"):
                    severity = str(row[col])
                elif str(col).__contains__("Country"):
                    country = str(row[col])

            #CREATING FOLDER WHERE MAP PRINT AND RESULT FILE WILL BE SAVED
            path_dir = path_dir_alerts + event + " - (" + from_date + "_" + to_date + ")"
            Path(path_dir).mkdir(exist_ok=True)
            
            print("\nAlert=> " + event + " | " + from_date + " | " + to_date + " | " + level + " | " + severity + " | " + country + "\n")

            if not self.find( "search", matching=0.97, waiting_time=10000):
                self.not_found("search")
            self.click(2000)

            #SELECTS THE GLOBAL DISASTER TYPE 
            self.tab()
            if event.__contains__("Earthquakes"):
                self.space(1000)
            self.tab()
            if event.__contains__("Tsunamis"):
                self.space(1000)
            self.tab()
            if event.__contains__("Floods"):
                self.space(1000)
            self.tab()
            if event.__contains__("Cyclones"):
                self.space(1000)
            self.tab()
            if event.__contains__("Volcanoes"):
                self.space(1000)
            self.tab()
            if event.__contains__("Droughts"):
                self.space(1000)
            self.tab()
            if event.__contains__("Forest Fires"):
                self.space(1000)

            #TYPE THE DATES FOR THE EVENT PERIOD 
            if not self.find( "from", matching=0.97, waiting_time=10000):
                self.not_found("from")
            self.click_relative(160, 11)
            self.control_a(1000)
            self.paste(from_date, 1000)
            self.enter()

            if not self.find( "to", matching=0.97, waiting_time=10000):
                self.not_found("to")
            self.click_relative(167, 10)
            self.control_a(1000)
            self.paste(to_date, 1000)
            self.enter()

            #SELECT THE LEVEL
            if not self.find( "level", matching=0.97, waiting_time=10000):
                self.not_found("level")
            self.click_relative(401, 12)
            if level == "Orange":
                self.enter()
            elif level == "Red":
                self.type_up()
                self.enter()
            elif level == "Green":
                self.type_down()
                self.enter()

            if severity != "nan":
                if not self.find( "severity", matching=0.97, waiting_time=10000):
                    self.not_found("severity")
                self.click_relative(95, 10)
                self.control_a(1000)
                self.paste(severity, 1000)
                self.enter()

            #TYPE THE COUNTRY NAME
            if not self.find( "country", matching=0.97, waiting_time=10000):
                self.not_found("country")
            self.click_relative(95, 10)
            self.paste(country, 1000)

            #SUBMITTING THE SEARCH
            if not self.find( "submit", matching=0.97, waiting_time=10000):
                self.not_found("submit")
            self.click(2000)
            self.scroll_down(500)
            self.wait(25000)

            #PRINT THE MAP AREA
            map_image = ImageGrab.grab(bbox=(400,350,1500,770))
            map_image.save(path_dir + "/Map-Alerts.png")

            #COLLECTING THE DESCRIPTIONS OF SURVEYED EVENTS 
            self.control_a(1000)
            self.control_c(1000)
            conteudoPagina = str(self.get_clipboard())
            description = conteudoPagina[conteudoPagina.find('RESULTS'):]
            description = description[:description.find('HOME(current)')]
            print(description)

            f = open(path_dir + "/RESULTS.txt", "w")
            f.write(description.replace("\t", "-"))
            f.close()

            self.moveTo(300, 350)
            self.scroll_up(500)
            self.wait(1000)
            self.key_f5(4000)

        self.control_w()
    
    def load(self):
        self.add_image("search", self.get_resource_abspath("search.png"))
        self.add_image("from", self.get_resource_abspath("from.png"))
        self.add_image("to", self.get_resource_abspath("to.png"))
        self.add_image("level", self.get_resource_abspath("level.png"))
        self.add_image("severity", self.get_resource_abspath("severity.png"))
        self.add_image("country", self.get_resource_abspath("country.png"))
        self.add_image("submit", self.get_resource_abspath("submit.png"))

    def not_found(self, label):
        print(f"Element not found: {label}")
    
if __name__ == '__main__':
    Bot.main()




