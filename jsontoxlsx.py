import json
import pandas as pd


def load_log():
    with open('autopoe/logdata.json') as json_file:  
        data = json.load(json_file)
        train = pd.DataFrame.from_dict(data)
        train.to_excel("autopoe/logdata.xlsx")



load_log()
