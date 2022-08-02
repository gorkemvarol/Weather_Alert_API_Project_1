import json
import requests
import gzip
import pandas as pd
import numpy as np
# import matplotlib.pyplot as plt
# import seaborn as sns
import csv
from time import sleep
import win32com.client as client
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


data = ('../weather_api/city.list.json.gz')
with gzip.open(data , 'rb') as f:
    json_content = json.loads(f.read())
    
# DISTRICT_1
KOHQ1 = [{'MN':'Wayzata'}]
# DISTRICT_2
SHAUN1 = [{'MN':'Bemidji'},{'MN':'Maple Lake'},{'MN':'Annandale'},{'MN':'Maple Plain'},{'MN':'Willmar'},{'MN':'Owatonna'},{'MN':'Waite Park'},{'MN':'Little Falls'},{'MN':'Albert Lea'},{'MN':'Waseca'},{'MN':'Austin'},{'MN':'Pillager'},{'MN':'Nisswa'},{'MN':'Baxter'},{'MN':'Princeton'},{'MN':'Alexandria'},{'MN':'St Cloud'},{'MN':'South Haven'},{'MN':'Pierz'},{'MN':'Clear Water'},{'MN':'Becker'},{'MN':'Rush City'},{'MN':'Hastings'},{'MN':'Big Lake'},{'MN':'Rochestar'},{'MN':'Buffalo'}]
# DISTRICT_3
EMILY1 = [{'ME':'Auburn'},{'ME':'Oxford'},{'ME':'Sanford'},{'ME':'Windham'},{'AR':'Paragould'},{'MO':'Billings'},{'MO':'Brookline'},{'MO':'Ozark'},{'MO':'Republic'},{'OH':'Tipp City'},{'NY':'Evans Mills'},{'NY':'Watertown'},{'TN':'Chattanooga'}]
# DISTRICT_4
JUSTIN1 = [{'SD':'Vermillion'},{'SD':'Elk Point'},{'SD':'Aberdeen'},{'WY':'Evansville'},{'KS':'Salina'},{'KS':'Hutchinson'},{'MT':'Billings'},{'WY':'Evansville'},{'WY':'Casper'},{'ND':'Valley City'},{'ND':'Minot'},{'ND':'Jamestown'},{'WY':'Cheyenne'},{'SD':'Milbank'},{'CA':'Twentynine Palms'}]



           
def get_no(listt):                
    list_=[]
    for num in range(len(listt)):
        for state,city in listt[num].items():
            for i in range(len(json_content)): 
                if json_content[i]['country'] == 'US' and json_content[i]['name'] == city and json_content[i]['state'] == state:
                    list_.append(json_content[i])
    return list_
                
                

#DISTRICT_1
KOHQ2 = get_no(KOHQ1)
#DISTRICT_2
SHAUN2 = get_no(SHAUN1)
#DISTRICT_3
EMILY2 = get_no(EMILY1)
#DISTRICT_4
JUSTIN2 = get_no(JUSTIN1)


def get_api_request(list2):
    list_3=[]
    for query in list2:

        api_url2 = 'http://api.openweathermap.org/data/2.5/weather?id='+str(query['id'])+'&appid=5f0573bd11f9c05220677beb94d32b6f&units=imperial'
        response = requests.get(api_url2)
        list_3.append(response.json())
    return list_3




# DISTRICT_1
KOHQ3 = get_api_request(KOHQ2)
# DISTRICT_2
SHAUN3 = get_api_request(SHAUN2)
# DISTRICT_3
EMILY3 = get_api_request(EMILY2)
# DISTRICT_4
JUSTIN3 = get_api_request(JUSTIN2)



def func(list4): 
    
    #Converting the data to Pandas DataFrame
    data = pd.DataFrame(list4)
    #Copying the data just in case 
    data2 = data.copy(deep=False)
    #Converting the series to str
    data2['weather'] = data2['weather'].astype(str)  
    #Cleaning weather series to get only the weather condition and id
    data2['weather1'] = data2['weather'].map(lambda weather:weather.split("main': '")[1].split("', 'description")[0])
    data2['id']= data2['weather'].map(lambda weather:weather.split("'id': ")[1].split(", 'main'")[0])
    data3 = data2[['weather1','id','name','main']]
    data3 = data3.rename(columns={'name':'city'},errors="raise")  
    #
    data3 = data3.loc[(data3['weather1'] == ('Snow')) & (data3['id'] == '601') | (data3['id'] == '602') ]
    #
    return data3



# DISTRICT_1
KOHQ4 = func(KOHQ3)
html_table_KOHQ4 = KOHQ4.to_html()
# DISTRICT_2
SHAUN4 = func(SHAUN3)
html_table_SHAUN4 = SHAUN4.to_html()
# DISTRICT_3
EMILY4 = func(EMILY3)
html_table_EMILY4 = EMILY4.to_html()
# DISTRICT_4
JUSTIN4 = func(JUSTIN3)
html_table_JUSTIN4 = JUSTIN4.to_html()



def conditional_email_alert(dataframe_):
    
    for x in dataframe_['id']:
        res = sum(1 for x in dataframe_['id'] if (x == '601')| (x == '602'))
        return res
    else:
        return 0
    


# DISTRICT_1
KOHQ5=conditional_email_alert(KOHQ4)
# DISTRICT_2
SHAUN5=conditional_email_alert(SHAUN4)
# DISTRICT_3
EMILY5 = conditional_email_alert(EMILY4)
# DISTRICT_4
JUSTIN5 = conditional_email_alert(JUSTIN4)


#DISTRICT_1


#Sending email to poeple who are listed in csv file
if KOHQ5 >= 1:


        # open distribution list
        with open('DISTRICT_1.csv', 'r', newline='') as f:
            reader = csv.reader(f)
            distro = [row for row in reader]

        # chunk distribution list into blocks of 30
        chunks = [distro[x:x+30] for x in range(0, len(distro), 30)]

        # create outlook instance
        outlook = client.Dispatch('Outlook.Application')


        # iterate through chunks and send mail
        for chunk in chunks:
            # iterate through each recipient in chunk and send mail
            for name, address in chunk:
                message = outlook.CreateItem(0)
                message.To = address
                message.Subject = "KO Storage Snow & Weather Report"
                message.HTMLBody = '<h4 style="font-family:verdana;">'+html_table_KOHQ4+'Real Time Weather Snow Report</h4><p style="font-family:verdana">Note: Snow weather report runs twice daily and reports to you via e-mail only when it is snowing <p style="font-family:verdana">Weather Source:</p><p style="font-family:verdana">https://openweathermap.org/ </p> <p style="font-family:verdana">See facilities above for the latest snowfall report</p><p style="font-family:verdana">Please contact Varol if you would like to subscribe to weather alerts or report any issues,</p><p style="font-family:verdana"> KO Storage</p><img src="KO (45).png" alt="KO Storage logo" width="300" height="150">'
#                 message.HTMLBody = str(cloudy)
                #message.Body = template.format(name)
                message.Send()

            # wait 60 seconds before sending next chunk
            sleep(10)

            