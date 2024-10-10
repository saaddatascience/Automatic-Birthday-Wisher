import pandas as pd
import datetime
import smtplib

# Enter authentication details 

GMAIL_ID = ''
GMAIL_PSWD = ''

def sendEmail(to , sub , msg):
    print(f'Email to {to} send with subject: {sub} and messagen {msg}')
    s = smtplib.SMTP('smtp.gmail.com', 587)  # Use port 587 for TLS
    s.starttls()
    s.login(GMAIL_ID , GMAIL_PSWD)
    
    s.sendmail(GMAIL_ID , to , f"subject: {sub}\n\m{msg}")
    s.quit()    
    

if __name__ == "__main__":
    # sendEmail(GMAIL_ID , "subject" , "test message ")
    
    df = pd.read_excel("data.xlsx.xlsx")
    
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")
    
    writeindex = []
     
     
    for index, item in df.iterrows():
        bday = item['Birthday'].strftime("%d-%m")
        print(bday)
        if(today==bday) and yearNow not in  str(item['Year']):
            sendEmail(item['Email'] , "Happy Birthday" , item['Dialogue'])
            writeindex.append(index)
    
    for i in writeindex:
        yr = df.loc[i , 'Year'] 
        df.loc[i , 'yer'] = str(yr) + ',' + str(yearNow)
    
    df.to_excel("data.xlsx.xlsx" , index = False)          