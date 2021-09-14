import json
import base64
import pandas as pd

def lambda_handler(event, context):
    # TODO implement
    #print(event);
    base64_message = event['Body']
    print(base64_message)
    #base64_bytes = base64_message.encode('utf-8')
    #message_bytes = base64.b64decode(base64_bytes)
    #message = message_bytes.decode('utf-8')
    message = str(base64.b64decode(base64_message))
    print(message)

    #excelStart=message.index('spreadsheetml.sheet')
    #excelStart=excelStart+27

    excelStart=message.index(';base64,')
    excelStart=excelStart+8

    excelEnd=message.index('------WebKitFormBoundary',excelStart)-4
    print('start: '+str(excelStart) +'  end: '+str(excelEnd))
    excelBody=message[excelStart:excelEnd]

    print("excelBody: " + excelBody)

    excelBytes=base64.b64decode(excelBody)
    print(excelBytes)

    filename='/tmp/test.xlsx'
    # Open the file with writing permission
    myfile = open(filename, 'wb')
    myfileByteArray = bytearray(excelBytes)
    myfile.write(myfileByteArray)
    # Write a line to the file
    #myfile.write(excelBytes)

    # Close the file
    myfile.close()

    sheet = pd.read_excel(filename)
    df = pd.DataFrame(sheet)

    html='<html><body><table>'
    html=html+'<tr><td>BU</td><td>Tag</td><td>Group</td><td>Type</td><td>Year</td><td>Month</td><td>Value</td></tr>'
    for index, row in df.iterrows():
        html=html+'<tr><td>'+row["BU"]+'</td><td>'+row["Tag"]+'</td><td>'+row["Group"]+'</td><td>'+row["Type"]+'</td><td>'+str(row["Year"])+'</td><td>'+str(row["Month"])+'</td><td>'+str(row["Value"])+'</td></tr>'
    #print row["c1"], row["c2"]
    #BU	Tag	Group	Type	Year	Month	Value
    #print(df)
    html=html+'</table></body><html>'
    print(html)

    return (html)
