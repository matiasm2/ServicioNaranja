import requests, poplib, ConfigParser, sys
from email import parser

configParser = ConfigParser.RawConfigParser()
#configFilePath = sys.argv[1]
configFilePath = "./env.cfg"
configParser.read(configFilePath)

def getSessionToken(url, userName, instanceName, password):
    url = url+'/api/core/security/login'
    data = {
        "InstanceName":instanceName,
        "Username":userName,
        "UserDomain": "",
        "Password":password
    }
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Content-Type': 'application/json'
    }
    response = requests.post(url, verify=False, headers=headers, json=data)
    #print response
    #print response.json()['RequestedObject']['SessionToken']
    return response.json()['RequestedObject']['SessionToken']

def postAttachment(url, sessionToken, attachmentName, attachmentBytes):
    url = url+'/api/core/content/attachment'
    data = {
        "AttachmentName":attachmentName,
        "AttachmentBytes":attachmentBytes
    }
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Authorization':'Archer session-id='+sessionToken,
        'Content-Type': 'application/json'
    }
    response = requests.post(url, verify=False, headers=headers, json=data)
    print response
    print response.json()['RequestedObject']['Id']
    return response.json()['RequestedObject']['Id']

def getAttachment(url, sessionToken, iD):
    url = url+'/api/core/content/attachment/' + str(iD)
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Authorization':'Archer session-id='+sessionToken,
        'Content-Type': 'application/json',
        'X-Http-Method-Override': 'GET'
    }
    response = requests.post(url, verify=False, headers=headers)
    print response
    print response.json()

def getApplicationMetadata(url, sessionToken, applicationId):
    url = url+'/api/core/system/application/' + str(applicationId)
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Authorization':'Archer session-id='+sessionToken,
        'Content-Type': 'application/json',
        'X-Http-Method-Override': 'GET'
    }
    response = requests.post(url, verify=False, headers=headers)
    print response
    print response.json()
    return response.json()

def getApplicationFieldDefinition(url, sessionToken, applicationId):
    url = url+'/api/core/system/fielddefinition/application/' + str(applicationId)
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Authorization':'Archer session-id='+sessionToken,
        'Content-Type': 'application/json',
        'X-Http-Method-Override': 'GET'
    }
    response = requests.post(url, verify=False, headers=headers)
    print response
    #print response.json()
    return response

def createJSON(moduleId, msg):
    valuesId = []
    values = []
    fd = getApplicationFieldDefinition(baseurl, sessionToken, moduleId)
    for field in fd.json():
        if field['RequestedObject']['LevelId']:
                levelId = field['RequestedObject']['LevelId']
        elif field['RequestedObject']['Alias'] == 'Texto':
            detalle = field['RequestedObject']['Id']
        elif field['RequestedObject']['Alias'] == 'Adjunto':
            adjunto = field['RequestedObject']['Id']
        elif field['RequestedObject']['Alias'] == 'Asunto':
            asunto = field['RequestedObject']['Id']
        elif field['RequestedObject']['Alias'] == 'De':
            de = field['RequestedObject']['Id']
        elif field['RequestedObject']['Alias'] == 'Para':
            para = field['RequestedObject']['Id']

    for part in msg.walk():
        print part.get_content_type()
        if part.get_content_type() in allowed_mimetypes:

            '''name = part.get_filename()
            print name

            data = part.get_payload(decode=True)
            data2 = part.get_payload(decode=False)
            f = file('archivines/'+str(name),'wb')
            f.write(data)
            iD = postAttachment('https://10.100.107.61', sessionToken, name, data2)
            json = createJSONUpdate('https://10.100.107.61', sessionToken, 239965, levelId, fieldId, iD)
            putContent('https://10.100.107.61', sessionToken, json)
            f.close()
            attachments.append(name)'''
        else:
            print msg


    data = {
        "Content":{
            "LevelId" : levelId,
            "FieldContents" : {
                 str(detalle): {
                     "Type" : 1,
                     "Value" : "hola",
                      "FieldId": str(detalle)
                  },
                  str(adjunto): {
                      "Type" : 11,
                      "Value" : values,
                       "FieldId": str(adjunto)
                   },
                  str(asunto): {
                      "Type" : 1,
                      "Value" : "hola",
                       "FieldId": str(asunto)
                   },
                   str(de): {
                       "Type" : 1,
                       "Value" : "hola",
                        "FieldId": str(de)
                    },
                    str(para): {
                        "Type" : 1,
                        "Value" : "hola",
                         "FieldId": str(para)
                     }
                 }
             }
         }
    print data
    return data

def createJSONUpdate(baseurl, sessionToken,contentId, levelId, fieldId, value):
    registro = getContentById(baseurl, sessionToken, contentId)
    values = registro.json()['RequestedObject']['FieldContents']['30536']['Value']
    if values:
        values.append(value)
    else:
        values = [value]

    data = {
        "Content":{
            "Id": contentId ,
            "LevelId" : levelId,
            "FieldContents" : {
                str(fieldId): {
                    "Type" : 11,
                    "Value" : values,
                     "FieldId": fieldId
                 }
                 }
             }
         }
    return data


def postContent(url, sessionToken, content):
    url = url+'/api/core/content'
    data = content
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Authorization':'Archer session-id='+sessionToken,
        'Content-Type': 'application/json'
    }
    response = requests.post(url, verify=False, headers=headers, json=data)
    print response.json()
    return response


def putContent(url, sessionToken, content):
    url = url+'/api/core/content'
    data = content
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Authorization':'Archer session-id='+sessionToken,
        'Content-Type': 'application/json'
    }
    response = requests.put(url, verify=False, headers=headers, json=data)
    return response

def getContentById(url, sessionToken, contentId):
    url = url+'/api/core/content/' + str(contentId)
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Authorization':'Archer session-id='+sessionToken,
        'Content-Type': 'application/json',
        'X-Http-Method-Override': 'GET'
    }
    response = requests.post(url, verify=False, headers=headers)
    print response
    return response

#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################

def fetchMail(pop_conn, delete_after=False):
    #Get messages from server:
    messages = [pop_conn.retr(i) for i in range(1, len(pop_conn.list()[1]) + 1)]
    # Concat message pieces:
    messages = ["\n".join(mssg[1]) for mssg in messages]
    #Parse message intom an email object:
    messages = [parser.Parser().parsestr(mssg) for mssg in messages]
    if delete_after == True:
        delete_messages = [pop_conn.dele(i) for i in range(1, len(pop_conn.list()[1]) + 1)]
    pop_conn.quit()
    return messages

def getAttachments(pop_conn):
    allowed_mimetypes = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","image/png","image/jpg", "image/jpeg"]
    messages = fetchMail(pop_conn)
    attachments = []
    for msg in messages:
        #print msg
        createJSON(425, msg)

    return attachments

#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################

baseurl = 'https://10.100.107.61'

#pop_conn = poplib.POP3_SSL(configParser.get('env', 'servername'))
#pop_conn.user(configParser.get('env', 'email'))
#pop_conn.pass_(configParser.get('env', 'emailpassword'))

#getAttachments(pop_conn)
#print (pop_conn.getwelcome())

sessionToken = getSessionToken(baseurl, configParser.get('env', 'archerusername'), configParser.get('env', 'archerinstancename'), configParser.get('env', 'archerpassword'))

hola = createJSON(configParser.get('env', 'moduleIdTickets'))
postContent(baseurl, sessionToken, hola)
