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
    #print response
    #print response.json()['RequestedObject']['Id']
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
    #print response
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
    #print response
    #print response.json()
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
    #print response
    #print response.json()
    return response

def createJSONDetalle(moduleId, subformFieldId, msg, subject, fromm, to, cc):
    fd = getApplicationFieldDefinition(baseurl, sessionToken, moduleId)

    for field in fd.json():
        if field['RequestedObject']['LevelId']:
                levelId = field['RequestedObject']['LevelId']
        if field['RequestedObject']['Alias'] == 'Detalle_del_correo':
            detalle = field['RequestedObject']['Id']
        if field['RequestedObject']['Alias'] == 'Asunto':
            asunto = field['RequestedObject']['Id']
        if field['RequestedObject']['Alias'] == 'De':
            de = field['RequestedObject']['Id']
        if field['RequestedObject']['Alias'] == 'Para':
            para = field['RequestedObject']['Id']
        if field['RequestedObject']['Alias'] == 'CC':
            ccId = field['RequestedObject']['Id']

    data = {
        "Content":{
            "LevelId" : levelId,
            "FieldContents" : {
                 str(detalle): {
                     "Type" : 1,
                     "Value" : msg,
                      "FieldId": str(detalle)
                  },
                  str(asunto): {
                      "Type" : 1,
                      "Value" : subject,
                       "FieldId": str(asunto)
                   },
                   str(de): {
                       "Type" : 1,
                       "Value" : fromm,
                        "FieldId": str(de)
                    },
                    str(para): {
                        "Type" : 1,
                        "Value" : to,
                         "FieldId": str(para)
                     },
                     str(ccId): {
                         "Type" : 1,
                         "Value" : cc,
                          "FieldId": str(ccId)
                      }
             }
         },
         "SubformFieldId": str(subformFieldId)
     }
    #print data
    return data


def createJSON(moduleId, msg):
    subforms = []
    values = []
    valdetalle = ''
    valasunto = msg['subject']
    valde = msg['from']
    valpara = msg['to']
    valcc = ''
    if msg['CC']:
        valcc = msg['CC']
        #print valcc
    fd = getApplicationFieldDefinition(baseurl, sessionToken, moduleId)

    for field in fd.json():
        if field['RequestedObject']['LevelId']:
                levelId = field['RequestedObject']['LevelId']
        if field['RequestedObject']['Alias'] == 'Texto':
            detalle = field['RequestedObject']['Id']
        if field['RequestedObject']['Alias'] == 'Adjunto':
            adjunto = field['RequestedObject']['Id']
        if field['RequestedObject']['Alias'] == 'Asunto':
            asunto = field['RequestedObject']['Id']
        if field['RequestedObject']['Alias'] == 'De':
            de = field['RequestedObject']['Id']
        if field['RequestedObject']['Alias'] == 'Para':
            para = field['RequestedObject']['Id']
        if field['RequestedObject']['Alias'] == 'CC':
            cc = field['RequestedObject']['Id']
        if field['RequestedObject']['Alias'] == 'Detalle_del_correo':
            ddc = field['RequestedObject']['Id']


    for part in msg.walk():
        #print part.get_content_type()
        #print part.get_content_type()
        if part.get_content_type() in allowed_mimetypes:

            name = part.get_filename()
            #print name
            data= part.get_payload(decode=False)
            #f = file('archivines/'+str(name),'wb')
            #f.write(data)
            #f.close()
            iD = postAttachment(baseurl, sessionToken, name, data)
            values.append(iD)
            print iD

            #json = createJSONUpdate(baseurl, sessionToken, 239965, levelId, fieldId, iD)
            #putContent(baseurl, sessionToken, json)

            #attachments.append(name)
        if part.get_content_type() == "text/html":
             body = part.get_payload(decode = True)
             charset = part.get_content_charset('iso-8859-1')
             texto = body.decode(charset, 'replace')
             valcc = ''
             if msg['CC']:
                 valcc = msg['CC']
             json = createJSONDetalle(configParser.get('env', 'moduleIdDetalle'), ddc, texto, msg['subject'], msg['from'], msg['to'], valcc)
             subformid = postContent(baseurl, sessionToken, json).json()['RequestedObject']['Id']
             subforms.append(subformid)
             valdetalle += texto
             #print valdetalle

    #print data
    '''
    for subfid in subforms:
        subformJSON = createJSONSubForm(subfid, configParser.get('env', 'moduleIdDetalle'), values)
        print subformJSON
        isSuccessful = putContent(baseurl, sessionToken, subformJSON)
        print isSuccessful.json()'''

    data = {
       "Content":{
           "LevelId" : levelId,
           "FieldContents" : {
                str(detalle): {
                    "Type" : 1,
                    "Value" : valdetalle,
                     "FieldId": str(detalle)
                 },
                 str(asunto): {
                     "Type" : 1,
                     "Value" : valasunto,
                      "FieldId": str(asunto)
                  },
                  str(cc): {
                      "Type" : 1,
                      "Value" : valcc,
                       "FieldId": str(cc)
                   },
                  str(de): {
                      "Type" : 1,
                      "Value" : valde,
                       "FieldId": str(de)
                   },
                   str(para): {
                       "Type" : 1,
                       "Value" : valpara,
                        "FieldId": str(para)
                    },
                    str(ddc): {
                        "Type" : 24,
                        "Value" : subforms,
                         "FieldId": str(ddc)
                     }

                }
            }
        }
    return [data, subforms, values]

def createJSONSubForm(subformId, moduleId, attachments):
    fd = getApplicationFieldDefinition(baseurl, sessionToken, moduleId)
    #print fd.json()
    for field in fd.json():
        if field['RequestedObject']['LevelId']:
                levelId = field['RequestedObject']['LevelId']
        if field['RequestedObject']['Alias'] == 'Adjunto':
            adjunto = field['RequestedObject']['Id']

    values = attachments

    data = {
        "Content":{
            "Id": subformId ,
            "LevelId" : levelId,
            "FieldContents" : {
                str(adjunto): {
                    "Type" : 11,
                    "Value" : values,
                    "FieldId": adjunto
                 }
            }
         }
     }
    return data


def createJSONAttachment(baseurl, contentId, moduleId, value):
    fd = getApplicationFieldDefinition(baseurl, sessionToken, moduleId)

    for field in fd.json():
        if field['RequestedObject']['LevelId']:
                levelId = field['RequestedObject']['LevelId']
        if field['RequestedObject']['Alias'] == 'Adjunto':
            adjunto = field['RequestedObject']['Id']
    registro = getContentById(baseurl, sessionToken, contentId)
    values = registro.json()['RequestedObject']['FieldContents'][str(adjunto)]['Value']
    if values:
        values.append(value)
    else:
        values = [value]




    data = {
        "Content":{
            "Id": contentId ,
            "LevelId" : levelId,
            "FieldContents" : {
                str(adjunto): {
                    "Type" : 11,
                    "Value" : values,
                     "FieldId": adjunto
                 }
                 }
             }
         }
    return data

def createJSONMail(baseurl, contentId, moduleId, subformId):
    fd = getApplicationFieldDefinition(baseurl, sessionToken, moduleId)

    for field in fd.json():
        if field['RequestedObject']['LevelId']:
                levelId = field['RequestedObject']['LevelId']
        if field['RequestedObject']['Alias'] == 'Detalle_del_correo':
            ddc = field['RequestedObject']['Id']

    registro = getContentById(baseurl, sessionToken, contentId)
    subforms = registro.json()['RequestedObject']['FieldContents'][str(ddc)]['Value']
    if subforms:
        subforms.append(subformId)
    else:
        subforms = [subformId]

    data = {
        "Content":{
            "Id": contentId ,
            "LevelId" : levelId,
            "FieldContents" : {
                str(ddc): {
                    "Type" : 24,
                    "Value" : subforms,
                     "FieldId": str(ddc)
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
    #print response.json()
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
    #print response
    return response

#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################
def createPOPConn():
    pop_conn = poplib.POP3_SSL(configParser.get('env', 'servername'))
    pop_conn.user(configParser.get('env', 'email'))
    pop_conn.pass_(configParser.get('env', 'emailpassword'))
    return pop_conn

def deleteMail(msg):
    isSuccessful = False
    pop_conn = createPOPConn()
    for i in range(1, len(pop_conn.list()[1]) + 1):
        mssg = pop_conn.retr(i)
        mssg = "\n".join(mssg[1])
        mssg = parser.Parser().parsestr(mssg)
        if mssg['Message-ID'] == msg['Message-ID'] and mssg['date'] == msg['date']:
            pop_conn.dele(i)

    pop_conn.quit()

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

def createRecord(pop_conn):
    #allowed_mimetypes = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","image/png","image/jpg", "image/jpeg"]
    messages = fetchMail(pop_conn)

    attachments = []
    print len(messages)

    for msg in messages:
        #print msg
        isSuccessful = False

        try:
            if 'Ticket#' in msg['subject']:
                contentId = msg['subject'].split("#")[1]
                if ' ' in msg['subject']:
                    contentId = msg['subject'].split("#")[1].split(' ')[0]
                    print contentId

                if  getContentById(baseurl, sessionToken, contentId).json()['IsSuccessful']:
                    isSuccessful = updateRecords(contentId, configParser.get('env', 'moduleIdTickets'), msg)

            else:
                info = createJSON(configParser.get('env', 'moduleIdTickets'), msg)
                print info
                isSuccessful = postContent(baseurl, sessionToken, info[0]).json()
                print isSuccessful
                isSuccessful = isSuccessful['IsSuccessful']

                for subfid in info[1]:
                    subform = createJSONSubForm(subfid, configParser.get('env', 'moduleIdDetalle'), info[2])
                    req = putContent(baseurl, sessionToken, subform)
                    isSuccessful = req.json()
                    print isSuccessful
                    isSuccessful = isSuccessful['IsSuccessful']


            if isSuccessful:
                    deleteMail(msg)

        except KeyError:
            print "Ocurrio KeyError:"
    '''if False is not in isSuccessful:
        fetchMail(pop_conn, delete_after = true)'''


def updateRecords(contentId, moduleId, msg):
    values = []
    fd = getApplicationFieldDefinition(baseurl, sessionToken, moduleId)
    isSuccessful = False

    for field in fd.json():
        if field['RequestedObject']['Alias'] == 'Detalle_del_correo':
            ddc = field['RequestedObject']['Id']

    for part in msg.walk():
        if part.get_content_type() in allowed_mimetypes:
            name = part.get_filename()
            data= part.get_payload(decode=False)
            iD = postAttachment(baseurl, sessionToken, name, data)
            #json = createJSONAttachment(baseurl, contentId, moduleId, iD)
            #putContent(baseurl, sessionToken, json)

            values.append(iD)


        if part.get_content_type() == "text/html":
             body = part.get_payload(decode = True)
             charset = part.get_content_charset('iso-8859-1')
             texto = body.decode(charset, 'replace')
             valcc = ''
             if msg['CC']:
                 valcc = msg['CC']
             json = createJSONDetalle(configParser.get('env', 'moduleIdDetalle'), ddc, texto, msg['subject'], msg['from'], msg['to'], valcc)
             subformId = postContent(baseurl, sessionToken, json).json()['RequestedObject']['Id']


             json2 = createJSONMail(baseurl, contentId, moduleId, subformId)
             putContent(baseurl, sessionToken, json2)

    subformJSON = createJSONSubForm(subformId, configParser.get('env', 'moduleIdDetalle'), values)
    print subformId
    print subformJSON
    req = putContent(baseurl, sessionToken, subformJSON)
    isSuccessful = req.json()['IsSuccessful']
    print req.json()
    return isSuccessful

#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################

baseurl = configParser.get('env', 'archerurl')
allowed_mimetypes = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/msword', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/pdf', 'text/csv', 'image/png', 'image/jpeg', 'image/gif']

pop_conn = createPOPConn()

#getAttachments(pop_conn)
#print (pop_conn.getwelcome())

sessionToken = getSessionToken(baseurl, configParser.get('env', 'archerusername'), configParser.get('env', 'archerinstancename'), configParser.get('env', 'archerpassword'))

createRecord(pop_conn)
#print getContentById(baseurl, sessionToken, 240420).json()
