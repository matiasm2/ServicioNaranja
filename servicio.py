import requests, poplib, ConfigParser, sys
from email import parser

configParser = ConfigParser.RawConfigParser()
#configFilePath = sys.argv[1]
configFilePath = "./env.cfg"
configParser.read(configFilePath)

ticketsIds = {
    "levelId": 369,
    "asunto": 30542,
    "cc": 30670,
    "de": 30663,
    "ddc": 31876,
    "para": 30664
}
ddcIds = {
    "levelId": 375,
    "adjunto": 31881,
    "asunto": 31880,
    "cuerpo": 31884,
    "de" : 31889,
    "para" : 31890,
    "cc" : 31891,
    "obs": 31893,
    "paramail": 31896,
    "ccmail": 31895
}


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
    print response.json()
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
    print response.json()
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
    print response.json()
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
    #print response
    print response.json()
    return response

def createJSONDetalle(msg, subject, fromm, to, tomail, cc, ccmail, obs):
    levelId = ddcIds['levelId']
    detalle = ddcIds['cuerpo']
    asunto = ddcIds['asunto']
    de = ddcIds['de']
    para = ddcIds['para']
    ccc = ddcIds['cc']
    obss = ddcIds['obs']
    mailpara = ddcIds['paramail']
    mailcc = ddcIds['ccmail']

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
                    str(ccc):{
                      "Type" : 1,
                      "Value" : cc,
                      "FieldId" : str(ccc)
                    },
                    str(mailpara):{
                      "Type" : 1,
                      "Value" : tomail,
                      "FieldId" : str(mailpara)
                    },
                    str(mailcc):{
                      "Type" : 1,
                      "Value" : ccmail,
                      "FieldId" : str(mailcc)
                    },
                    str(obss):{
                      "Type" : 1,
                      "Value" : obs,
                      "FieldId" : str(obss)
                    }
             }
         }
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
    levelId = ticketsIds['levelId']
    ddc = ticketsIds['ddc']
    asunto = ticketsIds['asunto']
    cc = ticketsIds['cc']
    de = ticketsIds['de']
    para = ticketsIds['para']




    for part in msg.walk():
        #print part.get_content_type()
        #print part.get_content_type()
        print '\n'
        print  part.get_content_type()
        print part.get_payload(decode=False)
        obs = ''
        if part.get_content_type() in allowed_mimetypes:
            name = part.get_filename()
            name = name.decode(charset, 'replace')
            if '?' in name:
                name = name.split('?')[3]
            print name
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
        if part.get_content_type() not in allowed_mimetypes and part.get_content_type() not in ["text/html", "text/plain", "multipart/alternative", "multipart/mixed"]:
            obs += 'Se encontro un adjunto ' + part.get_content_type() + ' que no se pudo cargar.'

        if part.get_content_type() == "text/html":
             charset = part.get_content_charset('iso-8859-1')
             body = part.get_payload(decode = True)
             texto = body.decode(charset, 'replace')
             valcc = ''
             valmailcc = ''
             if msg['CC']:
                valcc = ",".join(getEmails(msg['CC'])['names'])
                valmailcc =  ";".join(getEmails(msg['to'])['emails'])
             json = createJSONDetalle(texto, msg['subject'], ",".join(getEmails(msg['from'])['names']), ",".join(getEmails(msg['to'])['names']),  ";".join(getEmails(msg['to'])['emails']), valcc, valmailcc, obs)
             subformid = postContent(baseurl, sessionToken, json).json()['RequestedObject']['Id']
             subforms.append({"ContentID": subformid})
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
                    str(ddc): {
                        "Type" : 9,
                        "Value" : subforms,
                         "FieldId": str(ddc)
                     },
                    str(asunto): {
                       "Type" : 1,
                       "Value" : valasunto,
                        "FieldId": str(asunto)

                    },
                    str(cc):{
                       "Type" : 1,
                       "Value" : valcc,
                        "FieldId": str(cc)
                    },
                    str(de):{
                       "Type": 1,
                       "Value" : valde,
                       "FieldId": str(de)
                    },
                    str(para):{
                       "Type" : 1,
                       "Value" : valpara,
                       "FieldId" : str(para)
                    }

                }
            }
        }
    return [data, subforms, values]

def createJSONSubForm(ddcId, attachments):
    levelId = ddcIds['levelId']
    adjunto = ddcIds['adjunto']
    values = attachments

    data = {
        "Content":{
            "Id": ddcId ,
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
    levelId = ddcIds['levelId']
    adjunto = ddcIds['adjunto']
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
    levelId = ticketsIds['levelId']
    ddc = ticketsIds['ddc']


    registro = getContentById(baseurl, sessionToken, contentId)
    subforms = registro.json()['RequestedObject']['FieldContents'][str(ddc)]['Value']
    print subforms
    if subforms:
        subforms.append({"ContentID": subformId})
    else:
        subforms = [{"ContentID": subformId}]

    data = {
        "Content":{
            "Id": contentId ,
            "LevelId" : levelId,
            "FieldContents" : {
                str(ddc): {
                    "Type" : 9,
                    "Value" : subforms,
                     "FieldId": str(ddc)
                 }
                 }
             }
         }

    print data
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
    print response.json()
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
                contentId = isSuccessful['RequestedObject']['Id']
                print isSuccessful

                isSuccessful = isSuccessful['IsSuccessful']

                for subfid in info[1]:
                    subform = createJSONSubForm(subfid['ContentID'],  info[2])
                    req = putContent(baseurl, sessionToken, subform)
                    isSuccessful = req.json()['IsSuccessful']


            if isSuccessful:

                    deleteMail(msg)

        except KeyError:
            print "Ocurrio KeyError:"
    '''if False is not in isSuccessful:
        fetchMail(pop_conn, delete_after = true)'''



def updateRecords(contentId, moduleId, msg):
    values = []
    isSuccessful = False
    ddc = ticketsIds["ddc"]
    obs = ''

    for part in msg.walk():
        print '\n'
        print  part.get_content_type()
        print part.get_payload(decode=False)
        if part.get_content_type() in allowed_mimetypes:
            name = part.get_filename()
            if '?' in name:
                name = name.split('?')[3]
            print name
            data= part.get_payload(decode=False)
            print name
            iD = postAttachment(baseurl, sessionToken, name, data)
            #json = createJSONAttachment(baseurl, contentId, moduleId, iD)
            #putContent(baseurl, sessionToken, json)

            values.append(iD)

        if part.get_content_type() not in allowed_mimetypes and part.get_content_type() not in ["text/html", "text/plain", "multipart/alternative", "multipart/mixed"]:
            obs += 'Se encontro un adjunto ' + part.get_content_type() + ' que no se pudo cargar.'

        if part.get_content_type() == "text/html":
             charset = part.get_content_charset('iso-8859-1')
             body = part.get_payload(decode = True)
             texto = body.decode(charset, 'replace')
             valcc = ''
             if msg['CC']:
                valcc = ",".join(getEmails(msg['CC'])['names'])
                valmailcc =  ";".join(getEmails(msg['to'])['emails'])
             json = createJSONDetalle(texto, msg['subject'], ",".join(getEmails(msg['from'])['names']), ",".join(getEmails(msg['to'])['names']),  ";".join(getEmails(msg['to'])['emails']), valcc, valmailcc, obs)
             subformId = postContent(baseurl, sessionToken, json).json()['RequestedObject']['Id']


             json2 = createJSONMail(baseurl, contentId, moduleId, subformId)
             putContent(baseurl, sessionToken, json2)

    subformJSON = createJSONSubForm(subformId, values)
    print subformId
    print subformJSON
    req = putContent(baseurl, sessionToken, subformJSON)
    isSuccessful = req.json()['IsSuccessful']
    print req.json()
    return isSuccessful

def getEmails(emailsstring):
    emails = {
        "names": [],
        "emails": []
    }

    if ',' in emailsstring:
        emailsstring = emailsstring.split(",")
        for emailstring in emailsstring:
            emailaux = emailstring.split("<")
            emails['names'].append(emailaux[0])
            emails['emails'].append(emailaux[1].split(">")[0])
    else:
        emailaux = emailsstring.split("<")
        emails['names'].append(emailaux[0])
        emails['emails'].append(emailaux[1].split(">")[0])
    print emails
    return emails

def apiCall(url, sessionToken, content):
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



#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################

baseurl = configParser.get('env', 'archerurl')
allowed_mimetypes = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/msword', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/pdf', 'text/csv', 'image/png', 'image/jpeg', 'image/gif', 'application/x-zip-compressed', 'application/octet-stream']

pop_conn = createPOPConn()

#getAttachments(pop_conn)
#print (pop_conn.getwelcome())

sessionToken = getSessionToken(baseurl, configParser.get('env', 'archerusername'), configParser.get('env', 'archerinstancename'), configParser.get('env', 'archerpassword'))

createRecord(pop_conn)
#print getContentById(baseurl, sessionToken, 240420).json()
