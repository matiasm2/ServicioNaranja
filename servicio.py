import requests, poplib, ConfigParser, sys
from email import parser

configParser = ConfigParser.RawConfigParser()
#configFilePath = sys.argv[1]
configFilePath = "./env.cfg"
configParser.read(configFilePath)

'''ticketsIds = {
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
dataFeedGuid = "0C12DCCA-BD22-42B5-864A-8522A7DA2E82"'''

ticketsIds = {
    "levelId": 218, #Ultimo numero de url (/apps/ArcherApp/Home.aspx#search/70/75/542/false/default/368)
    "asunto": 16716,
    "cc": 22510,
    "de": 22508,
    "ddc": 23864,
    "para": 22509
}
ddcIds = {
    "levelId": 2266, #Ultimo numero de url (/apps/ArcherApp/Home.aspx#search/70/75/542/false/default/368)
    "adjunto": 23875,
    "asunto": 23874,
    "cuerpo": 23876,
    "de" : 23877,
    "para" : 23878,
    "cc" : 23879,
    "obs": 23881,
    "paramail": 23883,
    "ccmail": 23884
}
dataFeedGuid = "0C12DCCA-BD22-42B5-864A-8522A7DA2E82"


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
    return response.json()

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
    valde = ",".join(getEmails(msg['from'])['names'])
    valpara = ''
    valcc = ''
    if msg['TO']:
      valpara = ",".join(getEmails(msg['to'])['names'])
    if msg['CC']:
      valcc = ",".join(getEmails(msg['CC'])['names'])
    levelId = ticketsIds['levelId']
    ddc = ticketsIds['ddc']
    asunto = ticketsIds['asunto']
    cc = ticketsIds['cc']
    de = ticketsIds['de']
    para = ticketsIds['para']




    for part in msg.walk():
        obs = ''
        if part.get_content_type() in allowed_mimetypes:
            name = part.get_filename()
            name = name.decode(charset, 'replace')
            if '?' in name:
                name = name.split('?')[3]
            data= part.get_payload(decode=False)
            #f = file('archivines/'+str(name),'wb')
            #f.write(data)
            #f.close()
            iD = postAttachment(baseurl, sessionToken, name, data)
            if iD['IsSuccessful']:
                values.append(iD['RequestedObject']['Id'])
            else:
                obs += iD['ValidationMessages'][0]['ResourcedMessage']
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
             if msg['TO']:
               tonames = ",".join(getEmails(msg['to'])['names'])
               toemails = ";".join(getEmails(msg['to'])['emails'])
             if msg['CC']:
               valcc = ",".join(getEmails(msg['CC'])['names'])
               valmailcc =  ";".join(getEmails(msg['CC'])['emails'])
             json = createJSONDetalle(texto, msg['subject'], ",".join(getEmails(msg['from'])['names']), tonames,  valmailcc, valcc, toemails, obs)
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
    headers = getHeaders(sessionToken)
    isSuccessful = True
    for msg in messages:
        #print msg
        if not isSuccessful:
            print 'Fallo'
            break

        try:
            if 'Ticket#' in msg['subject']:
                contentId = msg['subject'].split("#")[1]
                if ' ' in msg['subject']:
                    contentId = msg['subject'].split("#")[1].split(' ')[0]
                    print contentId

                if  getContentById(baseurl, sessionToken, contentId).json()['IsSuccessful']:
                    print 'Actualiza uno'
                    isSuccessful = updateRecords(contentId, configParser.get('env', 'moduleIdTickets'), msg)


            else:
                info = createJSON(configParser.get('env', 'moduleIdTickets'), msg)
                isSuccessful = postContent(baseurl, sessionToken, info[0]).json()
                contentId = isSuccessful['RequestedObject']['Id']
                print 'Crea uno nuevo'
                print isSuccessful


                isSuccessful = isSuccessful['IsSuccessful']

                for subfid in info[1]:
                    subform = createJSONSubForm(subfid['ContentID'],  info[2])
                    req = putContent(baseurl, sessionToken, subform)
                    print 'Crea un detalle nuevo'
                    print req.json()
                    isSuccessful = req.json()['IsSuccessful']
            if isSuccessful:
                    deleteMail(msg)

        except KeyError:
            print "Ocurrio KeyError:"
    print isSuccessful
    if isSuccessful:
            dfContent = {"DataFeedGuid": dataFeedGuid, "IsReferenceFeedsIncluded": False}
            apiCall(baseurl+'/api/core/datafeed/execution', headers, dfContent)




def updateRecords(contentId, moduleId, msg):
    values = []
    isSuccessful = False
    ddc = ticketsIds["ddc"]
    obs = ''

    for part in msg.walk():
        obs = ''
        if part.get_content_type() in allowed_mimetypes:
            name = part.get_filename()
            if '?' in name:
                name = name.split('?')[3]
            data= part.get_payload(decode=False)
            iD = postAttachment(baseurl, sessionToken, name, data)
            if iD['IsSuccessful']:
                values.append(iD['RequestedObject']['Id'])
            else:
                print iD['ValidationMessages'][0]['ResourcedMessage']
                print iD['ValidationMessages']
                obs += iD['ValidationMessages'][0]['ResourcedMessage']
                print obs
            print iD

        if part.get_content_type() not in allowed_mimetypes and part.get_content_type() not in ["text/html", "text/plain", "multipart/alternative", "multipart/mixed"]:
            obs += 'Se encontro un adjunto ' + part.get_content_type() + ' que no se pudo cargar.'

        if part.get_content_type() == "text/html":
             charset = part.get_content_charset('iso-8859-1')
             body = part.get_payload(decode = True)
             texto = body.decode(charset, 'replace')
             valcc = ''
             valmailcc = ''
             tonames = ''
             toemails = ''
             if msg['TO']:
                 tonames = ",".join(getEmails(msg['to'])['names'])
                 toemails = ";".join(getEmails(msg['to'])['emails'])
             if msg['CC']:
                valcc = ",".join(getEmails(msg['CC'])['names'])
                valmailcc =  ";".join(getEmails(msg['CC'])['emails'])
             json = createJSONDetalle(texto, msg['subject'], ",".join(getEmails(msg['from'])['names']), tonames,  valmailcc, valcc, toemails, obs)
             subformId = postContent(baseurl, sessionToken, json).json()['RequestedObject']['Id']


             json2 = createJSONMail(baseurl, contentId, moduleId, subformId)
             req = putContent(baseurl, sessionToken, json2)
             print req

    subformJSON = createJSONSubForm(subformId, values)
    print subformId
    req = putContent(baseurl, sessionToken, subformJSON)
    isSuccessful = req.json()['IsSuccessful']
    print req.json()
    return isSuccessful

def getEmails(emailsstring):
    print emailsstring


    if '"' in emailsstring:
        emailsstring = ''.join(emailsstring.split('"'))
    if "\n" in emailsstring:
        emailsstring = ''.join(emailsstring.split('\n'))
    if "\t" in emailsstring:
        emailsstring = ''.join(emailsstring.split('\t'))
    emails = {
        "names": [],
        "emails": []
    }


    if ',' in emailsstring:
        emailsstring = emailsstring.split(",")
        for emailstring in emailsstring:
            emailaux = emailstring.split("<")
            emailaux[0] = emailaux[0].strip()
            if len(emailaux) == 1:
                    emails['names'].append(emailaux[0])
                    emails['emails'].append(emailaux[0])
            else:
                emails['names'].append(emailaux[0])
                emails['emails'].append(emailaux[1].split(">")[0])
    else:
        emailaux = emailsstring.split("<")
        emailaux[0] = emailaux[0].strip()
        if len(emailaux) == 1:
            emails['names'].append(emailaux[0])
            emails['emails'].append(emailaux[0])
        else:
            emails['names'].append(emailaux[0])
            emails['emails'].append(emailaux[1].split(">")[0])
    return emails

def getHeaders(sessionToken = ''):
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Content-Type': 'application/json'
    }

    if sessionToken != '':
        headers['Authorization'] = 'Archer session-id='+sessionToken

    return headers

def apiCall(url, headers, content):
    data = content
    headers = headers
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
