#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests, poplib, ConfigParser, sys
from email import parser
from email.header import decode_header

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
    "asunto": 24100,
    "cc": 24095,
    "de": 22508,
    "ddc": 23864,
    "para": 24094,
    "realmentececticket": 24066,
    "rticketno": 72098,
    "trabierto": 24111,
    "trabiertosi": 72137,
    "cerrarticket": 23854,
    "cticketno": 72100,
    "estado": 23852,
    "estadocerrado": 71959,
    "reabrirt": 23852,
    "rticketsi": 72129,
    "delacola": 17022,
    "servicios": 66631,
    "soc": 72123,
    "instalacion": 66628
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
    ##print response
    #print response.json()
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
    #print response.json()
    #print response.json()

def getApplicationMetadata(url, sessionToken, applicationId):
    url = url+'/api/core/system/application/' + str(applicationId)
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Authorization':'Archer session-id='+sessionToken,
        'Content-Type': 'application/json',
        'X-Http-Method-Override': 'GET'
    }
    response = requests.post(url, verify=False, headers=headers)
    ##print response
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
    ##print response
    #print response.json()
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
    ##print data
    return data


def createJSON(moduleId, msg):
    subforms = []
    values = []
    valasunto = decode_header(msg['subject'])[0][0].decode("iso-8859-1")
    decola = getEmails(msg['from'])['emails']
    paracola = []
    cccola = []
    valde = ",".join(getEmails(msg['from'])['emails'])
    valpara = 'pruebanova@outlook.com.ar'
    valcc = 'pruebanova@outlook.com.ar'
    valmailcc = ''
    if msg['TO']:
      valpara = ",".join(getEmails(msg['to'])['emails'])
      paracola = getEmails(msg['to'])['emails']
      tonames = valpara
      toemails = ";".join(getEmails(msg['to'])['emails'])
    if msg['CC']:
        cccola = getEmails(msg['cc'])['emails']
        valcc = ",".join(getEmails(msg['CC'])['emails'])
        valmailcc =  ";".join(getEmails(msg['CC'])['emails'])
    levelId = ticketsIds['levelId']
    ddc = ticketsIds['ddc']
    asunto = ticketsIds['asunto']
    cc = ticketsIds['cc']
    de = ticketsIds['de']
    para = ticketsIds['para']
    delacola = ticketsIds['delacola']
    hasHTML = False
    body = ''
    obs = ''
    attNames = []
    attBase64 = []
    for part in msg.walk():

        #print part.get_content_type()
        if part.get_content_type() in allowed_mimetypes:
            name = part.get_filename()
            name = name.decode(charset, 'replace')
            print name
            if '?' in name:
                name = name.split('?')[3].decode('iso-8859-1')
            data= part.get_payload(decode=False)
            #f = file('archivines/'+str(name),'wb')
            #f.write(data)
            #f.close()
            attNames.append(name)
            attBase64.append(part.get_payload().replace('\n',''))
            iD = postAttachment(baseurl, sessionToken, name, data)
            if iD['IsSuccessful']:
                values.append(iD['RequestedObject']['Id'])
            else:
                obs += iD['ValidationMessages'][0]['ResourcedMessage']
            #print iD

            #json = createJSONUpdate(baseurl, sessionToken, 239965, levelId, fieldId, iD)
            #putContent(baseurl, sessionToken, json)

            #attachments.append(name)
        if part.get_content_type() not in allowed_mimetypes and part.get_content_type() not in ["text/html", "text/plain", "multipart/alternative", "multipart/mixed"]:
            obs += 'Se encontro un adjunto ' + part.get_content_type() + ' que no se pudo cargar.'
        if part.get_content_type() == "text/plain":
            #print part.get_payload()

            if part.get_payload().endswith('='):
                texto = part.get_payload().decode('base64')
                body = texto
            else:
                body = part.get_payload()
        if part.get_content_type() == "text/html":
            hasHTML=True
            charset = part.get_content_charset('iso-8859-1')
            body = part.get_payload(decode = True)
            texto = body.decode(charset, 'replace')
            body = texto
            json = createJSONDetalle(texto, decode_header(msg['subject'])[0][0].decode("iso-8859-1"), ",".join(getEmails(msg['from'])['emails']), tonames, toemails, valcc, valmailcc, obs)
            subformid = postContent(baseurl, sessionToken, json).json()['RequestedObject']['Id']
            subformJSON = createJSONSubForm(subformid, values, body)
            req = putContent(baseurl, sessionToken, subformJSON)
            subforms.append({"ContentID": subformid})
    if not hasHTML:
        json = createJSONDetalle(body, decode_header(msg['subject'])[0][0].decode("iso-8859-1"), ",".join(getEmails(msg['from'])['emails']), tonames, toemails, valcc, valmailcc, obs)
        subformid = postContent(baseurl, sessionToken, json).json()['RequestedObject']['Id']
        subforms.append({"ContentID": subformid})

    '''for subfid in subforms:
        subformJSON = createJSONSubForm(subfid, configParser.get('env', 'moduleIdDetalle'), values)
        #print subformJSON
        isSuccessful = putContent(baseurl, sessionToken, subformJSON)
        #print isSuccessful.json()'''

    for name in attNames:
        print name
        if name in body:
            i = attNames.index(name)
            cid = body.index('cid:'+name)+len('cid:'+name)
            st = 'cid:'+name+body[cid:cid+18]
            body = body.replace(st, 'http://novatickets/'+name.split('.')[1]+attBase64[i])
    print body
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
                       "Value" : valmailcc,
                        "FieldId": str(cc)
                    },
                    str(de):{
                       "Type": 1,
                       "Value" : valde,
                       "FieldId": str(de)
                    },
                    str(para):{
                       "Type" : 1,
                       "Value" : toemails,
                       "FieldId" : str(para)
                    },
                    str(delacola): {
                        "Type" : 4,
                        "Value" : {
                            'ValuesListIds': [],
                            'OtherText': None
                        },
                        "FieldId": str(delacola)
                    }

                }
            }
    }

    if emailservicios in decola or emailservicios in paracola or emailservicios in cccola:
        data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [ticketsIds['servicios']]
    if emailsoc in decola or emailsoc in paracola or emailsoc in cccola:
        data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [ticketsIds['soc']]
    if emailinstalacion in decola or emailinstalacion in paracola or emailinstalacion in cccola:
        data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [ticketsIds['instalacion']]

    return [data, subforms, values, body]

def createJSONSubForm(ddcId, attachments, msg):
    levelId = ddcIds['levelId']
    adjunto = ddcIds['adjunto']
    detalle = ddcIds['cuerpo']
    values = attachments

    data = {
        "Content":{
            "Id": ddcId ,
            "LevelId" : levelId,
            "FieldContents" : {
                str(detalle): {
                     "Type" : 1,
                     "Value" : msg,
                      "FieldId": str(detalle)
                  },
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
    rticket = ticketsIds['realmentececticket']
    rticketno = ticketsIds['rticketno']
    trabierto = ticketsIds['trabierto']
    trabiertosi = ticketsIds['trabiertosi']
    cticket = ticketsIds['cerrarticket']
    cticketno = ticketsIds['cticketno']
    rrticket = ticketsIds['realmentececticket']
    rrticketsi = ticketsIds['rticketsi']
    estado = ticketsIds['estado']
    ecerrado = ticketsIds['estadocerrado']

    registro = getContentById(baseurl, sessionToken, contentId)
    #print registro.json()
    subforms = registro.json()['RequestedObject']['FieldContents'][str(ddc)]['Value']
    if subforms:
        subforms.append({"ContentID": subformId})
    else:
        subforms = [{"ContentID": subformId}]

    data = {
        "Content":{
            "Id": contentId,
            "LevelId" : levelId,
            "FieldContents" : {
                str(ddc): {
                    "Type" : 9,
                    "Value" : subforms,
                    "FieldId": str(ddc)
                },
                str(rticket): {
                    "Type" : 4,
                    "Value" : {
                        'ValuesListIds': [],
                        'OtherText': None
                    },
                    "FieldId": str(rticket)
                },
                str(cticket): {
                    "Type" : 4,
                    "Value" : {
                        'ValuesListIds': [],
                        'OtherText': None
                    },
                    "FieldId": str(cticket)
                },
                str(trabierto): {
                    "Type" : 4,
                    "Value" : {
                        'ValuesListIds': [],
                        'OtherText': None
                    },
                    "FieldId": str(trabierto)
                }

            }
        }
    }

    #print data['Content']['FieldContents'][str(rticket)]
    for val in registro.json()['RequestedObject']['FieldContents'][str(estado)]['Value']['ValuesListIds']:
        #print val
        #print ecerrado
        if val == ecerrado:
            data['Content']['FieldContents'][str(rticket)]['Value']['ValuesListIds']=[int(rticketno)]
            data['Content']['FieldContents'][str(cticket)]['Value']['ValuesListIds']=[int(cticketno)]
            data['Content']['FieldContents'][str(trabierto)]['Value']['ValuesListIds']=[int(trabiertosi)]

    #print data
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
    #print response.json()
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
def createPOPConn(servername, email, password):
    pop_conn = poplib.POP3_SSL(servername)
    pop_conn.user(email)
    pop_conn.pass_(password)
    return pop_conn

def deleteMail(msg, servername, email, password):
    pop_conn = createPOPConn(servername, email, password)
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

def createRecords(servername, email, password):
    #allowed_mimetypes = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","image/png","image/jpg", "image/jpeg"]
    messages = ""
    try:
        pop_conn = createPOPConn(servername, email, password)
        messages = fetchMail(pop_conn)
    except poplib.error_proto:
        print "Ocurrio un error con la cuenta "+email+" ¿contraseña erronea?"


    attachments = []
    #print len(messages)
    headers = getHeaders(sessionToken)
    isSuccessful = True
    for msg in messages:
        ##print msg
        if not isSuccessful:
            #print 'Fallo'
            break

        try:
            #print msg['subject']
            #print decode_header(msg['subject'])[0][0].decode("iso-8859-1")
            if 'Ticket#' in decode_header(msg['subject'])[0][0].decode("iso-8859-1"):
                contentId = decode_header(msg['subject'])[0][0].decode("iso-8859-1").split("#")[1]
                if ' ' in decode_header(msg['subject'])[0][0].decode("iso-8859-1"):
                    contentId = decode_header(msg['subject'])[0][0].decode("iso-8859-1").split("#")[1].split(' ')[0]
                    #print contentId

                if  getContentById(baseurl, sessionToken, contentId).json()['IsSuccessful']:
                    #print 'Actualiza uno'
                    isSuccessful = updateRecords(contentId, configParser.get('env', 'moduleIdTickets'), msg)


            else:
                info = createJSON(configParser.get('env', 'moduleIdTickets'), msg)
                isSuccessful = postContent(baseurl, sessionToken, info[0]).json()
                if isSuccessful['IsSuccessful']:
                    contentId = isSuccessful['RequestedObject']['Id']
                #print 'Crea uno nuevo'
                #print isSuccessful


                isSuccessful = isSuccessful['IsSuccessful']

                for subfid in info[1]:
                    subform = createJSONSubForm(subfid['ContentID'],  info[2], info[3])
                    req = putContent(baseurl, sessionToken, subform)
                    #print 'Crea un detalle nuevo'
                    #print req.json()
                    isSuccessful = req.json()['IsSuccessful']
            if isSuccessful:
                    deleteMail(msg, servername, email, password)

        except KeyError:
            print "Ocurrio KeyError:"
    #print isSuccessful
    return isSuccessful




def updateRecords(contentId, moduleId, msg):
    values = []
    subformId = ''
    isSuccessful = False
    ddc = ticketsIds["ddc"]
    obs = ''
    valcc = 'pruebanova@outlook.com.ar'
    valmailcc = ''
    tonames = 'pruebanova@outlook.com.ar'
    toemails = ''
    if msg['TO']:
        tonames = ",".join(getEmails(msg['to'])['emails'])
        toemails = ";".join(getEmails(msg['to'])['emails'])
    if msg['CC']:
       valcc = ",".join(getEmails(msg['CC'])['emails'])
       valmailcc =  ";".join(getEmails(msg['CC'])['emails'])

    hasHTML = False
    body = ''
    obs = ''
    attNames = []
    attBase64 = []

    for part in msg.walk():
        #print part.get_content_type()
        if part.get_content_type() in allowed_mimetypes:
            #print part
            name = part.get_filename()
            if name == None:
			   extension = part.get_content_type().split('/')[1]
			   name = "Adjunto" + "." + extension
			   #print name
            if '?' in name:
                name = name.split('?')[3]
            data= part.get_payload(decode=False)
            attNames.append(name)
            attBase64.append(part.get_payload().replace('\n',''))
            iD = postAttachment(baseurl, sessionToken, name, data)
            if iD['IsSuccessful']:
                values.append(iD['RequestedObject']['Id'])
            else:
                #print iD['ValidationMessages'][0]['ResourcedMessage']
                #print iD['ValidationMessages']
                obs += iD['ValidationMessages'][0]['ResourcedMessage']
                #print obs
            #print iD

        if part.get_content_type() not in allowed_mimetypes and part.get_content_type() not in ["text/html", "text/plain", "multipart/alternative", "multipart/mixed"]:
            obs += 'Se encontro un adjunto ' + part.get_content_type() + ' que no se pudo cargar.'
        if part.get_content_type() == "text/plain":
            #print part.get_payload()
            if part.get_payload().endswith('='):
                texto = part.get_payload().decode('base64')

                msg, subject, fromm, to, tomail, cc, ccmail, obs
                json = createJSONDetalle(texto, decode_header(msg['subject'])[0][0].decode("iso-8859-1"), ",".join(getEmails(msg['from'])['emails']), tonames, toemails, valcc, valmailcc, obs)
                subform = postContent(baseurl, sessionToken, json).json()
                subformId = subform['RequestedObject']['Id']
                json2 = createJSONMail(baseurl, contentId, moduleId, subformId)
                req = putContent(baseurl, sessionToken, json2)
                #print req.json()
                isSuccessful = req.json()['IsSuccessful']

        if part.get_content_type() == "text/html":
            hasHTML = True
            charset = part.get_content_charset('iso-8859-1')
            body = part.get_payload(decode = True)
            texto = body.decode(charset, 'replace')
            body = texto
            json = createJSONDetalle(texto, decode_header(msg['subject'])[0][0].decode("iso-8859-1"), ",".join(getEmails(msg['from'])['emails']), tonames, toemails, valcc, valmailcc, obs)
            subformId = postContent(baseurl, sessionToken, json).json()['RequestedObject']['Id']
            json2 = createJSONMail(baseurl, contentId, moduleId, subformId)
            req = putContent(baseurl, sessionToken, json2)
            isSuccessful = req.json()['IsSuccessful']
            #print req.json()

    if not hasHTML:
        json = createJSONDetalle(body, decode_header(msg['subject'])[0][0].decode("iso-8859-1"), ";".join(getEmails(msg['from'])['emails']), tonames, toemails, valmailcc, valmailcc, obs)
        subform = postContent(baseurl, sessionToken, json).json()
        subformId = subform['RequestedObject']['Id']
        json2 = createJSONMail(baseurl, contentId, moduleId, subformId)
        req = putContent(baseurl, sessionToken, json2)
        #print req.json()
        isSuccessful = req.json()['IsSuccessful']

    for name in attNames:
        print name
        if name in body:
            i = attNames.index(name)
            cid = body.index('cid:'+name)+len('cid:'+name)
            st = 'cid:'+name+body[cid:cid+18]
            print st
            body = body.replace(st, 'http://novatickets/'+name.split('.')[1]+attBase64[i])
    print body
    subformJSON = createJSONSubForm(subformId, values, body)
    ##print subformId
    req = putContent(baseurl, sessionToken, subformJSON)
    isSuccessful = req.json()['IsSuccessful']
    ##print req.json()
    return isSuccessful

def getEmails(emailsstring):
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
    #print response.json()

    return response



#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################

baseurl = configParser.get('env', 'archerurl')
allowed_mimetypes = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/msword', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/pdf', 'text/csv', 'image/png', 'image/jpeg', 'image/gif', 'application/x-zip-compressed']
servername = configParser.get('env', 'servername')
emailservicios =  configParser.get('env', 'emailservicios')
emailsoc =  configParser.get('env', 'emailsoc')
emailinstalacion =  configParser.get('env', 'emailinstalacion')
emailpassword = configParser.get('env', 'emailpassword')

data = {
    "InstanceName": configParser.get('env', 'archerinstancename'),
    "Username": configParser.get('env', 'archerusername'),
    "UserDomain": "",
    "Password": configParser.get('env', 'archerpassword')
}

sessionToken = apiCall(baseurl+'/api/core/security/login', getHeaders(), data). json()['RequestedObject']['SessionToken']

isSuccessful = createRecords(servername, emailservicios, emailpassword)
isSuccessful = createRecords(servername, emailsoc, emailpassword)
isSuccessful = createRecords(servername, emailinstalacion, emailpassword)

headers = getHeaders(sessionToken)

# Si todo esta OK se ejecutan los datafeeds de mapeo de contactos y merge de ticket.
if isSuccessful:
        guidDFContactos = configParser.get('env', 'GuidDFContactos')
        guidDFTicketMerge = configParser.get('env', 'GuidDFTicketMerge')
        dfContent = {"DataFeedGuid": guidDFContactos, "IsReferenceFeedsIncluded": False}
        apiCall(baseurl+'/api/core/datafeed/execution', headers, dfContent)
        dfContent = {"DataFeedGuid": guidDFTicketMerge, "IsReferenceFeedsIncluded": True}
        apiCall(baseurl+'/api/core/datafeed/execution', headers, dfContent)
##print getContentById(baseurl, sessionToken, 240420).json()
raw_input("Press Enter to continue...")
