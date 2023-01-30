
import tokenHelper
import requests
import json
import pandas as pd
import os
import zipfile
import createBearerToken as Oauth

rootDir = "c:\\DEMO\\1\\"
workspace = "ws001"

# Set up session
session = requests.session()

bearerToken = tokenHelper.getToken()

if len(bearerToken) > 2:

    # Header for the MetaData File
    file = open(rootDir + '\\' + "files.csv", 'w')
    file.writelines("Document Number, ResponsibleUser, Disipline, DateCreated, Stage, Type\r")
    file.close()

    session.headers = {'Content-Type': 'application/json', 'Accept': 'text/plain',
                       'Authorization': f"Bearer {bearerToken}"}

    # Get the list of approved documents to export
    r = session.get(f'{Oauth.getBaseUrl()}/odata/documentsdata?$orderby=DateModified&$top=500')
    df = pd.DataFrame(json.loads(r.text))

    dfDocuments = pd.json_normalize(df['value'])

    for index, row in dfDocuments.iterrows():
        # access data using column names
        sDocumentNumber = row['DocumentNumber']
        # Now get the Static DocumentId to Download
        # This also has Document Number, ResponsibleUser, Disipline, DateCreated, Stage, Type
        r = session.get(f'{Oauth.getBaseUrl()}/api/{workspace}/documents/search?cap=500&format=group&query=' + sDocumentNumber)
        df = pd.DataFrame(json.loads(r.text))
        dfSearchResults = pd.json_normalize(df['Documents'])

        for indexSrch, rowSrch in dfSearchResults.iterrows():
            # access data using column names
            jsonStr2 = rowSrch[0]
            if len(jsonStr2):
                filenumber = jsonStr2[0]['NavigationId']
                if sDocumentNumber.replace(" ", "") == jsonStr2[0]['Name'].replace(" ", ""):
                    sVersionFolder = rootDir + '\\' + str(jsonStr2[0]['Name'])
                    sVersion = str(jsonStr2[0]['Version'])
                    isExist = os.path.exists(sVersionFolder)
                    if (isExist != True):
                        # Document Number, ResponsibleUser, Disipline, DateCreated, Stage, Type
                        print(sDocumentNumber + "," + rowSrch[13] + "," + rowSrch[9] + "," + rowSrch[10]  + "," + rowSrch[20] + "," + rowSrch[21])
                        file = open(rootDir + '\\' + "files.csv", 'a')
                        file.writelines(sDocumentNumber + "," + rowSrch[13] + "," + rowSrch[9] + "," + rowSrch[10]  + "," + rowSrch[20] + "," + rowSrch[21] + "\r")
                        file.close()
                        # Save File
                        os.mkdir(sVersionFolder)
                        r = session.get(f'{Oauth.getBaseUrl()}/api/workspaces/' + row['WorkspaceSlug'] + '/documents/' + str(
                            filenumber) + '/export/full?fileFormat=word&includeAttachments=true&includeIssues=true')

                        local_filename = sVersionFolder + '\\' + sDocumentNumber + '.zip'
                        local_filename_doc = sVersionFolder + '\\' + sDocumentNumber + '.docx'

                        totalbits = 0
                        if r.status_code == 200:
                            with open(local_filename, 'wb') as f:
                                for chunk in r.iter_content(chunk_size=1024):
                                    if chunk:
                                        totalbits += 1024
                                        f.write(chunk)
                        # Unzip the File
                        with zipfile.ZipFile(local_filename, "r") as zip_ref:
                            zip_ref.extractall(sVersionFolder)

