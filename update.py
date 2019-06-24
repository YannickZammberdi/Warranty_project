from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
GoogleAuth.DEFAULT_SETTINGS['client_config_file'] = "C:\warranty_book\data\client_secrets.json"
def update():
    gauth = GoogleAuth()
    drive = GoogleDrive(gauth)
    file_list = drive.ListFile({'q':"'1FYVfJik6SxQv_1RIreXJjCh38fLGdtAY' in parents and trashed=false",
                    'corpora': 'teamDrive', 'teamDriveId': '0ALUZBWQEt-b2Uk9PVA',
                    'includeTeamDriveItems': True, 'supportsTeamDrives': True}).GetList()
    for file1 in file_list:
        if "." in file1['title']:
            file1.GetContentFile("C:/warranty_book/data/"+file1['title'])
        else:
            sub_list = drive.ListFile({'q':"'" + file1['id'] + "' in parents and trashed=false",
                    'corpora': 'teamDrive', 'teamDriveId': '0ALUZBWQEt-b2Uk9PVA',
                    'includeTeamDriveItems': True, 'supportsTeamDrives': True}).GetList()
            for file2 in sub_list:
                file2.GetContentFile("C:/warranty_book/data/" + file1['title'] + "/" +file2['title'])
                print('title: %s, id: %s' % (file2['title'], file2['id']))
        print('title: %s, id: %s' % (file1['title'], file1['id']))