from arcgis.gis import GIS
import arcgis
import datetime
import pandas as pd

username = 'AGOLUSERNAME'
password = 'AGOLPASSWORD'
gis = GIS("https://www.arcgis.com", username, password)

# the id of the feature service resulting from the Survey.
# This is currently set to the R1_V4_FSVeg_Spatial_Walkthrough Form
layer = gis.content.get('b4508d233a3e4d698243920b80cb7e00')

outputFolder = r'C:\Your\Folder'
current_date = datetime.date.today()

sharing = layer.shared_with
sharing = sharing['groups']

emailList = []
for sharedGroup in sharing:
    group = arcgis.gis.Group(gis, sharedGroup.id)
    members = group.get_members()
    members = members['users']
    for member in members:
        R1Users = gis.users.search(member)
        if R1Users[0].email not in emailList:
            emailList.append(R1Users[0].email)

finalDataFrame = pd.DataFrame(emailList, columns=['User Email'])

finalDataFrame.to_excel(
    '{0}/NEPA_Webmap_Shared_Users_{1}{2}{3}.xlsx'.format(
        outputFolder, current_date.strftime('%m'),
        current_date.strftime('%d'), current_date.year), index=False)
