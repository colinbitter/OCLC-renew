import pandas as pd
from bookops_worldcat import WorldcatAccessToken
import glob
from bookops_worldcat import MetadataSession
from pathlib import Path
import os
import time

token = WorldcatAccessToken(
        key="InsertKey",
        secret="InsertSecret",
        scopes=["WorldCatMetadataAPI"],
        principal_id="InsertPrincipal_ID",
        principal_idns="InsertPrincipal_idns")

path1 = str(Path.home() / "Downloads")
allFiles = glob.glob(path1 + "/*.xlsx")   # file needs header
for file_ in allFiles:
    fileName = os.path.basename(file_)
    fileName = os.path.splitext(fileName)[0]
    data_df = pd.read_excel(file_, names=['MMSID', 'ALMAOCLC'], engine='openpyxl')

    df1 = data_df['ALMAOCLC'].tolist()

    dff3 = []
    dff4 = []

    with MetadataSession(authorization=token, timeout=20) as session:
        for x in df1:
            result = session.holding_get_status(oclcNumber=x)
            result2 = result.json()['content']
            df = pd.json_normalize(result2)
            dff1 = df['currentOclcNumber']
            dff2 = df['holdingCurrentlySet']
            for i in dff1:
                dff3.append(i)
            df4 = pd.DataFrame(dff3)
            for i in dff2:
                dff4.append(i)
            df5 = pd.DataFrame(dff4)

    data_df['BATCH'] = df4
    data_df['BATCH'] = pd.to_numeric(data_df['BATCH'])
    data_df['BATCHMATCH'] = (data_df['ALMAOCLC'] == data_df['BATCH'])
    data_df['HELD'] = df5
    allTrue = data_df[(data_df['BATCHMATCH'] == True) & (data_df['HELD'] == True)].index
    data_df.drop(allTrue, inplace=True)
    if data_df.empty is False:
        data_df.to_csv(path1 + '/AlmaOclcResult_{}.csv'.format(fileName), index=False)
    if data_df.empty is True:
        print(fileName + ' is clear')
    time.sleep(130)
c = input("press close to exit")
