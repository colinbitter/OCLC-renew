import pandas as pd
from bookops_worldcat import WorldcatAccessToken
import glob
from bookops_worldcat import MetadataSession
from pathlib import Path

token = WorldcatAccessToken(
    key="InsertKey",
    secret="InsertSecret",
    scopes=["WorldCatMetadataAPI"],
    principal_id="InsertPrincipal_ID",
    principal_idns="InsertPrincipal_idns")

session = MetadataSession(authorization=token, timeout=20)

downloads_path = str(Path.home() / "Downloads")
path1 = downloads_path
allFiles = glob.glob(path1 + "/*.xlsx")  # xlsx file needs header
data_df = pd.DataFrame()
list_ = []
for file_ in allFiles:
    data_df = pd.read_excel(file_, names=['localID', 'OCLC'], engine='openpyxl')

df1 = data_df['OCLC'].tolist()

dff3 = []
dff4 = []
with MetadataSession(authorization=token) as session:
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
data_df['BATCHMATCH'] = (data_df['OCLC'] == data_df['BATCH'])
data_df['HELD'] = df5

data_df.to_csv(path1 + "/OCLCresult.csv", index=False)
