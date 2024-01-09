import pandas as pd
df = pd.read_csv("/path/recentLoginList.csv", skip_blank_lines=False)
df = df[['UserPrincipalName', 'Name', 'AppInteractLastLogin', 'TokenInteractLastLogin', 'IsLicensed', 'IsGuestUser']]
df.to_csv("/path/LastLoginDateReport.csv", index=False)