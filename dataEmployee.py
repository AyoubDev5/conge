import pandas as pd
from sqlalchemy import create_engine

df = pd.read_excel("data.xlsx", sheet_name="employee")

engine = create_engine("sqlite:///gestion_rh.db")

ids = pd.read_sql("SELECT id FROM employes", con=engine)
df = df[~df["id"].isin(ids["id"])]

df.to_sql("employes", con=engine, if_exists="append", index=False)

# df = pd.read_excel("data.xlsx", sheet_name="periodes")

# df.to_sql("periodes", con=engine, if_exists="append", index=False)

print("Data successfully imported into the database!")