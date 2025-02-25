import pandas as pd
from sqlalchemy import create_engine

df = pd.read_excel("data.xlsx", sheet_name="holidays")

engine = create_engine("sqlite:///gestion_rh.db")

ids = pd.read_sql("SELECT date_holiday FROM holidays", con=engine)
df = df[~df["date_holiday"].isin(ids["date_holiday"])]

df.to_sql("holidays", con=engine, if_exists="append", index=False)

print("Data successfully imported into the database!")