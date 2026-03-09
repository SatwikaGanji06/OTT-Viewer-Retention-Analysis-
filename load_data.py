import pandas as pd
from sqlalchemy import create_engine

engine = create_engine(
    "postgresql://postgres:postgres123@localhost:5432/mydb"
)

df = pd.read_csv("data.csv")

df.to_sql(
    "sales_data",     # TABLE NAME
    engine,
    if_exists="replace",
    index=False
)

print("Data loaded successfully")
