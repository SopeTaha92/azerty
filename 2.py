


import pandas as pd 



data = {
    'nom': ['Alice', 'Bob', 'Charlie', 'Diana'],
    'age': [25, 30, 35, 28],
    'ville': ['Paris', 'Lyon', 'Marseille', 'Paris'],
    'salaire': [3000, 4000, 3500, 3200]
}


df_brute = pd.DataFrame(data)
print(df_brute)

df = df_brute.copy()
print(df.head(2))
print(df.tail(1))
print(df.shape)