import pandas as pd

try:
    df = pd.read_excel('test sheet Fed ex .xlsx')
    print("Columns in 'test sheet Fed ex .xlsx':")
    for col in df.columns:
        print(col)
    print("\nFirst row sample:")
    print(df.iloc[0].to_dict())
except Exception as e:
    print(f"Error reading excel file: {e}")
