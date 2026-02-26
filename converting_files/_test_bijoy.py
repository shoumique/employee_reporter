import pandas as pd

df = pd.read_csv('Data_Format_unicode.csv', dtype=str)

# Check the specific columns the user flagged
watch = [
    'Super-Newmerray / Normal',
    'Rural/Urben',
    'Last Promotion SL',
    'Current Branch Duration',
    'UT Status',
    'Age_Calculator',
    'M/F',
    'খসড়া চার্জশীট',
    'নাম',
    'পিতার নাম',
    'নথি নং',
]

print('=== Column headers (first 15) ===')
for c in list(df.columns)[:15]:
    print(' ', repr(c))

print()
print('=== Spot-check columns ===')
for col in df.columns:
    if any(w in col for w in watch):
        sample = df[col].dropna().iloc[:3].tolist() if col in df else []
        print(f'  {col!r}: {sample}')

print()
print('=== Row 2 sample fields ===')
row = df.iloc[1]
for key in ['নাম', 'পিতার নাম', 'নিজ জেলা', 'যোগদানকৃত পদবী', 'Rural/Urben', 'Current Branch Duration']:
    if key in df.columns:
        print(f'  {key!r}: {row[key]!r}')
