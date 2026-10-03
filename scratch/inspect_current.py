import re

with open('index.html', encoding='utf-8') as f:
    text = f.read()

ids = sorted(set(re.findall(r'id=["\']([a-zA-Z0-9_\-]+)["\']', text)))
print(f'Total IDs in index.html: {len(ids)}')
for id_name in ids:
    print(id_name)
