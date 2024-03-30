from transformers import pipeline


classifier = pipeline("sentiment-analysis",   
                      "blanchefort/rubert-base-cased-sentiment")
from openpyxl import load_workbook
Book = load_workbook("Train.xlsx")

rev_sheet = Book["Rev"]

revs = []

count = 0

for item in rev_sheet:
    count += 1


for i in range(2, count + 1):
    text = ''.join(rev_sheet["B" + str(i)].value.replace(u'\xa0', u' ').replace(u'\n', u' '))
    if len(text) > 512:
        text = text[:(512-len(text))]
    revs.append(text)


tonal = []

for review in revs:
    verdict = classifier(review)[0]
    if verdict['label'] == 'NEUTRAL':
        verdict['label'] = 'NEGATIVE'
    to = verdict['label']
    tonal.append(to)

for i in range(len(tonal)):
    rev_sheet["C" + str(i + 2)] = tonal[i]

Book.save("Train.xlsx")

