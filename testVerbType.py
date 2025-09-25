import spacy
import pandas as pd


nlp = spacy.load("de_core_news_sm")    

# read excel file
sentences = pd.ExcelFile("result.xlsx")
sheetname = "Sheet1"
df = pd.read_excel(sentences, sheetname, header = 0)

print("still going!")

def FV(sentence):
    doc = nlp(str(sentence))
    for token in doc:
        if "VerbForm=Fin" in token.morph:
            if "VerbType=Mod" in token.morph or token.tag_.startswith("VM"):
                return "modal"
            elif "VerbType=Aux" in token.morph or token.tag_.startswith("VA"):
                return "aux"
            elif token.lemma_ == "sein":
                return "cop"
            else:
                return "main"
    return ""

print("almost done!")

df["verb"] = df["utterance"].apply(FV)
#df["verb"] = df["verb"].replace("ROOT", "fv") --> rename ROOT verbs to fv
df.to_excel("result.xlsx", index = False)

print ("done!")