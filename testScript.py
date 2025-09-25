import spacy
import pandas as pd

#nlp = spacy.load("de_core_news_sm")    
#sent = "Ich flog nach Rom."
#doc = nlp(sent)
#for token in doc:
    #print(token.text,list(token.morph))


# REAL CODE

nlp = spacy.load("de_core_news_sm")    

# read excel file
sentences = pd.ExcelFile("pythonTest.xlsx")
sheetname = "MC"
df = pd.read_excel(sentences, sheetname, header = 0)
#dfe = df["E"]
#dfh = df["H"]

print("still going!")

#debug version
#def FV(sentence):
    #doc = nlp(str(sentence))
    #for i, token in enumerate(doc):
        #if "VerbForm=Fin" in token.morph:
            #if i > 0:
                #print(f"Sentence: {sentence} | Finite verb: {token.text} | Word before: {doc[i-1].text}")
                #return doc[i-1].text
            #else:
                #print(f"Sentence: {sentence} | Finite verb: {token.text} | No word before")
                #return ""
    #print(f"Sentence: {sentence} | No finite verb found")
    #return ""

def FV(sentence):
    doc = nlp(str(sentence))
    symbols = {'@', '<', '(', ')', '>', '&', ']', '[', '/', '2', '0', '1'}
    for i, token in enumerate(doc):
        if "VerbForm=Fin" in token.morph:
            j = i - 1
            while j >= 0:
                if not any(sym in doc[j].text for sym in symbols):
                    return doc[j].text
                j -= 1
            return ""
    return ""

print("almost done!")

df["top"] = df["utterance"].apply(FV)
df.to_excel("result.xlsx", index = False)

print ("done!")