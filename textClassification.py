"""
First we add a dataset of text including persian mail, book, proposal, article, proceedings and dissertation.
"""
from docx import Document
import pandas as pd
"""
Then we should add a function to find number of words in each document.
"""


def setDocSize(dataFrame):
    documentSize = 0
    for i in range(len(dataFrame[0])):
        documentSize += len(dataFrame[0][i].split(" "))
    return documentSize


if __name__ == "__main__":
    document = Document("test.docx")
    paras = [p.text for p in document.paragraphs if p.text]
    df = pd.DataFrame(paras)
    docSize = setDocSize(df)


