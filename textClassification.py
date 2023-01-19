"""
First we add a dataset of text including persian mail, book, proposal, article, proceedings and dissertation.
"""
from docx import Document
import pandas as pd


def setDocSize(dataFrame):
    """
    Then we should add a function to find number of words in each document.
    """
    documentSize = 0
    for i in range(len(dataFrame[0])):
        documentSize += len(dataFrame[0][i].split(" "))
    return documentSize


def firstOccurrence(string):
    """
    Add a firstOccurrence to check where a string occurs in the document for the first time.
    """
    for i in range(len(df[0].str.count(string))):
        if df[0].str.count(string)[i] == 1:
            return i
    return -1


if __name__ == "__main__":
    document = Document("test.docx")
    paras = [p.text for p in document.paragraphs if p.text]
    df = pd.DataFrame(paras)
    docSize = setDocSize(df)


