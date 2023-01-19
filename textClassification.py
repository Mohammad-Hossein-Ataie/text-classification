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


def checkOccurence(words):
    """
    A function to determine the occurrence of the list of words. Actually finds a list of strings and informs us of each
    occurrence in a final list.)
    """
    rlist = []
    for j in range(len(words)):
        flag = 0
        if len(words[j]) != 2:
            rlist.append(firstOccurrence(words[j]))
        else:
            for x in range(len(words[j])):
                if firstOccurrence(words[j][x]) == -1:
                    continue
                else:
                    flag = 1
                    rlist.append(firstOccurrence(words[j][x]))
            if flag == 0:
                rlist.append(-1)
    return rlist


def isAscending(wordsList):
    """
    Function to check the order of occurrence of words
    """
    previous = wordsList[0]
    for number in wordsList:
        if number < previous:
            return False
        previous = number
    return True


def setScoreBaseOnOccurrence(listOfScore):
    """
    A scoring function based on word order and occurrence
    """
    for y in range(len(TotalWords)):
        wordsList = checkOccurence(TotalWords[y])
        sumation = 0
        length = len(wordsList)
        for p in range(length):
            if wordsList[p] != -1:
                sumation += 1
        score = sumation / length
        if isAscending(wordsList) == True and wordsList[-1] != -1:
            score += 0.5
        listOfScore.append(score)
    return listOfScore


def updateScoreBaseOnDocumentSize(listOfScore):
    """
    A function that updated each document score base on its size.
    """
    if docSize > 300:
        listOfScore[1] -= 0.5
    if docSize > 4000:
        listOfScore[2] -= 0.6
        listOfScore[4] -= 0.7
        listOfScore[0] -= 0.5
    if docSize > 8000:
        listOfScore[3] -= 0.4
    return listOfScore


def showResult(resultList):
    maxres = max(resultList)
    index = resultList.index(maxres)
    if index == 0:
        return "صورت جلسه"
    elif index == 1:
        return "نامه"
    elif index == 2:
        return "تمرین"
    elif index == 3:
        return "مقاله"
    elif index == 4:
        return "پروپوزال"
    elif index == 5:
        return "کتاب"
    elif index == 6:
        return "پایان نامه"


if __name__ == "__main__":
    """
    Specifying words related to each type of document.
    """
    words_soratJalase = [["عنوان", "موضوع"], "تاریخ", "زمان", ["بحث", "مباحث"], ["تصمیمات", "مصوبات"], "امضا"]
    words_naame = [["عنوان", "موضوع"], "سلام", "احترام", "تشکر", "امضا"]
    words_tamrin = [["نام", "خانوادگی"], "شماره دانشجویی", "تمرین", ["پرسش", "سوال"], ["پاسخ", "جواب"]]
    words_article = [["مولف ", "نویسنده"], "چکیده", "کلید¬واژه", "مقدمه", "ابزار", "روش", "نتایج", "جمع¬بندی",
                     ["منابع ", "ارجاع"]]
    words_proposal = ["عنوان ", " پروپوزال", ["استاد ", "راهنما"], "داوری", "نام دانشجو", "شماره", ["تعریف ", "مسئله"],
                      ["پیشینه ", "پژوهش"], "مراجع", "روش", "خارجی"]
    words_book = [" فهرست", ["پیشگفتار ", "درباره نویسنده"], " مقدمه", ["نویسنده ", "مترجم"], "فصل"]
    words_payanNaame = [["عنوان ", "موضوع"], "پایان¬نامه", ["استاد ", "راهنما"], "نگارش",
                        ["فهرست مطالب ", "فهرست اشکال"], "چکیده", ["فصل ", "فصول"], ["منابع ", "مراجع"],
                        ["ضمیمه ", "پیوست"]]
    TotalWords = [words_soratJalase, words_naame, words_tamrin, words_article, words_proposal, words_book,
                  words_payanNaame]

    document = Document("test.docx")
    paras = [p.text for p in document.paragraphs if p.text]
    df = pd.DataFrame(paras)
    docSize = setDocSize(df)
    listOfScores = []
    listOfScores = setScoreBaseOnOccurrence(listOfScores)
    listOfScores = updateScoreBaseOnDocumentSize(listOfScores)
    print(showResult(listOfScores))
