"""
First we add a dataset of text including persian mail, book, proposal, article, proceedings and dissertation.
"""
import docx
import pandas as pd
import json


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


def find_word_occurrence(words):
    """
    A function to find the first occurrence of each word in a list of words.
    """
    result = []
    for word in words:
        if len(word) == 2:
            sub_results = [firstOccurrence(w) for w in word if firstOccurrence(w) != -1]
            if sub_results:
                result.extend(sub_results)
            else:
                result.append(-1)
        else:
            result.append(firstOccurrence(word))
    return result


def is_ascending_order(words_list):
    """
    Function to check if the words occur in ascending order
    """
    return all(word1 <= word2 for word1, word2 in zip(words_list, words_list[1:]))


def setScoreBasedOnOccurrence(listOfScore, totalWords):
    """
    A scoring function based on word order and occurrence
    """
    for y in range(len(totalWords)):
        wordsList = find_word_occurrence(totalWords[y])
        total_occurrences = sum(1 for word in wordsList if word != -1)
        score = total_occurrences / len(wordsList)
        if is_ascending_order(wordsList) and wordsList[-1] != -1:
            score += 0.5
        listOfScore.append(score)
    return listOfScore


def updateScoreBaseOnDocumentSize(listOfScore, documentSize):
    """
    A function that updates each document score based on its size.
    """
    if documentSize > 300:
        listOfScore[1] -= 0.5
    if documentSize > 4000:
        listOfScore[2] -= 0.6
        listOfScore[4] -= 0.7
        listOfScore[0] -= 0.5
    if documentSize > 8000:
        listOfScore[3] -= 0.4
    return listOfScore


def showResult(resultList):
    """
    Function to display the result based on the scores given in the previous sections.
    """
    result_map = {
        0: "صورت جلسه",
        1: "نامه",
        2: "تمرین",
        3: "مقاله",
        4: "پروپوزال",
        5: "کتاب",
        6: "پایان نامه"
    }

    maxres = max(resultList)
    index = resultList.index(maxres)
    return result_map.get(index, "Invalid index")


if __name__ == "__main__":
    """
    Specifying words related to each type of document.
    """
    with open('config.json', encoding="utf8") as f:
        config = json.load(f)

    words_soratJalase = config['words_soratJalase']
    words_naame = config['words_naame']
    words_tamrin = config['words_tamrin']
    words_article = config['words_maqale']
    words_proposal = config['words_proposed']
    words_book = config['words_ketab']
    words_payanNaame = config['words_payaneNam']
    TotalWords = [words_soratJalase, words_naame, words_tamrin, words_article, words_proposal, words_book,
                  words_payanNaame]

    document = docx.Document("dataset/article.docx")
    paras = [p.text for p in document.paragraphs if p.text]
    df = pd.DataFrame(paras)
    docSize = setDocSize(df)
    listOfScores = []
    listOfScores = setScoreBasedOnOccurrence(listOfScores, TotalWords)
    listOfScores = updateScoreBaseOnDocumentSize(listOfScores, docSize)
    print(showResult(listOfScores))

