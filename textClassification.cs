// 
// First we add a dataset of text including persian mail, book, proposal, article, proceedings and dissertation.
// 
namespace Namespace {
    
    using docx;
    
    using pd = pandas;
    
    using json;
    
    using System;
    
    using System.Linq;
    
    using System.Collections.Generic;
    
    public static class Module {
        
        // 
        //     Then we should add a function to find number of words in each document.
        //     
        public static object setDocSize(object dataFrame) {
            var documentSize = 0;
            foreach (var i in Enumerable.Range(0, dataFrame[0].Count)) {
                documentSize += dataFrame[0][i].split(" ").Count;
            }
            return documentSize;
        }
        
        // 
        //     Add a firstOccurrence to check where a string occurs in the document for the first time.
        //     
        public static object firstOccurrence(object @string) {
            foreach (var i in Enumerable.Range(0, df[0].str.count(@string).Count)) {
                if (df[0].str.count(@string)[i] == 1) {
                    return i;
                }
            }
            return -1;
        }
        
        // 
        //     A function to find the first occurrence of each word in a list of words.
        //     
        public static object find_word_occurrence(object words) {
            var result = new List<object>();
            foreach (var word in words) {
                if (word.Count == 2) {
                    var sub_results = (from w in word
                        where firstOccurrence(w) != -1
                        select firstOccurrence(w)).ToList();
                    if (sub_results) {
                        result.extend(sub_results);
                    } else {
                        result.append(-1);
                    }
                } else {
                    result.append(firstOccurrence(word));
                }
            }
            return result;
        }
        
        // 
        //     Function to check if the words occur in ascending order
        //     
        public static object is_ascending_order(object words_list) {
            return all(from _tup_1 in zip(words_list, words_list[1]).Chop((word1,word2) => (word1, word2))
                let word1 = _tup_1.Item1
                let word2 = _tup_1.Item2
                select word1 <= word2);
        }
        
        // 
        //     A scoring function based on word order and occurrence
        //     
        public static object setScoreBasedOnOccurrence(object listOfScore, object totalWords) {
            foreach (var y in Enumerable.Range(0, totalWords.Count)) {
                var wordsList = find_word_occurrence(totalWords[y]);
                var total_occurrences = (from word in wordsList
                    where word != -1
                    select 1).Sum();
                var score = total_occurrences / wordsList.Count;
                if (is_ascending_order(wordsList) && wordsList[-1] != -1) {
                    score += 0.5;
                }
                listOfScore.append(score);
            }
            return listOfScore;
        }
        
        // 
        //     A function that updates each document score based on its size.
        //     
        public static object updateScoreBaseOnDocumentSize(object listOfScore, object documentSize) {
            if (documentSize > 300) {
                listOfScore[1] -= 0.5;
            }
            if (documentSize > 4000) {
                listOfScore[2] -= 0.6;
                listOfScore[4] -= 0.7;
                listOfScore[0] -= 0.5;
            }
            if (documentSize > 8000) {
                listOfScore[3] -= 0.4;
            }
            return listOfScore;
        }
        
        // 
        //     Function to display the result based on the scores given in the previous sections.
        //     
        public static object showResult(object resultList) {
            var result_map = new Dictionary<object, object> {
                {
                    0,
                    "صورت جلسه"},
                {
                    1,
                    "نامه"},
                {
                    2,
                    "تمرین"},
                {
                    3,
                    "مقاله"},
                {
                    4,
                    "پروپوزال"},
                {
                    5,
                    "کتاب"},
                {
                    6,
                    "پایان نامه"}};
            var maxres = max(resultList);
            var index = resultList.index(maxres);
            return result_map.get(index, "Invalid index");
        }
        
        static Module() {
            @"
    Specifying words related to each type of document.
    ";
        }
        
        public static object config = json.load(f);
        
        public static object words_soratJalase = config["words_soratJalase"];
        
        public static object words_naame = config["words_naame"];
        
        public static object words_tamrin = config["words_tamrin"];
        
        public static object words_article = config["words_maqale"];
        
        public static object words_proposal = config["words_proposed"];
        
        public static object words_book = config["words_ketab"];
        
        public static object words_payanNaame = config["words_payaneNam"];
        
        public static object TotalWords = new List<object> {
            words_soratJalase,
            words_naame,
            words_tamrin,
            words_article,
            words_proposal,
            words_book,
            words_payanNaame
        };
        
        public static object document = docx.Document("dataset/article.docx");
        
        public static object paras = (from p in document.paragraphs
            where p.text
            select p.text).ToList();
        
        public static object df = pd.DataFrame(paras);
        
        public static object docSize = setDocSize(df);
        
        public static object listOfScores = new List<object>();
        
        public static object listOfScores = setScoreBasedOnOccurrence(listOfScores, TotalWords);
        
        public static object listOfScores = updateScoreBaseOnDocumentSize(listOfScores, docSize);
    }
}
