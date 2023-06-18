import numpy as np
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import nltk
from nltk.corpus import stopwords

nltk.download('stopwords')
InbuiltStopwords = stopwords.words('english')

# Importing Stop Words:
with open('C:/Users/Computer/Desktop/Blackcoffer Internship/StopWords/StopWords_Auditor.txt','r', encoding='latin-1') as f1, open('C:/Users/Computer/Desktop/Blackcoffer Internship/StopWords/StopWords_Currencies.txt','r', encoding='latin-1') as f2, open('C:/Users/Computer/Desktop/Blackcoffer Internship/StopWords/StopWords_DatesandNumbers.txt','r', encoding='latin-1') as f3, open('C:/Users/Computer/Desktop/Blackcoffer Internship/StopWords/StopWords_Generic.txt','r', encoding='latin-1') as f4, open('C:/Users/Computer/Desktop/Blackcoffer Internship/StopWords/StopWords_GenericLong.txt','r', encoding='latin-1') as f5, open('C:/Users/Computer/Desktop/Blackcoffer Internship/StopWords/StopWords_Geographic.txt','r', encoding='latin-1') as f6, open('C:/Users/Computer/Desktop/Blackcoffer Internship/StopWords/StopWords_Names.txt','r', encoding='latin-1') as f7:
    MyStopWords = ([f1.readlines(),f2.readlines(),f3.readlines(),f4.readlines(),f5.readlines(),f6.readlines(),f7.readlines()])

MyStopWords = sum(MyStopWords, [])
MyStopWords = [item.replace('\n', '').replace(' ', '').replace('\t', '').replace('\r', '').split('|')for item in MyStopWords]
MyStopWords.append(InbuiltStopwords)
MyStopWords = list(np.concatenate(MyStopWords))

# Importing Positive Words:
with open('C:/Users/Computer/Desktop/Blackcoffer Internship/MasterDictionary/positive-words.txt','r', encoding='latin-1') as pos:
    PositiveWords = pos.read().split("\n")

# Importing Negative Words:
with open('C:/Users/Computer/Desktop/Blackcoffer Internship/MasterDictionary/negative-words.txt','r', encoding='latin-1') as neg:
    NegativeWords = neg.read().split("\n")

UrlFile = pd.read_excel('C:/Users/Computer/Desktop/Blackcoffer Internship/Input.xlsx')
#ColUrlId = []
#ColUrl = []
ColTitle = []
ColContent = []
ColSentence = []
ColCharCount = []
ColSentenceCount = []
ColToken = []
ColTokenCount = []
ColFilToken = []
ColFilTokenCount = []
ColPosScore = []
ColNegScore = []
ColPolarityScore = []
ColSubjectivityScore = []
ColAverageSenLen = []
ColComplexWordCount = []
ColPercComplexWords = []
ColFogIndex = []
ColAverageWordsPerSentence = []
ColSyllablePerWord = []
ColAvgWordLength = []
ColPersonalPronoun = []

for urls, UrlId in zip(UrlFile['URL'], UrlFile['URL_ID']):
    try:
        headers = {"Accept-Language": "es-ES,es;q=0.9",
                   "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (HTML, like Gecko) Chrome/94.0.4606.81 Safari/537.36"}
        PageData = requests.get(urls, headers=headers)
        FilPageData = BeautifulSoup(PageData.content, 'html.parser')
        if FilPageData.find('h1'):
            #ColUrl.append(urls)
            #ColUrlId.append(UrlId)
            print(UrlId)
            Title = FilPageData.find('h1')  # Getting the Page Title
            Title = re.sub(r'<[^>]+>', '', str(Title))  # Removing Unwanted Tags
            Title = Title.replace("/n", " ")  # Removing New-Line Variables
            ColTitle.append(Title)
            Content = FilPageData.findAll(attrs={'class': 'td-post-content'})  # Getting Page Contents
            Content = Content[0].text.replace('\n', "").replace('\t', '')  # Removing New-Line Variables
            ColContent.append(Content)
            Sentence = re.split(r'[?!.]', Content)
            ColSentence.append(Sentence)
        else:
            #ColUrl.append(None)
            #ColUrlId.append(None)
            ColTitle.append(None)
            ColContent.append(None)
            ColSentence.append(None)
    except Exception as e:
        print(e)
        continue

def FunCount():
    for ele in ColContent:
        if ele is not None:
            CharCount = 0
            CharCount += len(ele)
            ColCharCount.append(CharCount)
        else:
            ColCharCount.append(None)
    for ele in ColSentence:
        if ele is not None:
            SentenceCount = 0
            SentenceCount += len(ele)
            ColSentenceCount.append(SentenceCount)
            #print(ColCharCount)
            #print(ColSentenceCount)
        else:
                ColSentenceCount.append(None)


def FunToken():
    Punc = '''!()-[]{};:–'"“”`'\,<>./?@#$%^&*_~'''
    for ele in ColContent:
        if ele is not None:
            PuncContent = ele.translate(str.maketrans('','',Punc))
            PuncContent = PuncContent.split(' ')
            ColToken.append(PuncContent)
        else:
            ColToken.append(None)

def FunFilToken():
    temp = []
    for ele in ColToken:
        if ele is not None:
            for e in ele:
                if e not in MyStopWords:
                    temp.append(e)
            ColFilToken.append(temp)
            temp = []
        else:
            ColFilToken.append(None)

    for ele1, ele2 in zip(ColToken, ColFilToken):
        if ele1 and ele2 is not None:
            ColTokenCount.append(len(ele1))
            ColFilTokenCount.append(len(ele2))
        else:
            ColTokenCount.append(None)
            ColFilTokenCount.append(None)


def FunPosNegWords():
    for ele in ColFilToken:
        if ele is not None:
            PosScore = 0
            NegScore = 0
            for e in ele:
                if e in PositiveWords:
                    PosScore += 1
                elif e in NegativeWords:
                    NegScore += 1
            ColPosScore.append(PosScore)
            ColNegScore.append(NegScore)
        else:
            ColPosScore.append(None)
            ColNegScore.append(None)


def FunPolarityScore():
    for ele1, ele2 in zip(ColPosScore, ColNegScore):
        if ele1 and ele2 is not None:
            PolarityScore = (ele1-ele2)/((ele1+ele2)+0.000001)
            ColPolarityScore.append(PolarityScore)
        else:
            ColPolarityScore.append(None)


def FunSubjectivityScore():
    for ele1, ele2, ele3 in zip(ColPosScore, ColNegScore, ColFilTokenCount):
        if ele1 and ele2 and ele3 is not None:
            SubjectivityScore = (ele1+ele2)/(ele3+0.000001)
            ColSubjectivityScore.append(SubjectivityScore)
        else:
            ColSubjectivityScore.append(None)


def FunAnalysisReadability():
    for ele1, ele2 in zip(ColCharCount, ColSentenceCount):
        if ele1 and ele2 is not None:
            AverageSenLen = ele1 / ele2
            ColAverageSenLen.append(AverageSenLen)
        else:
            ColAverageSenLen.append(None)

    for ele in ColFilToken:
        if ele is not None:
            ComplexWordCount = 0
            for ele2 in ele:
                VowelCount = 0
                V = 'AEIOUaeiou'
                for e in ele2:
                    if e in V:
                        VowelCount += 1
                if VowelCount >= 2:
                    ComplexWordCount += 1
                if ele2.endswith("es" or "ed"):
                    ComplexWordCount -= 1

            ColComplexWordCount.append(ComplexWordCount)
        else:
            ColComplexWordCount.append(None)

    for ele1, ele2 in zip(ColComplexWordCount, ColTokenCount):
        if ele1 and ele2 is not None:
            PercComplexWords = (ele1 / ele2) * 100
            ColPercComplexWords.append(PercComplexWords)
        else:
            ColPercComplexWords.append(None)

    for ele1, ele2 in zip(ColAverageSenLen, ColPercComplexWords):
        if ele1 and ele2 is not None:
            FogIndex = 0.4 * (ele1 + ele2)
            ColFogIndex.append(FogIndex)
        else:
            ColFogIndex.append(None)

    for ele1, ele2 in zip(ColTokenCount, ColSentenceCount):
        if ele1 and ele2 is not None:
            AverageWordPerSentence = ele1 / ele2
            ColAverageWordsPerSentence.append(AverageWordPerSentence)
        else:
            ColAverageWordsPerSentence.append(None)

    for ele1, ele2 in zip(ColComplexWordCount, ColFilTokenCount):
        if ele1 and ele2 is not None:
            SyllablePerWord = ele1 / ele2
            ColSyllablePerWord.append(SyllablePerWord)
        else:
            ColSyllablePerWord.append(None)

    for ele1, ele2 in zip(ColCharCount, ColTokenCount):
        if ele1 and ele2 is not None:
            AvgWordLength = ele1 / ele2
            ColAvgWordLength.append(AvgWordLength)
        else:
            ColAvgWordLength.append(None)


def FunPersonalPronoun():
    PP = ['I', 'My', 'We', 'we', 'my', 'Ours', 'ours', 'Us', 'us']
    Count = 0
    for ele in ColToken:
        if ele is not None:
            if ele in PP:
                Count += 1
                ColPersonalPronoun.append(ele)
        else:
            ColPersonalPronoun.append(None)


FunCount()  # Function to calculate Char count on Website.
FunToken()  # Function to generate Tokens.
FunFilToken()
FunPosNegWords()
FunPolarityScore()
FunSubjectivityScore()
FunAnalysisReadability()
#FunPersonalPronoun()

FinalFile = pd.DataFrame(list(zip(ColPosScore, ColNegScore, ColPolarityScore, ColSubjectivityScore, ColAverageSenLen, ColPercComplexWords, ColFogIndex, ColAverageWordsPerSentence, ColComplexWordCount, ColFilTokenCount, ColSyllablePerWord, ColAvgWordLength)),  columns=None)


with pd.ExcelWriter("C:/Users/Computer/Desktop/Blackcoffer Internship/Output Data Structure.xlsx", engine="openpyxl", mode="a", if_sheet_exists='overlay',) as writer:
    FinalFile.to_excel(writer, startrow=1, startcol=2, index_label=False, index=False, header=False )

