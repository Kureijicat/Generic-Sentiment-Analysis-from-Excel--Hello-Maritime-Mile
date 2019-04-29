import xlrd 
import xlwt
import re 
import matplotlib.pyplot as plt

from textblob import TextBlob 

'''
Author: Joseph Rafferty, the licence text (MIT) is at the bottom of this file.


This code relises on libraries for correct operation
please ensure that matplotlib, xlrd and xlwt are avaiable.

    pip install xlrd
    pip install xlwt
    pip install matplotlib


    
Below are some global variables to identify:
    -the input file, 
    -column of the file to analyse (zero indexed)
    -the name of the excel sheet to store an output
    (in that order)

Ensure all files to be read exist and all locations to be 
written to are accessible by the process.


'''

gInputFile = "c://root//text.xlsx"
gFocusColumn = 3
gOutputFile = "C://root//sentimentOutput.xls"


'''
Below is the sentiment analysis class, this does the bulk of the work using NLTK and textbloob
'''
class SentiClass(object): 

    def __init__(self): 
        '''
        #init placeholder
        '''

    def cleanText(self, iText): 
        '''
        This cleans the text for parsing
        '''
        try: 
            return ' '.join(re.sub("(@[A-Za-z0-9]+)|([^0-9A-Za-z \t]) |(\w+:\/\/\S+)", " ", iText).split()) 
        
        except:
            return ''
  
    def getTextSentimentPolarity(self, iText): 

        '''
        This uses the textblob apis to perform sentiment analysis.
            This is dependent on the availablity of a NLTK corpus.
            If this doesnt exist on the running environment, issue:
            
            python -m textblob.download_corpora

        '''
        try: 
            
            #Get some analyis of the text
            analysis = TextBlob(iText) 
            return analysis.sentiment.polarity
            
        except Exception as e:
            print(e.message)
            

    def processTextFromXls(self, iPath):
        
        # empty list to store parsed text 
        textFragments = [] 
  
        sentimentSummary = {}
        sentimentSummary['positive']   = 0 
        sentimentSummary['negative']   = 0
        sentimentSummary['neutral']   = 0
        
        provideOutput("Extracting records")
    

        wb = xlrd.open_workbook(iPath) 
        sheet = wb.sheet_by_index(0) 
        sheet.cell_value(0, 0) 

        provideOutput("Performing Analysis")

        for i in range(sheet.nrows): 
            
            parsedRecord = {}
            parsedRecord['text'] = self.cleanText(sheet.cell_value(i, gFocusColumn))
            parsedRecord['sentimentPolarity'] = self.getTextSentimentPolarity(parsedRecord['text'])

            try: 
                # set sentiment 
                if parsedRecord['sentimentPolarity'] > 0: 
                    parsedRecord['sentimentClass']  =  'positive'
                    sentimentSummary['positive']+=1

                elif parsedRecord['sentimentPolarity'] == 0: 
                    parsedRecord['sentimentClass']  =  'neutral'
                    sentimentSummary['neutral']+=1
                else: 
                    parsedRecord['sentimentClass']  = 'negative'
                    sentimentSummary['negative']+=1 
            
            except :
                print("s")
        
            textFragments.append(parsedRecord)

        return textFragments, sentimentSummary

'''
End of the sentiment analysis class
'''

#A reusable function to provide output in a standard style
def provideOutput(iText):
    print("")
    print("--------------------------------")
    print("   ", iText)
    print("--------------------------------")
    print("")


'''
A generic function to write data for an excel file.

    iPath is the full path of the file to be written to.
    iData is the list of dicts to be written, this function 
        operates best when dicts are uniform in the list.
'''

def writeExcel(iPath, iData):

    #Indices to track target cell location
    colTrack = 0
    rowTrack = 0

    #Create the workbook and add a sheet
    book = xlwt.Workbook()
    sh = book.add_sheet("Output")

    #Write the headers
    headings = list(iData[0].keys())

    for text in headings:
        sh.write(rowTrack,colTrack ,text)
        colTrack+= 1


    #Write the data
    for record in iData:
        #Increment the row track
        rowTrack+= 1

        #Reset the Col track
        colTrack = 0
        for v in record.values():
            sh.write(rowTrack,colTrack,v)
            colTrack+=1

    #Save the book
    book.save(iPath)


#Generate a chart from a provided dict, this is fairly generic
def generateChart(iData):

    labels = list(iData.keys())
    sizes = list(iData.values())
    fig1, ax1 = plt.subplots()
    ax1.pie(sizes, labels=labels, autopct='%1.1f%%', shadow=True, startangle=90)
    ax1.axis('equal') 
    plt.show()
  
#The main function
def main(): 
    provideOutput("Author: Joe Rafferty")

    #Create an instance of the Sentiment Analyser
    api = SentiClass() 

    #process excel file content
    corpus, sentimentSummary = api.processTextFromXls(gInputFile)     

    #Write an Excel file with the sentiment analysis
    writeExcel(gOutputFile, corpus)

    #Generate a chart from the sentiment summary scores
    generateChart(sentimentSummary)

    print("")
    print("--------------------------------")
    print("   Sentiment output:")
    print("   ")
    print("   Postive: ", sentimentSummary['positive'])
    print("   Neutral: ", sentimentSummary['neutral'])
    print("   Negative: ", sentimentSummary['negative'])
    print("--------------------------------")
    print("")

#The stub
if __name__ == "__main__": 
    # calling main function 
    main() 


'''

License text - do not remove

MIT License

Copyright (c) 2019 Joe Rafferty

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''