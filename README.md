
# Generic Sentiment Analysis from Excel: Hello Maritime Mile

### Introduction

This software will perform sentiment analysis on text present in an excel file. This analysis will be performed using textblob and the NLTK corpus. 


### Prerequisites

This code relies upon the following Python packages:
* xlrd
* xlwt
* matplotlib

As such these need to be installed, it is recommended to use PIP using the following commands.

    pip install xlrd
    pip install xlwt
    pip install matplotlib 

In addition, the NLTK corpus will need to be installed in order for the textblob process to perform analysis. This will need to be installed to your active python locale.
It is reccomended to achieve this using the following command.

    python -m textblob.download_corpora

Note: this code is somewhat generic and can be modified for other applications, read the comments in line to understand how.

### Output
 The software provides outputs in a multitude of ways:
 * a summary in the console output  
 * a graphical representation in the form of a pie chart
 * an excel file with detailled sentiment scores for each text  sample provided