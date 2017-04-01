# Steven Chan - 
# srchan 2016.01.25
# bleuscores - calculates BLEU scores for 3-column CSV files, given:
# - column 1: original translation
# - column 2: reference translation (human)
# - column 3: computer translation being evaluated
# Each CSV file represents an individual engine
#
# Revised 2016.03.01 to accommodate:
# - ANOVA statistics for p-values and between-group testing
# - finding CSV files by recursing through subdirectories
# - calculating mean, standard error of means per "engine"
 
import csv
import nltk.translate
import nltk.tokenize
import os
 
import numpy
import scipy.stats  # for f_oneway()
 
 
## VARIABLES
rootdir = '.'
bleuweights = [0.25, 0.25, 0.25, 0.25] # for n-gram, bigram, etc.
bleuScores = []
 
engines = {} # insert values into this engine.
 
print '----------'
print '----------'
print 'Per-file statistics'
 
 
for subdir, dirs, files in os.walk(rootdir):
    print 'Subdirectory: ' + str(subdir)
    # print files
    for filename in files: # os.listdir(subdir):
        if not filename.endswith('.csv'):
            continue # skip non-csv file
        csvFilename = os.path.join(subdir, filename)
        engine = filename.lower()
 
        csvRows = []
        bleuScores = []
 
        currentFile = open(csvFilename, 'rU')
        currentReader = csv.reader(currentFile, dialect=csv.excel)
        print ('-- Filename :' + filename)
        for row in currentReader:
            if currentReader.line_num == 1:
                continue  #skip first row
 
            #reference = nltk.tokenize.word_tokenize(row[1].encode("ascii","ignore"))
            #hypothesis = nltk.tokenize.word_tokenize(row[2].encode("ascii","ignore"))
            reference = row[1].split(" ")
            hypothesis = row[2].split(" ")
 
            references = [reference]
            bleuScore = nltk.translate.bleu(references, hypothesis, bleuweights)
            bleuScores.append(bleuScore)
 
            # print ('Row #' + str(currentReader.line_num) + ' scored ' + str(bleuScore) + ' ' + str(row))
 
 
 
        currentFile.close()
 
        # print ('-- BLEU Scores: ' + str(bleuScores))
 
        averageResult = float(sum(bleuScores)) / len(bleuScores)
        print ('---- BLEU Average: ' + str(averageResult))
 
           # Write out the CSV file.
           # csvFileObj = open(os.path.join('headerRemoved', csvFilename), 'w',
           #              newline='')
           # csvWriter = csv.writer(csvFileObj)
           # for row in csvRows:
           #     csvWriter.writerow(row)
           # csvFileObj.close()
        try:   # Append if already exists
            engines[engine].extend(bleuScores)
        except KeyError: 
            engines[engine] = bleuScores
 
 
allValues = []
 
print '----------'
print '----------'
print 'Per-engine statistics'
 
# write out statistics:
for enginename, values in engines.iteritems():
    print 'ENGINE: ' + str(enginename)  # prints engine name
    
    valuesNdarray = numpy.array(values)
 
    # prep for ANOVA statistics
    allValues.append(values)
 
    # mean statistics
    #averageResult = float(sum(values)) / len(values)
    meanValue = valuesNdarray.mean()
    print '    averageResult: ' + str(meanValue)
 
    # error of mean
    semValue = scipy.stats.sem(valuesNdarray)
    lowerValue = meanValue - (semValue * 1.96)
    upperValue = meanValue + (semValue * 1.96)
    print '    standard error of mean: ' + str(semValue)
    print '    confidence interval: (' + str(lowerValue) + ', ' + str(upperValue) + '}'
 
    # count
    print '    count: ' + str(len(values))
 
print '----------'
print '----------'
allValuesNdarray = numpy.array(allValues)
#print 'Raw results (Numpy array)'
#print str(allValuesNdarray)
print '----------'
print '----------'
print '----------'
print 'ANOVA (F-value, p-value): '
# TODO: this ASSUMES 5 engines being tested, for our experiment
print scipy.stats.f_oneway(allValuesNdarray[0],allValuesNdarray[1],allValuesNdarray[2],allValuesNdarray[3],allValuesNdarray[4])
