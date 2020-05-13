#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Code written by Sam Reeves for general use.
samtreeves@gmail.com


This script will read a file data.xlsx and search it for
instances of terms defined in a single column in criteria.xlsx.
It will output a separate .xlsx file for the results of each
search.


"""

# LOADING

# Import the data as pandas DataFrame objects

import pandas
#import numpy

# Set pandas to display all columns
pandas.set_option('display.max_columns', None)

data = pandas.read_excel("data.xlsx")
criteria = pandas.read_excel("criteria.xlsx")

# Create dummy data for testing
#data = pandas.DataFrame(numpy.random.randint(0,100,size=(10000, 20)), columns=list('ABCDEFGHIJKLMNOPQRST'))
#criteria = pandas.DataFrame(numpy.random.randint(0,10,size=(5, 1)), columns=['Term'])

# PREPROCESSING

# Make everything lowercase
data = data.apply(lambda x: x.astype(str).str.lower())
criteria = criteria.apply(lambda x: x.astype(str).str.lower())

# Assert string type
terms = list(criteria.Term)

# FUNCTIONS

# Input data and criteria (names as defaults)
# Output list of lines with matches
def find_indices(data=data, terms=terms):
    found = {}
    for term in terms:
        found[term] = []
        
    # For row and column of data, check to see if the term exists within the string
        for i in range(len(data)):
            match = False

            for j in range(len(data.columns)):
                if term in data.iloc[i,j]:
                    found[term].append(i)

    return found

# RUN MAIN FUNCTION

found = find_indices()

# SAVE OUTPUT
for term in terms:
    df = data.iloc[found[term]]
    pandas.DataFrame(df).to_excel(str(term) + '.xlsx')