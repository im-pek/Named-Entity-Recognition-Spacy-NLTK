import os
import os.path
from os import path
import json
from datetime import datetime
import xlsxwriter
from scipy.spatial import distance
import itertools
from itertools import permutations
from numpy import unique
from sklearn.cluster import AffinityPropagation
import spacy
from spacy import displacy 
import nltk
from nltk import word_tokenize, pos_tag, ne_chunk
import fileinput
import sys
import csv
import pandas as pd
import enchant
import string
import unidecode
import xlwings as xw

dateTimeObj = datetime.now()
dateObj = dateTimeObj.date()
dateStr = dateObj.strftime("%d %b %Y ") #current date to label output file

time = str(dateTimeObj.hour) + '`' + str(dateTimeObj.minute) + '`' + str(dateTimeObj.second) #current time to label output file

#Singaporean Chinese Surnames
    
df = pd.read_excel(r"L:\\My Documents\\Desktop\\Singaporean Chinese Surnames.xlsx") #change file name and path where necessary

surnames1=df['Surnames 1'].values.tolist()
surnames2=df['Surnames 2'].values.tolist()
surnames3=df['Surnames 3'].values.tolist()
surnames4=df['Surnames 4'].values.tolist()
surnames5=df['Surnames 5'].values.tolist()
surnames6=df['Surnames 6'].values.tolist()
surnames7=df['Surnames 7'].values.tolist()
surnames8=df['Surnames 8'].values.tolist()

surnames = surnames1 + surnames2 + surnames3 + surnames4 + surnames5 + surnames6 + surnames7 + surnames8

cleaned = [x for x in surnames if str(x) != 'nan']

unaccented_surnames = []

for item in cleaned:
    unaccented_string = unidecode.unidecode(item)
    if unaccented_string not in unaccented_surnames:
        unaccented_surnames.append(unaccented_string)

#print (unaccented_surnames)
#print ('\n')

list_of_surnames = []

punctuations = set(string.punctuation)

for item in unaccented_surnames:
    if any(char in punctuations for char in item):
        pass
    else:
        list_of_surnames.append(item)

for ix,surname in enumerate(list_of_surnames):
    if surname == 'Other' or surname == 'Mandarin' or surname == 'Cantonese':
        del list_of_surnames[ix]

#print (list_of_surnames)
#print ('\n')

final_list_of_surnames = []

for idx,word in enumerate(list_of_surnames):
    if word.isalnum():
        if not word.isalpha():
            if not word.isnumeric():
                alpha = word[:-1]
                final_list_of_surnames.append(alpha)
    if word.isalpha():
        final_list_of_surnames.append(word)

added = ['Shek', 'Pek', 'Bo', 'Po', 'Ke', 'Khu', 'Cu', 'Chun', 'Soo', 'Chai', 'Tsay', 'Ko', 'He', 'Chien', 'Tsai', 'Tsao', 'Yu', 'Tiah', 'Tien', 'Hsueh', "Ch'en", "T'ang", "Ts'ao", "Ch'eng", "Ts'ai", "P'eng", "P'an", "T'ien", "Ts'ui", "Ch'iu", "T'an", "Ch'in", "K'ung", "Ch'ang", "Ch'ien", "T'ao", 'Hang']

all_surnames = final_list_of_surnames + added

print (all_surnames)

#

all_months_number = ['1','2','3','4','5','6','7','8','9','10','11','12','01','02','03','04','05','06','07','08','09','10','11','12']
all_months_spelling = ['January','February','March','April','May','June','July','August','September','October','November','December','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
all_months = all_months_number + all_months_spelling

days = list(range(1,32))

day_permutations = days + ['01','02','03','04','05','06','07','08','09','1st','2nd','3rd','4th','5th','6th','7th','8th','9th','First','Second','Third','Fourth','Fifth','Sixth','Seventh','Eighth','Ninth']

all_days=[]

for day in day_permutations:
    new_day = str(day)
    all_days.append(new_day)

year_permutations = list(range(1850,2022))

all_years = []

for year in year_permutations:
    new_year = str(year)
    all_years.append(new_year)

combinations = list(itertools.product(all_days, all_months, all_years))

#print (combinations)

all_perms_and_combs = []

for sets in combinations:
    listed = list(sets)
    permuted = permutations(listed)
    all_perms_and_combs.append(permuted)

#print (all_perms_and_combs) #comment this out, else bandwidth exceeds

all_dates = []

for date in all_perms_and_combs:
    for spec_date in date:
        all_dates.append(list(spec_date))

full_dates = []

for joint in all_dates:
    dates = ' '.join(joint)
    full_dates.append(dates)
    
#    
    
#working_dir_path = pathlib.Path().absolute() #directory path of where this python file is saved

working_dir_paths = "L:\\My Documents\\Desktop" #chosen directory of file inputs
arr = os.listdir(working_dir_paths) #list of file names in this directory


alls_full_directories = []

for array in arr:
    full_files = str(working_dir_paths) + '\\' + array
    alls_full_directories.append(full_files)

for index,file_name in enumerate(alls_full_directories):
    
    this_file_name = arr[index]
    
    print ('\n')    
    print (this_file_name)
    print ('\n')
        
    final_output_file_name = 'L:\\My Documents\\Desktop\\4 May\\Structured Details - ' + this_file_name + '.xlsx' #CHANGE DATE AND FILE DIRECTORY ACCORDINGLY
    
    workbook = xlsxwriter.Workbook(final_output_file_name)
    worksheet = workbook.add_worksheet()
        
    with open(file_name, 'rb') as f:
        if f:
            dic_of_dics = json.load(f, strict = False)
        else:
            pass
        
    dictlist = []
    
    for key, value in dic_of_dics.items():
        temp = [key,value]
        dictlist.append(temp)

    lists=[]

    for item in dictlist:
        for thing in item:
            lists.append(list(thing))

    ###print (lists)
    #for item in lists:
    #    print (item)
    #    print ('\n')

    ###print (len(lists[3])) #number of 'Text' entries in json

    #print lists[3][1]

    __DOBs = []
    all_texts= []
    draft_all_texts = ''
    all_possible_IDs = []
    check_if_in_dates = []
    condition_isnumber = True
    
    us = enchant.Dict("en_US")
    uk = enchant.Dict("en_UK")
    
    for ix,item in enumerate(lists[3]):
        dictionaryy=lists[3][ix]
        get_text=dictionaryy.get('Text')
        
        if get_text != None:
            get_text = get_text.replace('$','S')
            get_text = get_text.replace('Mr','')
            get_text = get_text.replace('Ms','')
            get_text = get_text.replace('MR','')
            get_text = get_text.replace('MS','')
            get_text = get_text.replace('mr','')
            get_text = get_text.replace('ms','')
            get_text = get_text.replace('MRS','')
            get_text = get_text.replace('Mrs','')
            get_text = get_text.replace('mrs','')
            get_text = get_text.replace('MISS','')
            get_text = get_text.replace('Miss','')
            get_text = get_text.replace('miss','')
            get_text = get_text.replace('Mr.','')
            get_text = get_text.replace('Ms.','')
            get_text = get_text.replace('MR.','')
            get_text = get_text.replace('MS.','')
            get_text = get_text.replace('mr.','')
            get_text = get_text.replace('ms.','')
            get_text = get_text.replace('MRS.','')
            get_text = get_text.replace('Mrs.','')
            get_text = get_text.replace('mrs.','')
            get_text = get_text.replace('MISS.','')
            get_text = get_text.replace('Miss.','')
            get_text = get_text.replace('miss.','')
            
            for_checking_get_text = ''.join(e for e in get_text if e.isalnum())
            
            possible_IDs_mix = []
            possible_IDs_num = []
                    
                    
            if not for_checking_get_text.isalpha():
                if not for_checking_get_text.isnumeric():
                    if get_text not in possible_IDs_mix:
                        possible_IDs_mix.append(get_text)
            if for_checking_get_text.isnumeric():
                if get_text not in possible_IDs_num:
                    possible_IDs_num.append(get_text)
            
            possible_IDs = possible_IDs_mix + possible_IDs_num
            
            for item in possible_IDs:
                if len(item) > 5: #CHANGE THIS DEPENDING ON DESIRED MINIMUM LENGTH OF ID (SYMBOLS, ALPHABETS, NUMBERS ALL INCLUSIVE)
                    split_item = item.split()
                    for elem in split_item:                        
                        for_IDs_elem = elem.translate(str.maketrans('', '', string.punctuation))
                        
                        
                        if len(for_IDs_elem) > 5: #CHANGE THIS DEPENDING ON DESIRED MINIMUM LENGTH OF ID (SYMBOLS, ALPHABETS, NUMBERS ALL INCLUSIVE)
                            if not for_IDs_elem.isalpha():
                                if for_IDs_elem not in all_possible_IDs:
                                    all_possible_IDs.append(for_IDs_elem)
                       
                        if len(for_IDs_elem) > 5: #CHANGE THIS DEPENDING ON DESIRED MINIMUM LENGTH OF ID (SYMBOLS, ALPHABETS, NUMBERS ALL INCLUSIVE)                 
                            if for_IDs_elem.isnumeric():
                                if item not in all_possible_IDs:
                                    all_possible_IDs.append(item)
                                    
    #print ('\n')
    #print (all_possible_IDs)          
            
            if '19' in get_text:
                new_get_text = get_text.replace('-',' ')
                new_get_text = new_get_text.replace('/',' ')
                __DOBs.append(new_get_text)
            
            if get_text not in draft_all_texts:
                if get_text not in all_possible_IDs:
                    all_texts.append(get_text)
                    draft_all_texts += ' ' + get_text

    all__texts = []
    
    for item in all_texts:
        splits = item.split()
        new_item = ''
        for ix,elem in enumerate(splits):
            new_element = elem.capitalize()
            if ix == 0:
                new_item += new_element
            if ix > 0:
                new_item += ' ' + new_element
        all__texts.append(new_item)
    
    #
    
    final_all_texts = []
        
    for stringed in all__texts:
        translator = str.maketrans(string.punctuation, ' '*len(string.punctuation)) #map punctuation to space    
        new_stringed = stringed.translate(translator)
        final_all_texts.append(new_stringed)
    
    print ('ALL TEXTS', final_all_texts)
    print ('\n')
    
    #
    
    text =' '.join(final_all_texts)

    #print ('TEXT', text)
    #print ('\n')

    #no punctuation version of original text
        
    translator = str.maketrans(string.punctuation, ' '*len(string.punctuation)) #map punctuation to space    
    no_punc_text = text.translate(translator)
    
    print ('NO PUNC TEXT', no_punc_text)
    print ('\n')
    
    #Spacy to remove GPE, LANGUAGE, LOC
    #text is the original text

    nlp = spacy.load('en_core_web_sm')
        
    classify = nlp(no_punc_text)
    
    entities=[(i, i.label_, i.label) for i in classify.ents]
    
    #print (entities)
    
    list_remove=[]
    
    for item in entities:
        condition = False
        if item[1] == 'GPE' or item[1] == 'LANGUAGE' or item[1] == 'LOC':
            condition = True
        if condition:
            list_form=list(item)
            to_remove=str(list_form[0])
            list_remove.append(to_remove)
    
    worded_form = text.split()
    
    #print (worded_form)
    
    no_locations = []
    
    for element in worded_form:
        if element not in list_remove:
            no_locations.append(element)
    
    #print (no_locations)
    
    new__text = []
    
    for item in no_locations:
        if item not in new__text:
            new__text.append(item)
            
    new_text = ' '.join(new__text)
    
    #print (new_text)
    
    
    # In[77]:
    
    
    #REMOVE CONSONANT-ONLY WORDS
    
    approved_words = []
    
    for word in new__text:
        if word.isnumeric():
            approved_words.append(word)
        if word.isalnum():
            if not word.isalpha():
                if not word.isnumeric():
                    approved_words.append(word)
        if word.isalpha():
            for char in word:
                if char in 'aeiouAEIOU':
                    if word not in approved_words:
                        approved_words.append(word)
    
    new_text = ' '.join(approved_words)
    
    #print (new_text)
    
    
    # In[78]:
    
    
    #NLTK to identify English and anglicised Asian names
        
    list_output = list(ne_chunk(pos_tag(word_tokenize(new_text))))
    
    #print (list_output)
    #print ('\n')
    
    new_output=[]
    
    for item in list_output:
        list_item=list(item)
        new_output.append(list_item)
    
    #print (new_output)
    
    neww_output=[]
    
    for listed in new_output:
        listt=[]
        for bracket in listed:
            if type(bracket) == tuple:
                listed_bracket=list(bracket)
                listt.append(listed_bracket)   
            else:
                listt.append(bracket)
        neww_output.append(listt)
    
    #print (neww_output)
    #print ('\n')
    
    double_apostrophe = json.dumps(neww_output)
    
    final_pairs_string = str(double_apostrophe)
    
    #print (final_pairs_string)
    
    
    # In[79]:
    
    
    #SEGREGATE NAMES
    
    final_pairs_string = str(final_pairs_string)
    
    text_file_1 = open(r"L:\My Documents\Desktop\out.txt", "w")
    text_file_1.write("%s" % final_pairs_string)
    text_file_1.close()
    
    for i, line in enumerate(fileinput.input(r"L:\My Documents\Desktop\out.txt", inplace=1)):
        sys.stdout.write(line.replace('[', ''))
    
    for i, line in enumerate(fileinput.input(r"L:\My Documents\Desktop\out.txt", inplace=1)):
        sys.stdout.write(line.replace(']', ''))
       
    for i, line in enumerate(fileinput.input(r"L:\My Documents\Desktop\out.txt", inplace=1)):
        sys.stdout.write(line.replace('(', ''))
    
    for i, line in enumerate(fileinput.input(r"L:\My Documents\Desktop\out.txt", inplace=1)):
        sys.stdout.write(line.replace(')', ''))
       
    for i, line in enumerate(fileinput.input(r"L:\My Documents\Desktop\out.txt", inplace=1)):
        sys.stdout.write(line.replace('"', ''))
    
    for i, line in enumerate(fileinput.input(r"L:\My Documents\Desktop\out.txt", inplace=1)):
        sys.stdout.write(line.replace(', ', ','))
       
    if path.exists(r"L:\My Documents\Desktop\out.csv"):
        os.remove(r"L:\My Documents\Desktop\out.csv")
    os.rename(r"L:\My Documents\Desktop\out.txt", r"L:\My Documents\Desktop\out.csv")
    
    with open(r"L:\My Documents\Desktop\out.csv") as inputfile:
        reader = csv.reader(inputfile)
        output_1 = list(reader)
    
    a_list_output_1 = []
    
    for thing in output_1:
        for things in thing:
            a_list_output_1.append(things)
           
    #print (a_list_output_1)
    
    filtered = list(filter(None, a_list_output_1))
    
    #print (filtered)
    #print ('\n')
    
    def divide_chunks(l, n):
        # looping till length l
        for i in range(0, len(l), n):  
            yield l[i:i + n]
    
    x = list(divide_chunks(filtered, 2))
    
    #print (x)
    #print ('\n')
    
    independent_names = []
    all_fresh_names = []
    
    for idx,item in enumerate(x):
        fresh_names = []
        if idx == 0:
            if 'NNP' in item:
                if 'NNP' in x[idx+1]:
                    fresh_names.append(x[idx][0])
                else:
                    independent_names.append(x[idx][0])
        if idx > 0 and idx < len(x)-1:
            if 'NNP' in item:
                if 'NNP' not in x[idx-1] and 'NNP' in x[idx+1]:
                    fresh_names.append(x[idx][0])
                if 'NNP' in x[idx-1]:
                    fresh_names.append(x[idx][0])
                if 'NNP' not in x[idx-1] and 'NNP' not in x[idx+1]:
                    independent_names.append(x[idx][0])
        all_fresh_names.append(fresh_names)
    
    ready_for_en_removal = []
    
    for names in all_fresh_names:
        if len(names) != 0:
            ready_for_en_removal.append(names)
        if len(names) == 0:
            ready_for_en_removal.append(['---'])
    
    segregated_names = []
    
    for listed in ready_for_en_removal:
        for item in listed:
            segregated_names.append(item)
    
    #print (segregated_names)
    #print ('\n')
    #print (independent_names)
    
    
    # In[80]:
    
    
    #EXTRACT FIRST NAMES, LAST NAMES (includes anglicised english names)    
    #REMOVE ALL ENGLISH WORDS (and standalone numbers) FOUND IN ENGLISH DICTIONARY, EXCEPT NORMAL ENGLISH NAMES
    
    df = pd.read_excel(r"L:\My Documents\Desktop\first names - dataworld.xlsx") #change file name and path where necessary
    en_names=df['Name'].values.tolist()
    
    #print (en_names)
    
    d = enchant.Dict("en_US") #removes words found in English dictionary, and standalone digitised numbers
    
    en_to_remove=[]
        
    for ix,item in enumerate(segregated_names):
        if item == True or item == False:
            segregated_names.pop(ix)
    
    for ix,word in enumerate(segregated_names):
        if us.check(word) or uk.check(word):
            if word not in en_names: #AN ENGLISH WORD THAT IS NOT AN ENGLISH NAME
                en_to_remove.append(ix)
    
    #print (en_to_remove)
    
    for item in en_to_remove:
        segregated_names[item]=''
    
    #print (segregated_names)
    
    filteredsets = list(filter(None, segregated_names))
    
    #print (filteredsets)
    #print ('\n')
    
    for idx,item in enumerate(filteredsets):
        if item == '---':
            filteredsets[idx] = ''
    
    filtereddsets = '_'.join(filteredsets)
    
    #print (filtereddsets)
    #print ('\n')
    
    new_string = ''
    
    for ix,char in enumerate(filtereddsets):
        if ix > 0 and ix < len(filtereddsets)-1:
            if char == '_':
                if filtereddsets[ix-1] == '_' or filtereddsets[ix+1] == '_':
                    char = ' '
        if char == '_':
            if ix == 0:
                if filtereddsets[ix+1] == '_':
                    char = ' '
        if char == '_':
            if ix == len(filtereddsets)-1:
                if filtereddsets[ix-1] == '_':
                    char = ' '
        new_string += char
    
    #print (new_string)
    #print ('\n')
    
    full_segregated_names = new_string.split()
    
    #print (full_segregated_names)
    #print ('\n')
    
    fulled_segregated_names = []
    
    for item in full_segregated_names:
        items = item.split('_')
        fulled_segregated_names.append(items)
    
    fulls_segregated_names = []
    
    for elem in fulled_segregated_names:   
        appent = list(filter(None, elem))
        fulls_segregated_names.append(appent)
        
    #print (fulls_segregated_names)
    
    
    # In[81]:
    
    
    #EXTRACT INDEPENDENT NAMES
    
    en_to_remove_2=[]
        
    for ix,item in enumerate(independent_names):
        if item == True or item == False:
            independent_names.pop(ix)
    
    for ix,word in enumerate(independent_names):
        if us.check(word) or uk.check(word):
            if word not in en_names: #AN ENGLISH WORD THAT IS NOT AN ENGLISH NAME
                en_to_remove_2.append(ix)
    
    #print (en_to_remove)
    
    for item in en_to_remove_2:
        independent_names[item]=''
    
    #print (independent_names)
    
    filteredsets2 = list(filter(None, independent_names))
    
    #print (filteredsets2)
    
    for name in filteredsets2:
        fulls_segregated_names.append([name])
    
    #print (fulls_segregated_names)
    
    
    # In[82]:
    
    
    #EXTRACT DATES OF BIRTH
    #All punctuations, including '-' and '/' that separates dates, would have already been removed in the very first tokenisation
            
    DOBs = []

    for a_date in full_dates:
        if ' '+a_date+' ' in no_punc_text:
            if a_date not in DOBs:
                DOBs.append(a_date)    
    
    for date in __DOBs:
        if date not in DOBs:
            DOBs.append(date)
    
    final_DOBs = []
    
    for item in DOBs:
        if len(item) > 5:
            split_item = item.split()
            check_joined = ''.join(split_item)
            for elem in split_item:
                if check_joined.isnumeric():
                    if item not in final_DOBs:
                        if len(item) < 13:
                            final_DOBs.append(item)
                if elem in all_months_spelling:
                    if item not in final_DOBs:
                        if len(item) < 13:                        
                            final_DOBs.append(item)
                    
    finalised_full_names = []
    
    for names in fulls_segregated_names:
        for item in names:
            #print ('ITEM', item)
            for elem in final_all_texts:
                #print ('ELEM', elem)
                if item in elem:
                    finalised_full_names.append(elem)
                else:
                    if item not in finalised_full_names:
                        finalised_full_names.append(item)
                 
    en_to_remove1 = []
     
    for ix,name in enumerate(finalised_full_names):
        split_name = name.split()
        for indiv in split_name:
            if us.check(indiv) or uk.check(indiv):
                if indiv not in en_names:
                    if indiv not in all_surnames: #AN ENGLISH WORD THAT IS NOT AN ENGLISH NAME
                        en_to_remove1.append(ix)
 
    for item in en_to_remove1:
        finalised_full_names[item]=''
         
 #    filtered1 = list(filter(None, final_first_names2))
    
    finalised_full_names = [x for x in finalised_full_names if x]
    finalised_full_names = list(dict.fromkeys(finalised_full_names))
    
    def num_there(s):
        return any(i.isdigit() for i in s)
    
    final_full_names = []
    
    for item in finalised_full_names:
        if num_there(item) == False:
            final_full_names.append(item)
    
    print ('FINAL FULL NAMES', final_full_names)    
    print ('\n')

###
    
    duplicated_final_full_names = final_full_names
    
    all_matching = []
    
    for thing in duplicated_final_full_names:
        for another in final_full_names:
            if thing in another:              
                if thing != another:
                    all_matching.append(another)
    
    all_matching = list(dict.fromkeys(all_matching))
        
    new_all_matching = all_matching
    
    for thing in duplicated_final_full_names:
        if any(thing in s for s in all_matching):
            pass
        else:
            new_all_matching.append(thing)
    
###    
            
    final_names = []
    
    for ix,item in enumerate(new_all_matching):
        names = item.split()
        if names[0] in all_surnames:
            names.append(names[0])
            del names[0]
            final_names.append(names)
        else:
            final_names.append(names)
    
    #print ('\n')           
    #print ('FINAL NAMES', final_names)
    #print ('\n')


    all_first_name_last_name = []
    
    for naming in final_names:
        if len(naming) > 2:
            long_name = ' '.join(naming[:-1])
            all_first_name_last_name.append([long_name, naming[-1]])
        else:
            all_first_name_last_name.append(naming)
 
    #print ('\n')            
    #print ('all first name last name', all_first_name_last_name)
    #print ('\n')
    
    first_name_last_name = []
    
    for listed in all_first_name_last_name:
        if len(listed) == 1:
            listed.append('')
        first_name_last_name.append(listed)
        
    #print ('first name, last name', first_name_last_name)

    first_names = []
    last_names = []
    
    for pairs in first_name_last_name:
        first_names.append(pairs[0])
        last_names.append(pairs[1])

    print ('FINAL FIRST NAMES', first_names)
    print ('\n')
    print ('FINAL LAST NAMES', last_names)
    print ('\n')
    
    for_validating = []
  
    for element in final_DOBs:
        to_validate_ = element.split()
        to_validate = ''.join(to_validate_)
        for_validating.append(to_validate)

    print ('FINAL DOBs', final_DOBs)       
    print ('\n')

#
    
    final__IDs = list(dict.fromkeys(all_possible_IDs))
    
    final_IDs = []
    
    for item in final__IDs:
        new_item = item.upper()
        for_checking = new_item.translate(str.maketrans('', '', string.punctuation))
        if for_checking not in for_validating:
            if for_checking not in final_IDs:
                final_IDs.append(for_checking)
    
    finalised_IDs = []

    for ix,item in enumerate(final_IDs):
        new_item = item.split()
        condition_ok = True
        for entity in new_item:
            if entity.isalpha():
                condition_ok = False
        if condition_ok:
            finalised_IDs.append(item)
        
    print ('FINAL IDs', finalised_IDs)       
    print ('\n')

#    
    
    df1 = pd.DataFrame({'Full Name': new_all_matching})
    df2 = pd.DataFrame({'First Name': first_names})
    df3 = pd.DataFrame({'Last Name': last_names})
    df4 = pd.DataFrame({'Date of Birth': final_DOBs})
    df5 = pd.DataFrame({'ID': finalised_IDs})
        
    
    ### FINAL OUTPUT OF CODE IS IN THE FOLLOWING FILE ###
    
    
    xlsx_filename = 'L:\\My Documents\\Desktop\\4 May\\Final Details (' + this_file_name + ', ' + dateStr + time + ').xlsx'
    
    writer = pd.ExcelWriter(xlsx_filename, engine='xlsxwriter')
    
    # Write each dataframe to a different worksheet.
    df1.to_excel(writer, sheet_name='Sheet1', startcol=0, index = False)
    df2.to_excel(writer, sheet_name='Sheet1', startcol=1, index = False)
    df3.to_excel(writer, sheet_name='Sheet1', startcol=2, index = False)
    df4.to_excel(writer, sheet_name='Sheet1', startcol=3, index = False)
    df5.to_excel(writer, sheet_name='Sheet1', startcol=4, index = False)
    
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()    
    
# =============================================================================
#     # Autofit excel cell widths
#     
#     if final_full_names and first_names and last_names and final_DOBs and final_IDs:
#         app = xw.App()
#         wb = xw.Book(xlsx_filename)
#         ws1 = wb.sheets['Sheet1']
#         ws1.autofit()
#         wb.save()
#         app.quit()
# =============================================================================

    #to check output of code
    
    txt_filename = 'L:\\My Documents\\Desktop\\Details at a Glance (' + this_file_name + ', ' + dateStr + time + ').txt'
    text_file_2 = open(txt_filename, "w")
    text_file_2.write("%s" % str(new_all_matching)) #full names
    text_file_2.write("\n")    
    text_file_2.write("%s" % str(first_names))
    text_file_2.write("\n")
    text_file_2.write("%s" % str(last_names))
    text_file_2.write("\n")
    text_file_2.write("%s" % str(final_DOBs))
    text_file_2.write("\n")
    text_file_2.write("%s" % str(finalised_IDs))
    text_file_2.write("\n")

    text_file_2.close()    
    
    conditioner1 = False
    conditioner2 = False
    conditioner3 = True
    conditioner4 = True
    
    if len(final_DOBs) == 0 or len(finalised_IDs) == 0:
        conditioner1 = True
        
    if len(final_DOBs) == 0 and len(finalised_IDs) == 0:
        conditioner2 = True
    
    if conditioner1 and conditioner2:
        print ('CLASSIFY AS DOCUMENT')
        conditioner3 = False
        conditioner4 = False
    
    if len(new_all_matching) < 21 and conditioner3:
        print ('CLASSIFY AS ID')
        conditioner4 = False
    
    if conditioner4: #everything else
        print ('CLASSIFY AS DOCUMENT')