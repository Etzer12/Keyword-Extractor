"""
Name: Keyword Extractor.py
Author: Andrew Racine
Date: 3/31/17
Description: This module parses excel documents looking for specified keywords. It then outputs the frequency of the
keywords into a output file.
Requirements: xlrd, nltk,
Input: excel files where the 4th column has the data to be extracted.
Output: txt file with distribution data on keywords
Restrictions: Does not work with non excel files even if in the same csv format. It also will not work with other excel
files that aren't looking for the words specified in the dict. Changing the values in the dict may cause errors and is
not recommended.
"""

try:
    import xlrd
except ImportError:
    print "This algorithm requires xlrd to parse files. Please install xlrd and try again"
    quit()
import time
import operator

g_emotion_dict = {
                   "Yes 2 You": {},
                   "Khols Charge": {},
                   "Khols Cash": {},
                   "Referral": {},
                   "Friends and Family": {},
                   "Promotion": {}
}


def get_excel_sheets(path):
    xl_workbook = xlrd.open_workbook(path)
    sheet_names = xl_workbook.sheet_names()
    list_of_sheets = []
    for sheet in sheet_names:
        list_of_sheets.append(xl_workbook.sheet_by_name(sheet))
    return list_of_sheets


def create_keyword_dict():
    keywords = {
        "Missing points": 0,
        "Points expiring": 0,
        "Points expiration": 0,
        "Rewards": 0,
        "Birthday": 0,
        "Expiration": 0,
        "Points Balance": 0,
        "exchanges": 0,
        "returns": 0,
        "shipping": 0,
        "order status": 0,
        "redemption": 0,
        "redeem": 0,
        "stacking": 0,
        "stacking offers": 0,
        "donate points": 0,
        "Activate": 0,
        "MVC": 0,
        "negative points": 0,
        "gift card": 0,
        "Store": 0,
        "Online": 0,
        "Website": 0,
        "App": 0,
        "Mobile App": 0,
        "hate app": 0,
        "Kohl's app": 0,
        "Kohl's pay": 0,
        "Wallet": 0,
        "associates": 0,
        "account locked": 0,
        "mystery offer": 0,
        "secret sale": 0,
        "Double points": 0,
        "Triple points": 0,
        "coupon": 0,
         }

    return keywords


def create_sub_keyword_dict():
    keyword_dict = [create_keyword_dict(),
                     {
                         "Y2Y": 0,
                         "Yes2You Rewards": 0,
                         "Yes2You": 0,
                         "YesToYou": 0,
                         "Yes2u": 0,
                         "YesToU": 0,
                         "YesYou": 0
                     },
                     {
                         "Kohl's charge": 0,
                         "charge card": 0
                     },
                     {
                         "Kohl's Cash": 0,
                         "Kohl's Cash expiration": 0,
                         "Kohl's Cash expiring": 0,
                         "print Kohl's cash": 0,
                         "email kohl's cash": 0,
                         "print cash": 0,
                         "email cash": 0
                     },
                     {
                         "Referral": 0,
                         "Refer a friend": 0,
                         "Refer": 0
                     },
                     {
                         "friends and family": 0,
                         "friends & family": 0
                     },
                     {
                         "promotion": 0,
                         "promo code": 0
                     },
                     {
                         "love": 0,
                         "hate": 0,
                         "annoyed": 0,
                         "frustrated": 0,
                         "angry": 0,
                         "wonderful": 0,
                         "horrible": 0,
                         "nice": 0,
                         "helpful": 0,
                         "dissatisfied": 0,
                         "satisfied": 0,
                         "very satisfied": 0
                     }]

    return keyword_dict


def tally_results(curr_keywords, total_keywords):
    #  First do initial dict, then do the ones that need to only count once.
    for key, val in curr_keywords[0].iteritems():
        if val == 1:
            if key in total_keywords:
                total_keywords[key] += 1
            else:
                total_keywords[key] = 1

    lst_of_keywords = ["", "Yes 2 You", "Khols Charge", "Khols Cash", "Referral", "Friends and Family", "Promotion"]
    for x in range(1, 7):
        for key, val in curr_keywords[x].iteritems():
            if val == 1:
                if lst_of_keywords[x] in total_keywords:
                    total_keywords[lst_of_keywords[x]] += 1
                else:
                    total_keywords[lst_of_keywords[x]] = 1
                for key2, val2 in curr_keywords[7].iteritems():
                    if val2 == 1:
                        if key2 in g_emotion_dict[lst_of_keywords[x]]:
                            g_emotion_dict[lst_of_keywords[x]][key2] += 1
                        else:
                            g_emotion_dict[lst_of_keywords[x]][key2] = 1
                break

    return total_keywords


def main():
    start = time.time()
    file_path = raw_input("Enter Spreadsheet name: ")
    sheets = get_excel_sheets(file_path)
    keywords = create_sub_keyword_dict()
    output = open("{}_output.txt".format(file_path), "w")
    for dic in keywords:
        for word in dic:
            dic[word.lower().replace(" ", "").replace("'", "")] = dic.pop(word)
    total_keywords = {}
    total_comments_parsed = 0
    for sheet in sheets:
        col = sheet.col(3)
        for idx, cell_obj in enumerate(col):
            if idx == 0:
                continue
            total_comments_parsed += 1
            text = cell_obj.value.lower().replace(" ", "").replace("'", "").replace("\n", "")
            for dic in keywords:
                for word in dic:
                    if word in text:
                        dic[word] = 1
            total_keywords = tally_results(keywords, total_keywords)
            for dic in keywords:
                for word in dic:
                    dic[word] = 0
    total_keywords2 = sorted(total_keywords.items(), key=operator.itemgetter(1), reverse=True)
    output.write("{} comments parsed. Results:\n\nKeywords found and at what frequency:\n\n".format(total_comments_parsed))
    for key, val in total_keywords2:
        percent = round(((float(val)/total_comments_parsed) * 100), 2)
        output.write("{}: appeared {} times. {}% of all comments searched.\n".format(key, val, percent))
    for key, val in g_emotion_dict.iteritems():
        output.write("---------------------------------------------------------------\n")
        output.write(key + " emotion rates:\n")
        dic = sorted(val.iteritems(), key=operator.itemgetter(1), reverse=True)
        for key2, val2 in dic:
            percent = round(((float(val2) / total_keywords[key]) * 100), 2)
            output.write("{}: appeared {} times when {} was used. {}% of the time.\n".format(key2, val2, key, percent))
    end = time.time()
    print end - start


main()
