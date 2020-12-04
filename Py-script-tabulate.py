import pdfplumber
import tabula
import csv
import os


def getName(fp):
    name_beg = fp.rindex("Student Name : ") + 15
    name_end = fp.rindex("Semester : 8")
    name = (fp[name_beg:name_end]).strip()
    return name


def getUsn(fp):
    usn_beg = fp.rindex("University Seat Number : ") + 25
    usn_end = fp.rindex("Student Name : ")
    usn = (fp[usn_beg:usn_end]).strip("\n")
    return usn

def subcode(table, table_reval):
    subcode = list()
    try:
        for i in table['Subject']:
            subcode.append(i)
        for i in table_reval['Subject']:
            subcode.append(i)
    except:
        return subcode
    return subcode


def Internal(table, table_reval):
    InternalMarks = list()
    try:
        for i in table['Internal']:
            InternalMarks.append(int(i))
        for i in table_reval['Internal']:
            InternalMarks.append(int(i))
    except:
        return InternalMarks
    return InternalMarks


def External(table, table_reval):
    ExternalMarks = list()
    try:
        for i in table['External']:
            ExternalMarks.append(int(i))
        for i in table_reval['External']:
            ExternalMarks.append(int(i))
    except:
        return ExternalMarks
    return ExternalMarks


def Total(table, table_reval):
    TotalMarks = list()
    try:
        for i in table['Total']:
            TotalMarks.append(int(i))
        for i in table_reval['Total']:
            TotalMarks.append(int(i))
    except:
        return TotalMarks
    return TotalMarks


def Result(table, table_reval):
    Result = list()
    try:
        for i in table['Result']:
            Result.append(i)
        for i in table_reval['Result']:
            Result.append(i)
    except:
        return Result
    return Result


def getResultList(table, table_reval,fp):
    usn = getUsn(fp)
    name = getName(fp)
    subjectcode = subcode(table, table_reval)
    internal = Internal(table, table_reval)
    external = External(table, table_reval)
    total = Total(table, table_reval)
    result = Result(table, table_reval)
    l = len(internal)
    res = list()
    for i in range(l):
        resultList = list()
        resultList.append(usn)
        resultList.append(name)
        resultList.append(subjectcode[i])
        resultList.append(internal[i])
        resultList.append(external[i])
        resultList.append(total[i])
        resultList.append(result[i])
        resultList = tuple(resultList)
        res.append(resultList)
    return res
    # print(subjectcode[i] + " " + internal[i] + " " + external[i] + " " + total[i] + " " + result[i])
    #to print the result which is stored in the list as subjectcode, internal marks, external marks, total marks, result
    #print(resultList)


def openwithplumber(filename):
    with pdfplumber.open(filename) as pdf1:
        fp = pdf1.pages[0].extract_text()
        return fp

def openwithtabula(filename,fp):
    table = tabula.read_pdf(filename)[0].dropna()
    table_reval = tabula.read_pdf(filename)[1].dropna()
    result_list = getResultList(table,table_reval,fp)
    return result_list


def main():
    directory = r'\PDFs'
    print("Enter File Name: ")
    fname = input()
    with open(fname,'a',newline="") as f:
        # fieldNames = ['USN', 'Name', 'Subject Code', 'Internal Marks', 'External Marks', 'Total Marks', 'Result']
        thewriter = csv.writer(f)
        thewriter.writerow(['USN', 'Name', 'Subject Code', 'Internal Marks', 'External Marks', 'Total Marks', 'Result'])
        for entry in os.scandir(directory):
            print(entry)
            if (entry.path.endswith(".pdf") and entry.is_file()):   
                filename = entry.path
                fp = openwithplumber(filename)
                result_list = openwithtabula(filename,fp)
                for i in result_list:
                    thewriter.writerow([i[0],i[1],i[2],i[3],i[4],i[5],i[6]])

if __name__ == "__main__":
    main()
