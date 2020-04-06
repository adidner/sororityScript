from openpyxl import load_workbook

failingGirls = []
gpaMinimum = 2.0
nameGPAFile = "Base1.xlsx"
nameEmailFile = "Contact Sheet Format.xlsx"


def grabDataFromGPATable(targetfilename):

    workbook = load_workbook(filename=targetfilename)

    sheet = workbook.active

    firstNames = sheet["A"]
    lastNames = sheet["B"]
    GPAs = sheet["E"]





    for (first,last,gpa) in zip(firstNames, lastNames, GPAs):


        if isinstance(gpa.value,float) and gpa.value <= gpaMinimum:
            newObject = {
                "first": first.value.lower(),
                "last": last.value.lower(),
                "gpa": gpa.value
            }
            failingGirls.append(newObject)

    print(failingGirls)


function grabDataFromEmailTable(targetfilename):

    workbook2 = load_workbook(filename=targetfilename)

    sheet2 = workbook2.active


    firstNames2 = sheet2["D"]
    lastNames2 = sheet2["C"]
    emails = sheet2["G"]


    for currentGirl in failingGirls:
        for first, last, email in zip(firstNames2, lastNames2, emails):
            if currentGirl["first"] == first.value.lower() and currentGirl["last"] == last.value.lower():
                currentGirl["email"] = email.value
                sendEmail(currentGirl)


function sendEmail(currentGirl):
    print("NOT ACTUALLY SENDING EMAILS YET, still work to do here")



def main():
    grabDataFromGPATable(nameGPAFile)
    grabDataFromEmailTable(nameEmailFile)


if __name__ == '__main__':
    main()
