from operator import truediv
from pathlib import Path  #imports the path class from pathlib python package
from docx import Document #imports Document class from PyDocx package

folder_path = Path(r"C:\Users\lowry\Desktop\Unsorted Documents") #points to folder, rstring prevents backspace from causing issues
files = folder_path.iterdir()  #Returns an iterator of path objects for each file/ folder in Unsorted documents (.iterdir is from the path class
for file in files: #loops through files variable(iterator of path objects) temp naming each one file
    print(file)

file_path = Path(r"C:\Users\lowry\Desktop\Unsorted Documents\FIAR_SOW_Document.docx") #points to file
doc = Document(file_path) #doc becomes a document object from the pydocx package, holding everything inside the docx file, in the provided location

subject_keywords = {'Audit' : ('FIAR','audit','internal controls','NGA') #Dictionary containing subjects, and keywords that relate to them
                    }


def determine_subject(file):
    scores = {"Audit": 0, 'Budget': 0 }

    for paragraph in doc.paragraphs:  # loops through doc variable with .paragraphs attribute, which is a list of all the paragraphs of text (just paragraphs) in the document object
        print(paragraph.text) #prints paragraphs
        for subject, keywords in subject_keywords.items(): #loops through dictionary as a list using .items
            for keyword in keywords:  #checks (loops through) each individual keyword
                if keyword.lower() in paragraph.text.lower(): #checks if paragraph text matches any keywords during loop
                    # scores["Audit"] += 1
                    # for subject, score in scores.items():
                    #     print(f"{subject}: {score} matches")


""" What this means, Go through each paragraph, and check all subjects(keys) for keyword matches in that paragraph, then print good job everytime theres a match"""


determine_subject(file_path)

