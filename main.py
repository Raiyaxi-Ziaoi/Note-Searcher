import os
from docx import Document


def search(keyword):
    file_list = []
    found = 0
    for root, dirs, files in os.walk(os.getcwd()):
        for file in files:
            if os.path.join(root, file).endswith(".docx") or os.path.join(root, file).endswith(".txt"):
                file_list.append(os.path.join(root, file))
    for file in file_list:
        if file.endswith(".docx"):
            with open(file, "r", encoding='cp850') as f:
                file_contents = f.read()
                file_contents = file_contents.split("\n")
                lineIndex = 1
                doc = Document(file)
                for para in doc.paragraphs:
                    text = para.text
                    text = text.split("\n")
                    for line in text:
                        if keyword in line:
                            filePath = os.path.relpath(file, os.getcwd())
                            print(
                                f"\nFile: \"{filePath}\" Line: \"{lineIndex}\"")
                            lineIndex += 1
                            found += 1
                        else:
                            lineIndex += 1
                            break
        else:
            with open(file, "r", encoding='cp850') as f:
                file_contents = f.read()
                file_contents = file_contents.split("\n")
                lineIndex = 1
                for line in file_contents:
                    if keyword in line:
                        filePath = os.path.relpath(file, os.getcwd())
                        print(f"\nFile: {filePath}, Line: {lineIndex}")
                        lineIndex += 1
                        found += 1
                    else:
                        lineIndex += 1
                        break
    if found == 0:
        print("\nKeyword not found")


os.system('TITLE Note Searcher')
os.system('COLOR 0b')
print("\nWelcome to Note Searcher")
while True:
    print("\nType exit to exit")
    keyword = input("\nPlease enter a keyword to search for: ")
    search(keyword)
print("\nThank you for using Note Searcher")
input("\nPress enter to exit...")
