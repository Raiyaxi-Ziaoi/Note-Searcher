# Imports

import os
from docx import Document


def search(keyword):  # Searches
    file_list = []
    found = 0
    # Recursive directory search for files with .txt or .docx
    for root, dirs, files in os.walk(os.getcwd()):
        for file in files:
            if os.path.join(root, file).endswith(".docx") or os.path.join(root, file).endswith(".txt"):
                file_list.append(os.path.join(root, file))
    for file in file_list:
        if file.endswith(".docx"):  # File ends in .docx?
            with open(file, "r", encoding='cp850') as f:  # Open file with weird encoding
                file_contents = f.read()
                file_contents = file_contents.split("\n")
                lineIndex = 1
                doc = Document(file)
                for para in doc.paragraphs:  # Reads .docx file
                    text = para.text
                    text = text.split("\n")
                    for line in text:
                        if keyword in line:
                            filePath = os.path.relpath(file, os.getcwd())
                            print(
                                f"\nFile: \"{filePath}\" Line: \"{lineIndex}\"")  # Prints out details
                            lineIndex += 1
                            found += 1
                        else:
                            lineIndex += 1
                            break
        else:  # Must be .txt file now
            with open(file, "r", encoding='cp850') as f:  # More encoding magic
                file_contents = f.read()
                file_contents = file_contents.split("\n")
                lineIndex = 1
                for line in file_contents:
                    if keyword in line:
                        filePath = os.path.relpath(file, os.getcwd())
                        # Prints out details
                        print(f"\nFile: {filePath}, Line: {lineIndex}")
                        lineIndex += 1
                        found += 1
                    else:
                        lineIndex += 1
                        break
    if found == 0:  # Couldn't find keyword
        print("\nKeyword not found")


if __name__ == "__main__":  # Main
    os.system('TITLE Scrubber')
    os.system('COLOR 0b')
    print("Welcome to Scrubber")
    while True:
        print("\nType EXIT to exit the program.")
        keyword = input("\nPlease enter a keyword to search for: ")
        if keyword == "EXIT":
            break
        search(keyword)
    print("\nThank you for using Note Searcher")
    input("\nPress enter to exit...")
