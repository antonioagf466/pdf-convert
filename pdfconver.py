#this is only supposed to work if you have microsof word in your machine
from docx2pdf import convert
import os

def main():
    folder_path = 'D:/pdfsconvert/toconvert' #change this to whatever path you need in your pc
    
    folder_goal = 'D:/pdfsconvert/converted' #change this to whatever path you need in your pc
    
    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            file_extension = filename.split('.')[-1].lower()
            if file_extension == "docx":
               file_path = os.path.join(folder_path, filename)
               convert(file_path, folder_goal)
    
    print(f"The files were converted and moved from {folder_path} to {folder_goal}")


if __name__ == "__main__":
    main()