import os
import win32com.client

from tkinter import messagebox
import time


def print_debug(debug, message):
    if debug:
        print(message)

def listFiles(directory, extension=".docx"):
    """
    :param directory: directory to parse
    :param extension: extension to look for
    :return: listFiles, array with files
    """
    # list all files
    files = os.listdir(directory)
    listFiles = []
    for file in files:
        if file.endswith(extension):
            listFiles.append(file)
    return listFiles
    # return (files, listFiles)


def ensure_file_is_closed_and_exists(file, debug=False):
    command_executed = False
    error = False
    while not command_executed:
        if os.path.exists(file):
            try:
                os.rename(file, file)
                print_debug(debug, 'Access on file "' + file + '" is available!')
                time.sleep(1)
                if error:
                    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                command_executed = True
            except OSError as e:
                print('Access-error on file "' + file + '"! \n' + str(e))
                messagebox.showinfo(title="Script de correction automatique",
                                    message="Fermer la presentation " + file + " pour la nouvelle correction")
                error = True
        else:
            print("file don't exist", file)
            exit(2)


def open_presentation(ppt_app, file, debug=False):

    # self.app = win32com.client.Dispatch("PowerPoint.Application")
    # Open presentation
    presentation = ""
    print_debug(debug, file)
    filename = os.path.abspath(file)
    print_debug(debug, filename)
    try:
        presentation = ppt_app.Presentations.Open(filename)
    except Exception as e:
        print("erreur dans l'ouverture de la presentation" + str(e))
        quit(2)

    print_debug(debug, "ok, presentation open")
    return presentation


def close_powerpoint(debug):
    try:
        # Create a Word application object
        ppt_app = win32com.client.Dispatch("Powerpoint.Application")
        # quit without saving
        ppt_app.Quit(SaveChanges=0)

        print_debug(debug, "Ppt closed successfully")

    except Exception as e:
        print(f"Error: {e}")

def main():
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    files = listFiles(".", ".pptx")
    print(files)
    presentation = open_presentation(ppt_app, files[0], debug=True)

    ppt_app.Quit()
if __name__ == "__main__":
    main()
