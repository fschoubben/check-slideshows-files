from powerpoint_macros import *
from check_slideshows_tools import *
from student import Student

def main():
    stud = Student()
    # Create Powerpoint instance
    #ppt_app = win32com.client.Dispatch("PowerPoint.Application")

    py_win32_ppt_app = win32com.client.Dispatch("PowerPoint.Application")

    # on prends la liste des fichiers PDF
    lf = listFiles(".", ".pptx")
    print("files array : ", lf)
    # print(listefichiers)
    students=[]
    for f in lf:
        stud = Student()

        max = 0

        presentation = open_presentation(py_win32_ppt_app, f)
        add_macros(presentation)


        # check filename
        # voir par quoi c'est généré : Word ou LibO
        # nom fichier :  2023-01-TIC1—Nom- /2
        #verifier_nom_fichiers(f, default_start_of_filename, stud)
        #max += 2
        # check 2 formats
        #verifier_deux_formats_fichiers(f, listefichiers, 2, stud.scores, stud.reasons)
        #max += 2

        m=0
        tot_pages=0
        to_check=""
        try:
            check_shapes(py_win32_ppt_app, stud)
            max+=stud.max_points["slideshowObjectType"]
            check_animations(py_win32_ppt_app, stud, key="slideshowAnimation")
            max+=stud.max_points["slideshowAnimation"]

        except Exception as e:
            print("problème dans le check_shapes : No pdf file ? ")
        print(stud.firstname, " ", stud.name, " ", str(sum(stud.scores.values())), "sur ", max)
        print("========================================")
        students.append(stud)
        # time.sleep(5)
    py_win32_ppt_app.Quit()

    # generate xlsx results file
    #save_in_excel_file(students, groups)

    print("done")

if __name__ == "__main__":
    main()
