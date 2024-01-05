from powerpoint_macros import *
from check_slideshows_tools import *
from student import Student

working_directory = "."
excel_file_for_results = working_directory+"/"+"2024-01-auto-correct-slideshows-results.xlsx"


def main():
    stud = Student()
    # Create Powerpoint instance
    # ppt_app = win32com.client.Dispatch("PowerPoint.Application")

    py_win32_ppt_app = win32com.client.Dispatch("PowerPoint.Application")

    # on prends la liste des fichiers PDF
    lf = listFiles(".", ".pptx")
    print("files array : ", lf)
    # print(listefichiers)
    students = []
    for f in lf:
        stud = Student()
        files_parts = f.split("-")
        stud.name = files_parts[3]
        stud.firstname = files_parts[4]

        max_score = 0

        presentation = open_presentation(py_win32_ppt_app, f)
        add_macros(presentation)

        # check filename
        # voir par quoi c'est généré : Word ou LibO
        # nom fichier :  2023-01-TIC1—Nom- /2
        # verifiemax_scorer_nom_fichiers(f, default_start_of_filename, stud)
        # max_score += 2
        # check 2 formats
        # verifier_deux_formats_fichiers(f, listefichiers, 2, stud.scores, stud.reasons)
        # max_score += 2

        m = 0
        tot_pages = 0
        to_check = ""
        try:
            check_shapes(py_win32_ppt_app, stud)
            max_score += stud.max_points["slideshowObjectType"]
            check_animations(py_win32_ppt_app, stud, key="slideshowAnimation")
            max_score += stud.max_points["slideshowAnimation"]
            check_transitions(py_win32_ppt_app, stud, key="slideshowTransition")
            max_score += stud.max_points["slideshowTransition"]
            check_name_in_mask(py_win32_ppt_app, stud, key="slideshowNameInTemplate")
            max_score += stud.max_points["slideshowNameInTemplate"]

        except Exception as e:
            print("problème dans le check_shapes : No pdf file ? ")
        print(stud.name, " ", stud.firstname, " ", str(sum(stud.scores.values())), "sur ", max_score)
        print("========================================")
        students.append(stud)
        # time.sleep(5)
    py_win32_ppt_app.Quit()

    # generate xlsx results file
    # TODO : generate groups
    groups = {"S2", "Unknown"}
    save_in_excel_file(excel_file_for_results, students, groups)

    print("done")


if __name__ == "__main__":
    main()
