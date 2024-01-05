import sys
import os
import win32com.client

from tkinter import messagebox
import time

import re
# Documentation :
#  --> need to authorise macros in Powerpoint : https://support.pcmiler.com/en/support/solutions/articles/19000047036-what-does-error-1004-programmatic-access-to-visual-basic-project-is-not-trusted-mean-

remove_non_english = lambda s: re.sub(r'[^a-zA-Z0-9]', '', s)


### Macro's Definition
def define_macros():
    macros = []
    macros.append("""
Function CheckMasterSlideName(targetName As String) As Boolean
    ' Define the name to check
    ' Dim targetName As String
    ' targetName = "Test"  ' Replace with the actual name you're looking for

    ' Iterate through each slide master
    'Dim master As master
    Dim presDesign As Design
    Dim CL As CustomLayout
    Dim shp As Shape
    Dim mName As String, mCode As String, mPhone As String, fName As String
    For Each presDesign In ActivePresentation.Designs
        For Each CL In presDesign.SlideMaster.CustomLayouts
            ' check if name is in sub-slidemasters
            For Each shp In CL.Shapes
                If shp.HasTextFrame Then
                    If InStr(1, shp.TextFrame.TextRange, targetName, vbTextCompare) > 0 Then
                        ' The name is found
                        CheckMasterSlideName = True
                        Exit Function
                    End If
                End If
            Next shp
            For Each shp In presDesign.SlideMaster.Shapes
            ' check if name is in "master slideMaster"
                If shp.HasTextFrame Then
                    If InStr(1, shp.TextFrame.TextRange, targetName, vbTextCompare) > 0 Then
                        ' The name is found
                        CheckMasterSlideName = True
                        Exit Function
                    End If
                End If
            Next shp
        Next CL
    Next presDesign

    ' The name is not found
    CheckMasterSlideName = False
End Function""")
    macros.append("""
Function ReturnShapeTypes() As String
    Dim slide As slide
    Dim shape As shape
    Dim shapeTypes As Collection
    Dim result As String

    ' Initialize a collection to store unique shape types
    Set shapeTypes = New Collection

    ' Iterate through each slide
    For Each slide In ActivePresentation.Slides
        ' Iterate through each shape on the slide
        For Each shape In slide.Shapes
            ' Check if the shape type is not already in the collection
            On Error Resume Next
            shapeTypes.Add shape.Type, CStr(shape.Type)
            On Error GoTo 0
        Next shape
    Next slide

    ' Build the result string
    result = ""
    For Each item In shapeTypes
        result = result & item & ", "
    Next item

    ' Remove the trailing comma and space
    result = Left(result, Len(result) - 2)

    ' Return the result
    ReturnShapeTypes = result
End Function""")
    macros.append("""
Function HasAnimation() As Boolean
    Dim slide As slide
    Dim shape As shape
    Dim hasAnim As Boolean

    ' Initialize to False
    hasAnim = False

    ' Iterate through each slide
    For Each slide In ActivePresentation.Slides
        ' Check each shape on the slide for animation
        For Each shape In Slide.Shapes
            If shape.AnimationSettings.Animate <> msoFalse Then
                ' If at least one shape has animation, set the flag to True
                HasAnimation = True
                Exit Function
            End If
        Next shape
    Next slide

    ' Return the result
    HasAnimation = hasAnim
End Function""")
    macros.append("""
Function CheckTransitions() As Integer
    Dim slide As slide
    Dim transitionCount As Integer
    Dim result As Integer

    ' Initialize transition count and result
    transitionCount = 0
    result = 0

    ' Iterate through each slide
    For Each slide In ActivePresentation.Slides
        ' Check if the slide has a transition
        If slide.SlideShowTransition.EntryEffect <> ppAnimateLevelNone  Then
            ' Increment the transition count
            transitionCount = transitionCount + 1
        End If
    Next slide

    ' Set the result based on the transition count
    If transitionCount = 0 Then
        result = 0
    ElseIf transitionCount = 1 Then
        result = 2
    Else
        result = 1
    End If

    ' Return the result
    CheckTransitions = result
End Function""")

    return macros


def print_debug(debug, message):
    if debug:
        print(message)


def check_shapes(ppt_app, student, key="slideshowObjectType", debug=False):
    max_scores = student.max_points[key]
    why = ""
    to_check_manually = ""
    score = 0

    try:
        shapes_list = ppt_app.Run("ReturnShapeTypes")
        print_debug(debug, str(shapes_list))
        shapes=shapes_list.split(", ")
        shapes_types={ 1: "AutoShape", 3 : "Chart", 4 : "Comment", 5 : "Freeform", 7 : "Embedded OLE object", 8 : "Form control",
                       10 : "Linked OLE object", 11 : "Linked picture", 13 : "Picture", 14 : "PlaceHolder", 15 : "Text effect",
                       16 : "Media", 17 : "Text box", 19 : "Table", 20 : "Canvas", 21 : "Diagram", 24 : "SmartArt graphic",
                       26 : "Web video", 28 : "Graphic", 29 : "Linked graphic", 30 : "3D model", 31 : "Linked 3D model",
                       14 : "PlaceHolder"}
        # TODO : check what is 1 EXACTLY : Autoshape : "regular shape shapes" (square, circle...)
        # 6 : groups
        # 14 : is Placeholder, the default shape in some slides
        # shapes values : https://learn.microsoft.com/en-us/office/vba/api/office.msoshapetype
        accepted_values = [1, 3, 4, 5, 7, 8, 10, 11, 13, 15, 16, 17, 19, 20, 21, 24, 26, 28, 29, 30, 31]
        # TODO : next year ? object of shape 6 is a group
        # TODO 2024 : double check values, especially 1 and how to make difference between video and audio inserted
        shapes_nb=0
        for sh in shapes:
            try:
                sh=int(sh)
            except Exception as e:
                print("shape given is not an int ? ", str(e))
            if sh in accepted_values:
                shapes_nb += 1
                print_debug(debug, "shape found : "+shapes_types[sh])
        if shapes_nb >= 4:
            print_debug(debug, "OK, au moins 4 types d'objets présente. ")
            score = max_scores
        elif shapes_nb >= 2:
            print_debug(debug, "OK, mais seulement 2 ou 3 types d'objets présents. ")
            why += "moins de 4 types d'objets dans le document. "
            to_check_manually += "vérifier trtypes d'objets. "
            score = max_scores/2
        else:
            print_debug(debug, "pas de  transition")
            why += "moins de 2 type d'objets dans le document. "
            to_check_manually += "vérifier types d'objets. "
    except Exception as e:
        sys.stderr.write("error in word_macros.py\check_shapes " + str(e))

    student.scores[key] = score
    student.reasons[key] = why
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    print_debug(debug, "fin check_shapes ")
    return {}

def check_transitions(ppt_app, student, key="slideshowTransition", debug=False):
    max_scores = student.max_points[key]
    why = ""
    to_check_manually = ""
    score = 0

    try:
        transitions = ppt_app.Run("CheckTransitions")
        if transitions == 2:
            print_debug(debug, "OK, une seule transition présente. ")
            score = max_scores
        elif transitions == 1:
            print_debug(debug, "OK, mais plusieurs transition présente. ")
            why += "plus d'une transition dans le document. "
            to_check_manually += "vérifier transitions - "
            score = max_scores/2
        else:
            print_debug(debug, "pas de  transition")
            why += "pas de transition dans le document. "
            to_check_manually += "vérifier transitions - "
    except Exception as e:
        sys.stderr.write("error in word_macros.py\check_transitions " + str(e))

    student.scores[key] = score
    student.reasons[key] = why
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    print_debug(debug, "fin check_transitions ")
    return {}
def check_name_in_mask(ppt_app, student, key="slideshowNameInTemplate", debug=False):
    # TODO : ensure it works with student files...
    max_scores = student.max_points[key]
    why = ""
    to_check_manually = ""
    score = 0

    try:
        name_in_mask = ppt_app.Run("CheckMasterSlideName", student.name)
        if name_in_mask:
            print_debug(debug, "OK, Nom présent dans le masque. ")
            score = max_scores
        else:
            print_debug(debug, "pas de Nom présent dans le masque'")
            why += "pas de Nom présent dans le masque du document. "
            to_check_manually += "vérifier masque - "
    except Exception as e:
        sys.stderr.write("error in word_macros.py\check_name_in_mask " + str(e))

    student.scores[key] = score
    student.reasons[key] = why
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    print_debug(debug, "fin check_name_in_mask ")
    return {}
def check_animations(ppt_app, student, key="slideshowAnimation", debug=False):
    max_scores = student.max_points[key]
    why = ""
    to_check_manually = ""
    score = 0

    try:
        animation = ppt_app.Run("HasAnimation")
        if animation:
            print_debug(debug, "OK, Animation présente. ")
            score = max_scores
        else:
            print_debug(debug, "pas d'animation'")
            why += "pas d'animation' dans le document. "
            to_check_manually += "vérifier animation - "
    except Exception as e:
        sys.stderr.write("error in word_macros.py\check_animations " + str(e))

    student.scores[key] = score
    student.reasons[key] = why
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    print_debug(debug, "fin check_animations ")
    return {}


# def add_word_macros_pywin32():
def add_macros(presentation, debug=False):
    macros = define_macros()
    try:
        new_module = presentation.VBProject.VBComponents.Add(1)  # 1 correspond à vbext_ct_StdModule

        # Ajouter le code de la macro au module
        for m in macros:
            new_module.CodeModule.AddFromString(m)
    except Exception as e:
        print(f"Une erreur s'est produite dans l'ajout de la macro : {e}")
    print_debug(debug, "macros ajoutées")
    # try:
    #    word_count = doc.Run("CompterMots")
    #    print(f"Le nombre de mots dans le document est : {word_count}")
    # except Exception as e:
    #    print(f"Une erreur s'est produite dans le retour de la macro : {e}")
    # finally:
    # doc.Close(True)
    # Fermer l'application Word
    # ppt_app.Quit()


def close_powerpoint(debug):
    try:
        # Create a Word application object
        ppt_app = win32com.client.Dispatch("Powerpoint.Application")
        # quit without saving
        ppt_app.Quit(SaveChanges=0)

        print_debug(debug, "Ppt closed successfully")

    except Exception as e:
        print(f"Error: {e}")

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
def open_presentation(file, debug=False):

    # Create Powerpoint instance
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    # self.app = win32com.client.Dispatch("PowerPoint.Application")
    # Open presentation

    try:
        presentation = ppt_app.Presentations.Open(file)
    except Exception as e:
        print("erreur dans l'ouverture de la presentation" + str(e))
        quit(2)

    print_debug(debug,"ok, presentation open")
    return ppt_app, presentation


def main():
    debugging = True
    # TODO : keep ?    close_powerpoint(debugging)
    file_name_begin = "2024-01-S2-"
    # file_name = file_name_begin + "Test-1.pptx"
    file_name = file_name_begin + "Test-1.pptx"
    file_name = file_name_begin + "Test-4.pptx" # no transition
    # file_name = file_name_begin + "Test-5.pptx"  # 2 transitions
    # file_name = file_name_begin + "Test-6.pptx"  # 1 transition
    file_name = file_name_begin + "Test-7.pptx"  # Test Name in mask (not "all slides")
    # file_name = file_name_begin + "Test-8.pptx"  # Test Name in mask "all slides"

    # file_name = file_name_begin + "Delsalle-Lisa-PowerPoint.pptx"
    # file_name = file_name_begin + "Arens-Hélène-Diapo.pwp.pptx"
    # file_name = "2024-01-S2-Henrotte-Clémence-Diapo.pptx"
    # file_name = "2024-01-S2-LEE-Chae-Yeon-diapo.pptx"

    ensure_file_is_closed_and_exists(file_name)

    stud = Student()
    stud.name = "Test"
    stud.firstname = "1"

    #print(file_name)
    file = file_name
    path = os.getcwd()
    file_name = path + '/' + file_name
    print(file_name)

    (ppt_app, presentation)=open_presentation(file_name, debug=debugging)

    add_macros(presentation)
    print_debug(debugging, "macros added : OK")
    #
    # check_quote(ppt_app, stud, key="citation", debug=debug)
    # check_animations(ppt_app, stud, key="slideshowAnimation", debug=debugging)
    # check_transitions(ppt_app, stud, key="slideshowTransition", debug=debugging)
    check_name_in_mask(ppt_app, stud, key="slideshowNameInTemplate", debug=debugging)
    #check_shapes(ppt_app, stud, key="slideshowObjectType", debug=debugging)

    for key, value in stud.scores.items():
        if value !=0:
            print(key, ": ", value, "/", stud.max_points[key])
    presentation.Close()
    ppt_app.Quit()
    
if __name__ == "__main__":
    from student import Student
    main()

