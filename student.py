class Student:
    max_points = {#"format": 2, "nomFichiers": 2,"poids": 2, "orthographe": 0, "pages": 0,
                  #"styles": 4, "piedDePage": 4, "espaces": 4, "TDM": 2, "section": 2, "listes": 2,
                  #"tableau": 2,  "citation": 2, "noteBasPage": 2, "lien": 2, "images": 4,
                  "slideshowObjectType": 4, "slideshowAnimation": 2, "slideshowTransition": 1,
                  "slideshowNameInTemplate":1, "slideshowOtherTypeOfSlideshow": 2}

    def reset(self):
        self.scores = {}
        self.reasons = {}
        for key in self.max_points:
            self.scores[key]=0
            self.reasons[key]=""

        self.name = ""
        self.firstname = ""
        self.group = "Unknown"
        self.to_check_manually = ""
        self.to_check = set()

    def __init__(self):
        self.reset()
if __name__ == "__main__":
    st=Student()
    total = 0
    i=0
    for key, value in st.scores.items():
       #print(key," --> ", value, "/", Student.max_points[key])
        print('{:<15}  --> {:>3} / {:>3}'.format(key, value, Student.max_points[key]))
        total+=st.max_points[key]
        i+=1
    for key, value in st.reasons.items():
        print(key, "-->", value)
    print(st.reasons)
    print(st.scores)
    print("max score = ",total,", ",i,"elements")