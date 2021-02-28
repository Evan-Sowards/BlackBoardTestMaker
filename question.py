# This is where I'll make the object classes for Python

class Question:
    def __init__(self, text="", r1 = "", r2 = "", r3 = "", r4 = "", r5 = "", r6 = "", r7 = "", r8 = "",
                 w1= "", w2 = "", w3 = "", w4 = "", w5 = "", w6 = "", w7 = "", w8 = "", correct = "", incorrect = "",
                 num_right = 0):

        self.text = text
        self.r1 = r1
        self.r2 = r2
        self.r3 = r3
        self.r4 = r4
        self.r5 = r5
        self.r6 = r6
        self.r7 = r7
        self.r8 = r8
        self.w1 = w1
        self.w2 = w2
        self.w3 = w3
        self.w4 = w4
        self.w5 = w5
        self.w6 = w6
        self.w7 = w7
        self.w8 = w8
        self.correct = correct
        self.incorrect = incorrect
        self.num_right = num_right

    def print_question(self):
        print(self.text + "\n")
        print(self.r1 + "\n")
        print(self.r2 + "\n")
        print(self.r3 + "\n")
        print(self.r4 + "\n")
        print(self.r5 + "\n")
        print(self.r6 + "\n")
        print(self.r7 + "\n")
        print(self.r8 + "\n")
        print(self.w1 + "\n")
        print(self.w2 + "\n")
        print(self.w3 + "\n")
        print(self.w4 + "\n")
        print(self.w5 + "\n")
        print(self.w6 + "\n")
        print(self.w7 + "\n")
        print(self.w8 + "\n")
        print(self.correct + "\n")
        print(self.incorrect + "\n")
        print(str(self.num_right) + "\n")




