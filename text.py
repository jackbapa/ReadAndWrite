

class text():
    def __init__(self,filePathWithName):
        try:
            self.file = open(filePathWithName,"r+")
        except FileNotFoundError as _:
            self.file = open(filePathWithName,"w+")


    def readSetting(self):
        line = None
        while line is not "":
            line = self.file.readline()
            # print(line == "")
            start = line.index("#")
            print(start)
            # except ValueError as e:
                # print(e)
                # break




    def initSetting(self):
        pass


if __name__ == "__main__":

    f = text("d://a.text")
    f.readSetting()
