

class txt():
    def __init__(self,filePathWithName):

        self.settingMap = {}
        try:
            self.file = open(filePathWithName,"r+")
        except FileNotFoundError as _:
            self.file = open(filePathWithName,"w+")


    def readSetting(self):
        line = None
        settingKey = None
        tempMap = None
        tempList = []
        while line != "":
            line = self.file.readline()

            try:
                start = line.index("#")
                if tempMap is not None:
                    self.settingMap[settingKey] = tempMap
                settingKey = self.SuperStrip(line[start:],"r")

                if tempList.__len__() > 0:
                    self.settingMap[settingKey]["list"] = tempList
                    tempList = []
                tempMap = {}


            except Exception as e:
                if line.find("=") != -1:
                    lineArray = line.split("=")
                    key = self.SuperStrip(self.SuperStrip(lineArray[0],"r"),"l")
                    value = self.SuperStrip(self.SuperStrip(lineArray[1],"r"),"l")

                    tempMap[key] = value
                elif line.__len__() > 1:
                    # print(line)
                    tempList.append(self.SuperStrip(self.SuperStrip(line,"r"),"l"))
                    # print(tempList)
                    continue

        if tempMap is not None:
            self.settingMap[settingKey] = tempMap
        if tempList.__len__() > 0:
            self.settingMap[settingKey]["list"] = tempList

        print(self.settingMap)





    def initSetting(self):
        pass

    def SuperStrip(self,Instring:str,rightORLeft):

        if rightORLeft == "r":
            Instring = Instring.rstrip("\n") if Instring[-1:] == "\n" else Instring
            Instring = Instring.rstrip("\t") if Instring[-1:] == "\t" else Instring
            # print(Instring)

            while Instring[-1:] == " ":
                Instring = Instring.rstrip(" ")
        else:
            Instring = Instring.lstrip("\n") if Instring[-1:] == "\n" else Instring
            Instring = Instring.lstrip("\t") if Instring[-1:] == "\t" else Instring

            while Instring[:1] == " ":
                Instring = Instring.lstrip(" ")
        return Instring



if __name__ == "__main__":

    f = txt("./a.txt")
    f.readSetting()
