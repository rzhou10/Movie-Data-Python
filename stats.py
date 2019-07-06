from itertools import islice
import sys

class Stats():
    grossAvg = 0
    maxGross = 0
    highestFilm = ""
    lowestFilm = ""
    minGross = sys.maxsize
    runAvg = 0
    longestFilm = ""
    shortestFilm = ""
    maxRun = 0
    minRun = sys.maxsize

    def setStats(self, statList):
        statsAsDict = {}

        for i in islice(statList, 1, None):
            if i not in statsAsDict:
                statsAsDict[i] = 0
            statsAsDict[i] += 1

        return statsAsDict

    def setGrossStats(self, grossList, titleList):
        totalGross = 0
        nonNA = 0

        noFirst = list(islice(grossList, 1, None))

        for i in range(len(noFirst)):
            if noFirst[i] != "N/A":
                valToInt = int(noFirst[i])
                totalGross += valToInt
                nonNA += 1

                if valToInt > self.maxGross:
                    self.maxGross = valToInt
                    self.highestFilm = titleList[i + 1]
                elif valToInt < self.minGross:
                    self.minGross = valToInt
                    self.lowestFilm = titleList[i + 1]
        
        self.grossAvg = round(totalGross / nonNA, 2)

    def getGrossAvg(self):
        return self.grossAvg

    def getMaxGross(self):
        return self.maxGross

    def getMinGross(self):
        return self.minGross
    
    def getHighestFilm(self):
        return self.highestFilm
    
    def getLowestFilm(self):
        return self.lowestFilm
    
    def setRunStats(self, runList, titleList):
        totalRun = 0

        noFirst = list(islice(runList, 1, None))
        
        for i in range(len(noFirst)):
            valToInt = int(noFirst[i])
            totalRun += valToInt

            if valToInt > self.maxRun:
                self.maxRun = valToInt
                self.longestFilm = titleList[i + 1]
            elif valToInt < self.minRun:
                self.minRun = valToInt
                self.shortestFilm = titleList[i + 1]

        self.runAvg = round(totalRun / len(noFirst), 2)
    
    def getRunAvg(self):
        return self.runAvg
    
    def getMaxRun(self):
        return self.maxRun
    
    def getMinRun(self):
        return self.minRun
    
    def getLongestFilm(self):
        return self.longestFilm
    
    def getShortestFilm(self):
        return self.shortestFilm