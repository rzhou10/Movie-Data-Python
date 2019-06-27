import xlrd
from itertools import islice
import pandas as pd
from openpyxl import load_workbook
from stats import Stats

excelFile = "Movie Data.xlsx"
workbook = xlrd.open_workbook(excelFile)
sheet = workbook.sheet_by_index(0)

allTitles = [ str(sheet.cell_value(i, 0)) for i in range(sheet.nrows) ]
allYears = [ str(sheet.cell_value(i, 1)) for i in range(sheet.nrows) ]
allGenre = [ str(sheet.cell_value(i, 2)) for i in range(sheet.nrows) ]
allDirectors = [ str(sheet.cell_value(i, 3)) for i in range(sheet.nrows) ]
allRun = [ sheet.cell_value(i, 4) for i in range(sheet.nrows) ]
allGross = [ sheet.cell_value(i, 5) for i in range(sheet.nrows) ]
allRating = [ str(sheet.cell_value(i, 6)) for i in range(sheet.nrows) ]
allFran = [ str(sheet.cell_value(i, 7)) for i in range(sheet.nrows) ]
allStudio = [ str(sheet.cell_value(i, 8)) for i in range(sheet.nrows) ]

allStats = Stats()

dictYears = allStats.setStats(allYears)
dictGenres = allStats.setStats(allGenre)
dictDirectors = allStats.setStats(allDirectors)
dictRatings = allStats.setStats(allRating)
dictFran = allStats.setStats(allFran)
dictStudio = allStats.setStats(allStudio)

allStats.setGrossStats(allGross, allTitles)
allStats.setRunStats(allRun, allTitles)

print(f"Time Average: {allStats.getRunAvg()} min")
print(f"Shortest film: {allStats.getShortestFilm()} at {allStats.getMinRun()} min")
print(f"Longest film: {allStats.getLongestFilm()} at {allStats.getMaxRun()} min")

print(f"Lowest grossing: {allStats.getLowestFilm()} at ${allStats.getMinGross()}")
print(f"Highest grossing: {allStats.getHighestFilm()} at ${allStats.getMaxGross()}")
print(f"Gross Average: ${allStats.getGrossAvg()}")

print("Please check out the Excel for more stats!")

directorData = pd.DataFrame(
    {
        "Director" : [ key for key, value in dictDirectors.items() ],
        "Count" : [ value for key, value in dictDirectors.items() ],
    }
)
yearData = pd.DataFrame(
    {
        "Year" : [ key for key, value in dictYears.items() ],
        "Count" : [ value for key, value in dictYears.items() ],
    }
)
genreData = pd.DataFrame(
    {
        "Genre" : [ key for key, value in dictGenres.items() ],
        "Count" : [ value for key, value in dictGenres.items() ],
    }
)
ratingData = pd.DataFrame(
    {
        "Rating" : [ key for key, value in dictRatings.items() ],
        "Count" : [ value for key, value in dictRatings.items() ],
    }
)
franData = pd.DataFrame(
    {
        "Franchise" : [ key for key, value in dictFran.items() ],
        "Count" : [ value for key, value in dictFran.items() ],
    }
)
studioData = pd.DataFrame(
    {
        "Studio" : [ key for key, value in dictStudio.items() ],
        "Count" : [ value for key, value in dictStudio.items() ],
    }
)

book = load_workbook("Movie Data.xlsx")
directorSheet = book["Directors"]
yearSheet = book["Years"]
genreSheet = book["Genres"]
ratingSheet = book["Rating"]
franchiseSheet = book["Franchise"]
studioSheet = book["Studio"]

for index, row in directorData.iterrows():
    director = f"A{index + 1}"
    count = f"B{index + 1}"
    directorSheet[director] = row[0]
    directorSheet[count] = row[1]

for index, row in yearData.iterrows():
    year = f"A{index + 1}"
    count = f"B{index + 1}"
    yearSheet[year] = row[0]
    yearSheet[count] = row[1]

for index, row in genreData.iterrows():
    genre = f"A{index + 1}"
    count = f"B{index + 1}"
    genreSheet[genre] = row[0]
    genreSheet[count] = row[1]

for index, row in ratingData.iterrows():
    rating = f"A{index + 1}"
    count = f"B{index + 1}"
    ratingSheet[rating] = row[0]
    ratingSheet[count] = row[1]

for index, row in franData.iterrows():
    franchise = f"A{index + 1}"
    count = f"B{index + 1}"
    franchiseSheet[franchise] = row[0]
    franchiseSheet[count] = row[1]

for index, row in studioData.iterrows():
    studio = f"A{index + 1}"
    count = f"B{index + 1}"
    studioSheet[studio] = row[0]
    studioSheet[count] = row[1]

book.save("Movie Data.xlsx")
book.close()