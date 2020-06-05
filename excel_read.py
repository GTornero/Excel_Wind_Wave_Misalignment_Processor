import xlrd
import xlsxwriter
import numpy as np
import tkinter
from tkinter.filedialog import askopenfilename


def my_round_down(value, sig=1):
    return (value // sig) * sig


tkinter.Tk().withdraw()
path = askopenfilename()

windSpeeds = []
windDirs = []
windSectors = []
waveHeights = []
wavePPeriods = []
waveDirs = []
waveSectors = []

windRose = [0]*12
waveRose = [0]*12

print("What step in [m] do you want for the Significant Wave Height? (default value 0.25m)")
waveStep = float(input())

if waveStep == '':
    waveStep = 0.25

print("What is the Maximum Significant Wave height in [m]? (default 10m)")
maxWave = float(input())

if maxWave == '':
    maxWave = 10

print("What step in [m/s] do you want for the Wind Speed (default value 1m/s")
windStep = float(input())

if windStep == '':
    windStep = 1

print("What is the Maximum Wind Speed in [m/s]? (default 40m/s)")
maxWind = float(input())

if maxWind == '':
    maxWind = 40

windWaveTables = np.zeros((12, 12, int(maxWind / windStep) + 1, int(maxWave / waveStep) + 1))

with open(path, 'r') as f:

    inputWorkbook = xlrd.open_workbook(path)
    inputWorksheet = inputWorkbook.sheet_by_index(0)

    outWorkbook = xlsxwriter.Workbook("output.xlsx")
    outSheet = outWorkbook.add_worksheet("Wind Data")

    for y in range(inputWorksheet.nrows):
        windSpeeds.append(inputWorksheet.cell_value(y, 2))
        windDirs.append(round(inputWorksheet.cell_value(y, 3), 1))
        waveHeights.append(inputWorksheet.cell_value(y, 4))
        waveDirs.append(inputWorksheet.cell_value(y, 5))
        wavePPeriods.append(inputWorksheet.cell_value(y, 6))

        if round(inputWorksheet.cell_value(y, 3), 1) >= 345:
            windSector = 1
        else:
            windSector = int(((round(inputWorksheet.cell_value(y, 3), 1) + 15) / 30) + 1)

        if round(inputWorksheet.cell_value(y, 5), 1) >= 345:
            waveSector = 1
        else:
            waveSector = int(((round(inputWorksheet.cell_value(y, 5), 1) + 15) / 30) + 1)

        windSectors.append(windSector)
        waveSectors.append(waveSector)

        windRose[windSector - 1] += 1
        waveRose[waveSector - 1] += 1

        outSheet.write("A1", "Wind Speed [m/s]")
        outSheet.write("B1", "Wind Direction [deg]")
        outSheet.write("C1", "Wind Sector")
        outSheet.write("D1", "Significant Wave Height [m]")
        outSheet.write("E1", "Wave Direction [deg]")
        outSheet.write("F1", "Peak Period [s]")
        outSheet.write("G1", "Wave Sector")

        outSheet.write("A" + str(y + 2), windSpeeds[y])
        outSheet.write("B" + str(y + 2), windDirs[y])
        outSheet.write("C" + str(y + 2), windSectors[y])
        outSheet.write("D" + str(y + 2), waveHeights[y])
        outSheet.write("E" + str(y + 2), waveDirs[y])
        outSheet.write("F" + str(y + 2), wavePPeriods[y])
        outSheet.write("G" + str(y + 2), waveSectors[y])

        windWaveTables[windSector - 1, waveSector - 1, int(windSpeeds[y] // windStep), int(waveHeights[y] // waveStep)] += 1

    for y in range(len(windRose)):
        outSheet.write("J" + str(y + 2), windRose[y])
        outSheet.write("K" + str(y + 2), waveRose[y])

    tableCols = np.arange(0, maxWave + waveStep, waveStep)
    tableRows = np.arange(0, maxWind + windStep, windStep)

    test = windWaveTables.sum(axis=(0, 1))

    outSheet = outWorkbook.add_worksheet("Total")
    outSheet.write(2, 2, "Significant Weight Height Bin [m]")
    outSheet.write(4, 0, "Wind Speed Bin [m/s]")

    i = 0
    while i < len(tableCols):
        outSheet.write(3, 2 + i, tableCols[i])
        i += 1

    i = 0
    while i < len(tableRows):
        outSheet.write(4 + i, 1, tableRows[i])
        i += 1

    for y in range(test.shape[0]):
        for x in range(test.shape[1]):
            outSheet.write(4 + y, 2 + x, test[y][x])


    for windSector in range(1, 13):
        outSheet = outWorkbook.add_worksheet("Wind Sector " + str(windSector))
        for waveSector in range(1, 13):
            outSheet.write(2, 2 + ((waveSector - 1) * (len(tableCols) + 2)), "Wave Sector" + str(waveSector))
            i = 0
            while i < len(tableCols):
                outSheet.write(3, 2 + i + ((waveSector - 1) * (len(tableCols) + 2)), tableCols[i])
                i += 1

            i = 0
            while i < len(tableRows):
                outSheet.write(4 + i, 1 + ((waveSector - 1) * (len(tableCols) + 2)), tableRows[i])
                i += 1

            for x in range(int(maxWind / windStep) + 1):
                for y in range(int(maxWave / waveStep) + 1):
                    outSheet.write(4 + x, 2 + y + (((waveSector - 1) * (len(tableCols) + 2))), windWaveTables[windSector - 1][waveSector - 1][x][y])

    outWorkbook.close()
