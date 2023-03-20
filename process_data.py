# Ignore deprecation warnings emitted by openpyxl.
    # This bit needs to appear BEFORE importing openpyxl...
import warnings
warnings.simplefilter("ignore", DeprecationWarning)

import sys, argparse, pathlib, json, io, openpyxl

import pandas as pd

import numpy as np

from sklearn.linear_model import LinearRegression

fieldMap = {
    "Tipo de ensayo": "experimentType",
    "Nombre del m\ufffdtodo": "methodName",
    "Nombre": "name",
    "ID operador": "operatorID",
    "Empresa": "enterprise",
    "Nombre lab.": "labName",
    "Fecha ensayo": "experimentDate",
    "Temperatura": "temperature",
    "Humedad": "humidity",
    "Nota 1": "noteA",
    "Nota 2": "noteB",
    "Nota 3": "noteC",
    "Geometr\ufffda": "geometry",
    "Probeta": "probe",
    "Nombre probeta": "probeName",
    "Anchura": "width",
    "Espesor": "thickness",
    "Longitud": "length",
    "Di\ufffdmetro": "diameter",
    "Di\ufffdmetro int": "innerDiameter",
    "Di\ufffdmetro ext": "exteriorDiameter",
    "Espesor pared": "wallThickness",
    "\ufffdrea": "area",
    "Densidad lineal": "linearDensity",
    "Peso de pat\ufffdn": "railWeight",
    "Separa. rodillos de carga": "loadRollSeparation",
    "Separa. rodillos de soporte": "supportRollSeparation",
    "Separaci\ufffdn rodillos": "rollSeparation",
    "Tipo fijaci\ufffdn": "fasteningType",
    "Observaciones": "observations",
    "Anchura final": "finalWidth",
    "Espesor final": "finalThickness",
    "Longitud final": "finalLength",
    "Di\ufffdmetro final": "finalDiameter",
    "Di\ufffdmetro interior final": "finalInnerDiameter",
    "Di\ufffdmetro exterior final": "finalExteriorDiameter",
    "Espesor de pared final": "finalWallThickness",
    "\ufffdrea final": "finalArea",
    "Densidad lineal final": "finalLinearDensity",
    "Tiempo sec": "secTime",
    "Extensi\ufffdn mm": "extensionMM",
    "Carga N": "loadN",
    "Resistencia MPa": "resistanceMPa",
    "N\ufffdmero ciclos ": "nCycles",
    "N\ufffdmero total de ciclos ": "totalnCycles",
    "Total de repeticiones ": "repetitionTotal",
    "Deform. [Exten.] %": "deformationPercent",
    "Tenacidad gf/tex": "tenacity"
}

def parseArgs():
    parser = argparse.ArgumentParser(description = "Elasticity data analyser.")
    parser.add_argument("path", help = "File or directory containing raw data.")
    parser.add_argument("--separator", default = ';', help = "Raw data field separator.")
    parser.add_argument("--header-lines", type = int, default = 40, help = "Number of header lines.")
    parser.add_argument("--json", action = "store_true", help = "Dump processed data into a global JSON file.")
    parser.add_argument("--excel", action = "store_true", help = "Dump merged Excel files per probe.")
    return parser.parse_args()

def removeQuotes(raw: str) -> str:
    if raw[-1] == '"':
        raw = raw[:-1]
    if raw[-1] == ':':
        raw = raw[:-1]
    if raw[0] == '"':
        raw = raw[1:]
    return raw

def parseHeaderLine(parsedData: dict[str, dict[str]], line: list[str]):
    if len(line) == 1:
        # We've got some malformed lines coming from the machine...
        pass
    elif len(line) == 2:
        parsedData[fieldMap.get(removeQuotes(line[0]), "errField")] = {"value": removeQuotes(line[1]), "unit": "none"}
    elif len(line) == 3:
        parsedData[fieldMap.get(removeQuotes(line[0]), "errField")] = {"value": float(removeQuotes(line[1])), "unit": removeQuotes(line[2])}
    else:
        print(f"expected 1, 2 or 3 fields and got {len(line)}: {line}. Quitting...")
        sys.exit(-1)

def dumpExcel(path: str, parsedData):
    chosenColNames = ["tensionMPa", "elongationN", "extensionMM", "loadN"]

    excelBuff = io.BytesIO()
    excelWriter = pd.ExcelWriter(excelBuff)

    with pd.ExcelWriter(excelBuff) as writer:
        for fName, processedData in parsedData:
            df = pd.DataFrame(
                [[processedData["data"][colName][i] for colName in chosenColNames] for i in range(len(processedData["data"][chosenColNames[0]]))],
                columns = ["Tensión [MPa]", "Elongación [N]", "Extensión [mm]", "Carga [N]"]
            )
            df.to_excel(writer, sheet_name = fName, index = False, startrow = 8, startcol = 0)

    excelBuff.seek(0, io.SEEK_SET)

    wb = openpyxl.load_workbook(excelBuff)
    for fName, processedData in parsedData:
        ws = wb[fName]

        ws["A1"] = "Máximos:"
        ws["A2"] = "Tensión Máxima [MPa]:"
        ws["B2"] = processedData["data"]["maxTensionMPa"]
        ws["A3"] = "Elongación Máxima [N]:"
        ws["B3"] = processedData["data"]["maxElongationN"]
        ws["A4"] = "Ductilidad [%]:"
        ws["B4"] = processedData["data"]["ductility"]
        ws["A5"] = "Longitud final [mm]:"
        ws["B5"] = processedData["finalLength"]["value"]

    excelBuff.seek(0, io.SEEK_SET)
    pathlib.Path(path).write_bytes(openpyxl.writer.excel.save_virtual_workbook(wb))

def parseRawData(file: pathlib.Path, separator: str, headerLines: int) -> dict:
    print(f"Parsing file: {file.name}", file = sys.stderr)
    parsedData = {"data": {"tensionMPa": [], "elongationN": []}}
    for i, line in enumerate(file.read_text(encoding = "utf-8", errors = "replace").splitlines()):
        # print(f"Parsing line: {line}", file = sys.stderr)

        sLine = line.split(separator)

        if i < headerLines:
            parseHeaderLine(parsedData, sLine)

        elif i == headerLines:
            colNames = [fieldMap.get(removeQuotes(field), "extensionMM") for field in sLine]
            for colName in colNames:
                parsedData["data"][colName] = []

        else:
            for measure, value in zip(colNames, sLine):
                parsedData["data"][measure].append(float(value.replace(',', '.')))
            parsedData["data"]["tensionMPa"].append(parsedData["data"]["loadN"][-1] / 40)
            parsedData["data"]["elongationN"].append(parsedData["data"]["extensionMM"][-1] / 60)

    parsedData["data"]["maxTensionMPa"] = max(parsedData["data"]["tensionMPa"])
    parsedData["data"]["maxElongationN"] = max(parsedData["data"]["elongationN"])
    parsedData["data"]["ductility"] = (float(parsedData["finalLength"]["value"]) - 60) / 60

    return parsedData

def findYoung(probe, fullData):
    fitData = []
    for experimentName, data in fullData:
        # print(f"Processing experiment {experiment}", file = sys.stderr)
        fitPoints = int(len(data["data"]["elongationN"]) / 2)
        elongationX = np.array(data["data"]["elongationN"][:fitPoints]).reshape((-1, 1))
        tensionY = np.array(data["data"]["tensionMPa"][:fitPoints])
        model = LinearRegression(fit_intercept = False).fit(elongationX, tensionY)
        # print(f"Fit score (best is 1): {model.score(elongationX, tensionY)}", file = sys.stderr)
        # print(f"Intercept = {model.intercept_}; slope = {model.coef_}", file = sys.stderr)
        fitData.append((experimentName, model.coef_[0], model.score(elongationX, tensionY)))
    return fitData

def addYoungInfoToExcel(path: str, fitData):
    wb = openpyxl.load_workbook(filename = path)
    for expName, slope, score in fitData:
        ws = wb[expName]
        ws["A6"] = "E / Score (1 is best):"
        ws["B6"] = slope
        ws["C6"] = score

    wb.save(path)

def joinExcels(path: str):
    joinedExcels = pd.ExcelWriter("joinedProbes.xlsx")

    for file in pathlib.Path(path).iterdir():
        print(f"Mergeing Excel file {file.absolute()}...")
        currExcel = pd.ExcelFile(file.absolute())

        # Returns all the sheets as dataframes!
        dfs = pd.read_excel(currExcel, None)

        for sheetName, sheetDf in dfs.items():
            sheetDf.to_excel(joinedExcels, sheet_name = sheetName, startrow = 0, startcol = 0)

    joinedExcels.close()

def main():
    args = parseArgs()

    if args.path == "young":
        try:
            parsedFiles = json.loads(pathlib.Path("processed_data.json").read_text())
        except FileNotFoundError:
            print("couldn't find ./processed_data.json: remember to process the data beforehand")
            return -1

        for probe, fullData in parsedFiles.items():
            print(f"Processing probe {probe}", file = sys.stderr)

            # More info over at https://realpython.com/linear-regression-in-python/
            for experiment, data in fullData:
                print(f"Processing experiment {experiment}", file = sys.stderr)
                fitPoints = int(len(data["data"]["elongationN"]) / 2)
                elongationX = np.array(data["data"]["elongationN"][:fitPoints]).reshape((-1, 1))
                tensionY = np.array(data["data"]["tensionMPa"][:fitPoints])
                model = LinearRegression(fit_intercept = False).fit(elongationX, tensionY)
                print(f"Fit score (best is 1): {model.score(elongationX, tensionY)}", file = sys.stderr)
                print(f"Intercept = {model.intercept_}; slope = {model.coef_}", file = sys.stderr)

                sys.exit(-1)

    if args.path == "addYoung":
        parsedFiles = json.loads(pathlib.Path("processed_data.json").read_text())

        for probe, fullData in parsedFiles.items():
            print(f"Adding young data for probe {probe}...")
            addYoungInfoToExcel(f"./excels/{probe}.xlsx", findYoung(probe, fullData))

        return 0

    if args.path == "mergeExcels":
        joinExcels("./excels")

    if args.excel:
        parsedFiles = json.loads(pathlib.Path("processed_data.json").read_text())
        for probe, data in parsedFiles.items():
            print(f"Merging data for probe {probe}", file = sys.stderr)
            dumpExcel(f"./excels/{probe}.xlsx", data)
        return 0

    if pathlib.Path(args.path).is_dir():
        parsedFiles = {}
        for file in pathlib.Path(args.path).iterdir():
            if file.name.split('.')[1] != "raw":
                continue

            fileRoot = file.name.split('-')[0]
            if not parsedFiles.get(fileRoot, False):
                parsedFiles[fileRoot] = []

            parsedFiles[fileRoot].append([file.name, parseRawData(file, args.separator, args.header_lines)])

        if args.json:
            print("Dumping parsed data to a JSON file...", file = sys.stderr)
            pathlib.Path("processed_data.json").write_text(json.dumps(parsedFiles, indent = 2))

        if args.excel:
            for probe, data in parsedFiles.items():
                print(f"Merging data for probe {probe}", file = sys.stderr)
                dumpExcel(f"./excels/{probe}.xlsx", data)
        return 0

if __name__ == "__main__":
    main()
