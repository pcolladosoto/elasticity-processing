# Ignore deprecation warnings emitted by openpyxl.
    # This bit needs to appear BEFORE importing openpyxl...
import warnings
warnings.simplefilter("ignore", DeprecationWarning)

import sys, argparse, pathlib, json, io, openpyxl

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from sklearn.linear_model import LinearRegression

from scipy import stats

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

    parser.add_argument("--output", default = 'processed_data.json', help = "JSON file to output processed data to.")

    parser.add_argument("--input", default = 'processed_data.json', help = "JSON file to read processed data from.")
    parser.add_argument("--summary", default = 'summarised_probes.json', help = "JSON file to read summarised data from.")

    parser.add_argument("--separator", default = ';', help = "Raw data field separator.")

    parser.add_argument("--excel-dir", default = './excels', help = "Directory on which to dump generated Excels.")

    parser.add_argument("--header-lines", type = int, default = 40, help = "Number of header lines.")

    parser.add_argument("--excel", action = "store_true", help = "Dump summarised Excel files per probe.")

    parser.add_argument("--merge-excels", action = "store_true", help = "Merge existing Excels on `--excel-dir` into a big one.")

    parser.add_argument("--summarised-excel", action = "store_true", help = "Generate a one-sheet summary Excel.")
    parser.add_argument("--output-excel-summary", default = "ProbeSummary.xlsx", help = "Excel file to dump probe data summary to.")

    parser.add_argument("--summarised-json", action = "store_true", help = "Generate a JSON file with the aggregated statistics.")
    parser.add_argument("--output-json-summary", default = "summarised_probes.json", help = "JSON file to dump probe data summary to.")

    parser.add_argument("--shapiro", action = "store_true", help = "Pass the Shapiro-Wilk test on the summarised data.")
    parser.add_argument("--shapiro-output", default = "ShapiroProbes.xlsx", help = "Excel file to dump the Shapiro-Wilk test results to.")

    parser.add_argument("--shore", action = "store_true", help = "Pass the Shapiro-Wilk test on the Shore hardness data.")

    parser.add_argument("--plots", action = "store_true", help = "Generate plots.")
    parser.add_argument("--plots-custom", action = "store_true", help = "Generate custom plots.")
    parser.add_argument("--plot-dir", default = "./plots", help = "Directory to store plots to.")

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
        ws["A6"] = "E / Score (1 is best):"
        ws["B6"] = processedData["youngModule"]["E"]
        ws["C6"] = processedData["youngModule"]["score"]

    excelBuff.seek(0, io.SEEK_SET)
    pathlib.Path(path).write_bytes(openpyxl.writer.excel.save_virtual_workbook(wb))

def parseRawData(file: pathlib.Path, separator: str, headerLines: int) -> dict:
    print(f"Parsing file: {file.name}...", file = sys.stderr)
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

    parsedData["maxTensionMPa"] = max(parsedData["data"]["tensionMPa"])
    parsedData["maxElongationN"] = max(parsedData["data"]["elongationN"])
    parsedData["ductility"] = (float(parsedData["finalLength"]["value"]) - 60) / 60

    parsedData["youngModule"] = findYoung(parsedData["data"]["elongationN"], parsedData["data"]["tensionMPa"])

    return parsedData

def findYoung(elongation: list[float], tension: list[float]) -> dict:
    # Number of samples to take into account for the linear regression
    fitPoints = int(len(elongation) / 2)

    # Converting the vectors into NumPy arrays...
    elongationX = np.array(elongation[:fitPoints]).reshape((-1, 1))
    tensionY = np.array(tension[:fitPoints])

    # Time to fit the data!

    # Returning the results. Bear in mind the `fit_intercept` kwarg implies
    # we take the intercept with the Y axis to be 0.
    model = LinearRegression(fit_intercept = False).fit(elongationX, tensionY)
    return {"E": model.coef_[0], "score": model.score(elongationX, tensionY)}

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

def summarisedJSON(summaryJSONName: str, processedData: dict):
    summary = {}

    for probeName, probeData in processedData.items():
        print(f"Summarising data for probe {probeName}...", file = sys.stderr)
        summary[probeName] = []

        for experimentName, experimentData in probeData.items():
            summary[probeName].append({
                "name": experimentName,
                "maxTensionMPa": experimentData["maxTensionMPa"],
                "maxElongationN": experimentData["maxElongationN"],
                "ductility": experimentData["ductility"],
                "finalLength": experimentData["finalLength"]["value"],
                "youngModuleE": experimentData["youngModule"]["E"],
                "youngModuleScore": experimentData["youngModule"]["score"]
            })

    pathlib.Path(summaryJSONName).write_text(json.dumps(summary, indent = 4))

def shapiroWilkTest(resultExcelName: str, summarisedData: dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Shapiro-Wilk Test Results"

    ws.cell(row = 1, column = 1, value = "Probe Name")
    ws.cell(row = 1, column = 2, value = "Sample Type")
    ws.cell(row = 1, column = 3, value = "Statistic (W)")
    ws.cell(row = 1, column = 4, value = "p-value")

    currRow = 2
    for probeName, probeData in summarisedData.items():
        print(f"Working on data for probe {probeName}...", file = sys.stderr)

        try:
            probeData = [experiment for experiment in probeData if experiment["name"].split('.')[0].split('-')[1] not in ['9', '10']]
        except IndexError:
            print("Skipping decimation of non-standard probe", file = sys.stderr)

        # Skip running the test on probes with less than 3 measurements: it's nonsense!
        if len(probeData) < 3:
            continue

        ws.cell(row = currRow, column = 1, value = probeName)

        shapiroMaxTensions = stats.shapiro(np.array([experiment["maxTensionMPa"] for experiment in probeData]))
        ws.cell(row = currRow, column = 2, value = "Tension [MPa]")
        ws.cell(row = currRow, column = 3, value = shapiroMaxTensions.statistic)
        ws.cell(row = currRow, column = 4, value = shapiroMaxTensions.pvalue)
        currRow += 1

        shapiroMaxElongations = stats.shapiro(np.array([experiment["maxElongationN"] for experiment in probeData]))
        ws.cell(row = currRow, column = 2, value = "Elongation [N]")
        ws.cell(row = currRow, column = 3, value = shapiroMaxElongations.statistic)
        ws.cell(row = currRow, column = 4, value = shapiroMaxElongations.pvalue)
        currRow += 1

        shapiroYoungs = stats.shapiro(np.array([experiment["youngModuleE"] for experiment in probeData]))
        ws.cell(row = currRow, column = 2, value = "E")
        ws.cell(row = currRow, column = 3, value = shapiroYoungs.statistic)
        ws.cell(row = currRow, column = 4, value = shapiroYoungs.pvalue)
        currRow += 1

    wb.save(resultExcelName)

def summarisedExcel(summaryExcelName: str, processedData: dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summarised Data"

    ws.cell(row = 1, column = 1, value = "Probe Name")
    ws.cell(row = 1, column = 2, value = "Experiment Name")
    ws.cell(row = 1, column = 3, value = "Max. Tension [MPa]")
    ws.cell(row = 1, column = 4, value = "Max. Elongation [N]")
    ws.cell(row = 1, column = 5, value = "Ductility [%A]")
    ws.cell(row = 1, column = 6, value = "Final Length [mm]")
    ws.cell(row = 1, column = 7, value = "E")
    ws.cell(row = 1, column = 8, value = "E Fit Score")

    currRow = 2
    for probeName, probeData in processedData.items():
        print(f"Summarising data for probe {probeName}...", file = sys.stderr)
        for experimentName, experimentData in probeData.items():
            ws.cell(row = currRow, column = 1, value = probeName)
            ws.cell(row = currRow, column = 2, value = experimentName)
            ws.cell(row = currRow, column = 3, value = experimentData["maxTensionMPa"])
            ws.cell(row = currRow, column = 4, value = experimentData["maxElongationN"])
            ws.cell(row = currRow, column = 5, value = experimentData["ductility"])
            ws.cell(row = currRow, column = 6, value = experimentData["finalLength"]["value"])
            ws.cell(row = currRow, column = 7, value = experimentData["youngModule"]["E"])
            ws.cell(row = currRow, column = 8, value = experimentData["youngModule"]["score"])
            currRow += 1

    wb.save(summaryExcelName)

def shoreShapiro():
    try:
        # Note we need to manually inspect the Excel file to find the indices!
        shoreHardnessDf = pd.read_excel("durezaShoreD.xlsx").iloc[2:67, 2:22]
    except FileNotFoundError:
        print(f"Couldn't load {shoreHardnessFile}...", file = sys.stderr)
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Shapiro-Wilk Test Results"

    ws.cell(row = 1, column = 1, value = "Row index")
    ws.cell(row = 1, column = 2, value = "Statistic (W)")
    ws.cell(row = 1, column = 3, value = "p-value")

    for i in range(len(shoreHardnessDf)):
        print(f"Working on data for row {i}...", file = sys.stderr)

        ws.cell(row = i + 2, column = 1, value = i + 1)

        shapiroRow = stats.shapiro(np.array(shoreHardnessDf.iloc[i]))
        ws.cell(row = i + 2, column = 2, value = shapiroRow.statistic)
        ws.cell(row = i + 2, column = 3, value = shapiroRow.pvalue)

    wb.save("durezaShoreShapiro.xlsx")

def genPlot(plotDir: str, expName: str, elongation: list[float], tension: list[float]):
    plt.figure(layout = "constrained")
    plt.xlabel("Elongation [N]")
    plt.ylabel("Tension [MPa]")
    plt.title(expName)
    plt.plot(elongation, tension, "g")
    plt.savefig(f"{plotDir}/{expName}.png", bbox_inches = "tight")
    plt.close()

def genAggregatedPlot(plotDir: str, probeName: str, probeData: dict):
    plt.figure(figsize = (25, 12.5), layout = "constrained")
    plt.xlabel("Elongation [N]")
    plt.ylabel("Tension [MPa]")
    plt.title(probeName)

    for experimentName, experimentData in probeData.items():
        plt.plot(experimentData["data"]["elongationN"],
            experimentData["data"]["tensionMPa"], label = experimentName.split(".")[0])

    plt.legend()
    plt.savefig(f"{plotDir}/agg{probeName}.png", bbox_inches = "tight")
    plt.close()

def genCustomPlot(plotDir: str, parsedFiles: dict):
    for name in ["Prob2Gris.raw", "Prob10Tenaflex.raw", "Prob7.raw", "Prob1AmarilloCanario.raw"]:
        print(f"Probe {name}:")
        for experimentName, experimentData in parsedFiles[name].items():
            print(f"\t{len(experimentData['data']['elongationN'])}\t{len(experimentData['data']['tensionMPa'])}")

    # return

    plt.figure(figsize = (25, 12.5), layout = "constrained")
    plt.xlabel("Elongation [log(N)]")
    plt.ylabel("Tension [MPa]")
    plt.title("Prob2Gris, Prob10Tenaflex, Prob7 y Prob1AmarilloCanario")
    plt.xscale("log")

    for name in ["Prob2Gris.raw", "Prob10Tenaflex.raw", "Prob7.raw", "Prob1AmarilloCanario.raw"]:
        for experimentName, experimentData in parsedFiles[name].items():
            plt.plot(experimentData["data"]["elongationN"],
                experimentData["data"]["tensionMPa"], label = name)

    plt.legend()
    plt.savefig(f"{plotDir}/customGraph.png", bbox_inches = "tight")
    plt.close()

def main():
    args = parseArgs()

    if args.summarised_excel:
        try:
            parsedFiles = json.loads(pathlib.Path(args.input).read_text())
        except FileNotFoundError:
            print(f"Couldn't load {args.input}. Have you processed the data?", file = sys.stderr)
            return -1
        summarisedExcel(args.output_excel_summary, parsedFiles)

    if args.summarised_json:
        try:
            parsedFiles = json.loads(pathlib.Path(args.input).read_text())
        except FileNotFoundError:
            print(f"Couldn't load {args.input}. Have you processed the data?", file = sys.stderr)
            return -1
        summarisedJSON(args.output_json_summary, parsedFiles)

    if args.shapiro:
        try:
            summarisedData = json.loads(pathlib.Path(args.summary).read_text())
        except FileNotFoundError:
            print(f"Couldn't load {args.summary}. Have you summarised the data? You can use:\n" +
                  "\tpython3 process_data.py --summarised-json foo", file = sys.stderr)
            return -1
        shapiroWilkTest(args.shapiro_output, summarisedData)

    if args.shore:
        shoreShapiro()

    if args.merge_excels:
        joinExcels(args.excel_dir)
        return 0

    if args.excel:
        try:
            parsedFiles = json.loads(pathlib.Path(args.input).read_text())
        except FileNotFoundError:
            print(f"Couldn't load {args.input}. Have you processed the data?", file = sys.stderr)
            return -1

        for probe, data in parsedFiles.items():
            print(f"Dumping Excel for probe {probe}", file = sys.stderr)
            dumpExcel(f"./{args.excel_dir}/{probe}.xlsx", data)
        return 0

    if args.plots:
        try:
            parsedFiles = json.loads(pathlib.Path(args.input).read_text())
        except FileNotFoundError:
            print(f"Couldn't load {args.input}. Have you processed the data?", file = sys.stderr)
            return -1

        for probeName, probeData in parsedFiles.items():
            print(f"Generating plots for probe {probeName}...", file = sys.stderr)
            genAggregatedPlot(args.plot_dir, probeName, probeData)
            for experimentName, experimentData in probeData.items():
                genPlot(args.plot_dir, experimentName.split(".")[0], experimentData["data"]["elongationN"], experimentData["data"]["tensionMPa"])
        return 0

    if args.plots_custom:
        try:
            parsedFiles = json.loads(pathlib.Path(args.input).read_text())
        except FileNotFoundError:
            print(f"Couldn't load {args.input}. Have you processed the data?", file = sys.stderr)
            return -1

        # print(json.dumps(list(parsedFiles.keys()), indent = 2))

        genCustomPlot(args.plot_dir, parsedFiles)

    if pathlib.Path(args.path).is_dir():
        parsedFiles = {}
        for file in pathlib.Path(args.path).iterdir():
            if file.name.split('.')[1] != "raw":
                continue

            fileRoot = file.name.split('-')[0]
            if not parsedFiles.get(fileRoot, False):
                parsedFiles[fileRoot] = {}

            parsedFiles[fileRoot][file.name] = parseRawData(file, args.separator, args.header_lines)

        print("Dumping parsed data to a JSON file...", file = sys.stderr)
        pathlib.Path(args.output).write_text(json.dumps(parsedFiles, indent = 2))

        return 0

if __name__ == "__main__":
    main()
