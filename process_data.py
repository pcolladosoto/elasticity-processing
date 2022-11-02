# Ignore deprecation warnings emitted by openpyxl.
    # This bit needs to appear BEFORE importing openpyxl...
import warnings
warnings.simplefilter("ignore", DeprecationWarning)

import sys, argparse, pathlib, json, io, openpyxl

import pandas as pd

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
    parser.add_argument("raw_data", help = "File with a probe's raw data.")
    parser.add_argument("--separator", default = ';', help = "Raw data field separator.")
    parser.add_argument("--header-lines", type = int, default = 40, help = "Number of header lines.")
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
        parsedData[fieldMap[removeQuotes(line[0])]] = {"value": removeQuotes(line[1]), "unit": "none"}
    elif len(line) == 3:
        parsedData[fieldMap[removeQuotes(line[0])]] = {"value": float(removeQuotes(line[1])), "unit": removeQuotes(line[2])}
    else:
        print(f"expected 1, 2 or 3 fields and got {len(line)}: {line}. Quitting...")
        sys.exit(-1)

def dumpExcel(path: str, parsedData):
    chosenColNames = ["tensionMPa", "elongationN", "extensionMM", "loadN"]
    df = pd.DataFrame(
        [[parsedData["data"][colName][i] for colName in chosenColNames] for i in range(len(parsedData["data"][chosenColNames[0]]))],
        columns = ["Tensión [MPa]", "Elongación [N]", "Extensión [mm]", "Carga [N]"]
    )
    excelBuff = io.BytesIO()
    df.to_excel(excelBuff, sheet_name = "Datos", index = False, startrow = 5, startcol = 0)
    excelBuff.seek(0, io.SEEK_SET)

    wb = openpyxl.load_workbook(excelBuff)
    ws = wb.active

    ws["A1"] = "Máximos:"
    ws["A2"] = "Tensión Máxima [MPa]:"
    ws["B2"] = parsedData["data"]["maxTensionMPa"]
    ws["A3"] = "Elongación Máxima [N]:"
    ws["B3"] = parsedData["data"]["maxElongationN"]
    ws["A4"] = "Ductilidad [%]:"
    ws["B4"] = parsedData["data"]["ductility"]

    excelBuff.seek(0, io.SEEK_SET)
    pathlib.Path(path).write_bytes(openpyxl.writer.excel.save_virtual_workbook(wb))

def main():
    args = parseArgs()
    parsedData = {"data": {"tensionMPa": [], "elongationN": []}}

    for i, line in enumerate(pathlib.Path(args.raw_data).read_text(encoding = "utf-8", errors = "replace").splitlines()):
        sLine = line.split(args.separator)

        if i < args.header_lines:
            parseHeaderLine(parsedData, sLine)

        elif i == args.header_lines:
            colNames = [fieldMap[removeQuotes(field)] for field in sLine]
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

    dumpExcel(f"{args.raw_data.split('.')[0].split('/')[-1]}.xlsx", parsedData)
    pathlib.Path(f"{args.raw_data.split('.')[0].split('/')[-1]}.json").write_text(json.dumps(parsedData, indent = 2, ensure_ascii = False))

if __name__ == "__main__":
    main()
