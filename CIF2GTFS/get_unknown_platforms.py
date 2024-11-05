import wx
import os
import pandas as pd
import logging

logging.basicConfig(
    filename="UnknownPlatforms.log",
    encoding="utf-8",
    filemode="w",
    format="{asctime} - {levelname} - {message}",
    style="{",
    datefmt="%Y-%m-%d %H:%M",
    level=logging.INFO # Change to logging.DEBUG for more details
)



def file_select_dlg(message, wildcard):
    with wx.FileDialog(parent=None, message=message, wildcard=wildcard,
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:

        if dlg.ShowModal() == wx.ID_CANCEL:
            exit(0)
        pathname = dlg.GetPath()
    return pathname


def main():

    app = wx.App()

    wildcard = "All files (*.*)|*.*"

    cifPath = file_select_dlg("Please select the CIF file...", wildcard)
    logging.info(f"CIF File - {cifPath}")

    folder, cifName = os.path.split(cifPath)
    cifName = os.path.splitext(cifName)[0]

    outputPath = os.path.join(folder, f"PlatformUnknowns_{cifName}.csv")

    # List of lists with attributes [TrainUID, OrigTIPLOC, OrigDep, CurrentTIPLOC, CurrentArr, CurrentDep]
    unknownPlatformStops = []

    with open(cifPath) as cif:

        for record in cif:

            match record[:2]:
                case 'BS':
                    TrainUID = record[3:9]
                    if record[2] != 'N':
                        logging.warning(f"This Basic Schedule isn't a new record - TrainUID: {TrainUID}.")
                case 'LO':
                    OrigTIPLOC = record[2:10].strip()
                    OrigPlatNum = record[19:22].strip()
                    OrigDep = record[10:15].replace("H", "30").strip()

                    if OrigPlatNum == '':
                        unknownPlatformStops.append([TrainUID,  OrigTIPLOC, OrigDep, OrigTIPLOC, "", OrigDep])

                case 'LI':
                    InterTIPLOC = record[2:10].strip()
                    InterPlatNum = record[33:36].strip()
                    InterArr = record[10:15].replace("H","30").strip()
                    InterDep = record[15:20].replace("H","30").strip()
                    actCodes = [record[42:44].strip().upper(), record[44:46].strip().upper(), record[46:48].strip().upper(), record[48:50].strip().upper(), record[50:52].strip().upper(), record[52:54].strip().upper()]

                    stopAct = not set(actCodes).isdisjoint(['T', 'R', 'U', 'D'])

                    if InterPlatNum == '' and stopAct:
                        unknownPlatformStops.append([TrainUID,  OrigTIPLOC, OrigDep, InterTIPLOC, InterArr, InterDep])

                case 'LT':
                    DestTIPLOC = record[2:10].strip()
                    DestPlatNum = record[19:22].strip()
                    DestArr = record[10:15].replace("H","30").strip()

                    if DestPlatNum == '':
                        unknownPlatformStops.append([TrainUID,  OrigTIPLOC, OrigDep, DestTIPLOC, DestArr, ""])
                    
                case _:
                    logging.debug("Record ignored - "+record.rstrip('\n'))
    
    dfPlatUnknowns = pd.DataFrame(unknownPlatformStops, columns=['TrainUID', 'OrigTIPLOC', 'OrigDep', 'CurrentTIPLOC', 'CurrentArr', 'CurrentDep'])
    dfPlatUnknowns['Platform'] = ""

    dfPlatUnknowns.to_csv(outputPath, index=False)

    wx.MessageBox("Unknowns platform list created.", "Processing Complete", wx.OK | wx.ICON_INFORMATION)



if __name__ == '__main__':
    main()
        