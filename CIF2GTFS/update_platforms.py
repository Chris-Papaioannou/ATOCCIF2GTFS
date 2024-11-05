import os
import sys
import pandas as pd
import numpy as np
import shutil

sys.path.append(os.path.dirname(__file__))

import get_inputs as gi


def main():
    path = os.path.dirname(__file__)
    input_path = os.path.join(path, "input\\inputs.csv")

    updatePlatforms = gi.readPlatformUnknowns(input_path)
    if bool(updatePlatforms[0]):
        platformFilename = updatePlatforms[1]
        timetablePath = updatePlatforms[2]

        folderpath, filename = os.path.split(timetablePath)
        uneditedTimetablePath = os.path.join(folderpath, "Original_"+filename)

        shutil.copyfile(timetablePath, uneditedTimetablePath)

        dfPlatforms = pd.read_csv(platformFilename, dtype={'OrigDep':str,'CurrentArr':str,'CurrentDep':str})
        dfPlatforms['ID'] = list(zip(dfPlatforms.TrainUID, dfPlatforms.CurrentTIPLOC, dfPlatforms.CurrentArr, dfPlatforms.CurrentDep))
        platformsDict = dict(zip(dfPlatforms.ID, dfPlatforms.Platform))

        with open(uneditedTimetablePath) as cif: 
            with open(timetablePath, "w") as new_cif:

                for record in cif:

                    match record[:2]:
                        case 'BS':
                            TrainUID = record[3:9]
                            new_cif.write(record)
                        case 'LO':
                            OrigTIPLOC = record[2:10].strip()
                            OrigPlatNum = record[19:22].strip()
                            OrigDep = record[10:15].replace("H", "30").strip()

                            if OrigPlatNum == '':
                                if tuple((TrainUID, OrigTIPLOC, np.nan, OrigDep)) in platformsDict.keys():
                                    newPlatNum = platformsDict[tuple((TrainUID, OrigTIPLOC, np.nan, OrigDep))]
                                    new_record = record[:19]+str(newPlatNum).ljust(3, " ")+record[22:]
                                    new_cif.write(new_record)
                                else:
                                    new_cif.write(record)
                            else:
                                new_cif.write(record)

                        case 'LI':
                            InterTIPLOC = record[2:10].strip()
                            InterPlatNum = record[33:36].strip()
                            InterArr = record[10:15].replace("H","30").strip()
                            InterDep = record[15:20].replace("H","30").strip()
                            actCodes = [record[42:44].strip().upper(), record[44:46].strip().upper(), record[46:48].strip().upper(), record[48:50].strip().upper(), record[50:52].strip().upper(), record[52:54].strip().upper()]

                            stopAct = not set(actCodes).isdisjoint(['T', 'R', 'U', 'D'])

                            if InterPlatNum == '' and stopAct:
                                if tuple((TrainUID, InterTIPLOC, InterArr, InterDep)) in platformsDict.keys():
                                    newPlatNum = platformsDict[tuple((TrainUID, InterTIPLOC, InterArr, InterDep))]
                                    new_record = record[:33]+str(newPlatNum).ljust(3, " ")+record[36:]
                                    new_cif.write(new_record)
                                else:
                                    new_cif.write(record)
                            else:
                                new_cif.write(record)

                        case 'LT':
                            DestTIPLOC = record[2:10].strip()
                            DestPlatNum = record[19:22].strip()
                            DestArr = record[10:15].replace("H","30").strip()

                            if DestPlatNum == '':
                                if tuple((TrainUID, DestTIPLOC, DestArr,np.nan)) in platformsDict.keys():
                                    newPlatNum = platformsDict[tuple((TrainUID, DestTIPLOC, DestArr, np.nan))]
                                    new_record = record[:19]+str(newPlatNum).ljust(3, " ")+record[22:]
                                    new_cif.write(new_record)
                                else:
                                    new_cif.write(record)
                            else:
                                new_cif.write(record)
                            
                        case _:
                            new_cif.write(record)

if __name__ == "__main__":
    main()
