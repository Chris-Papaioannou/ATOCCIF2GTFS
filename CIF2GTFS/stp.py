import os
import sys
import pandas as pd
import logging
import traceback
import numpy as np

sys.path.append(os.path.dirname(__file__))

import get_inputs as gi

logging.basicConfig(
    filename="ModelBuilder.log",
    encoding="utf-8",
    filemode="a",
    format="{asctime} - {levelname} - {message}",
    style="{",
    datefmt="%Y-%m-%d %H:%M",
    level=logging.INFO # Change to logging.DEBUG for more details
)



def getSTP(cifPath, folder):
    
    permanent = []
    shortTerm = []
    overlay = []
    cancellation = []


    with open(cifPath) as cif:

        for record in cif:

            match record[:2]:
                case 'HD':
                    calendarFrom = record[48:54]
                    calendarTo = record[54:60]

                case 'BS':

                    if record[2] != 'N':
                        logging.warning(f"This Basic Schedule isn't a new record - TrainUID: {TrainUID}.")

                    TrainUID = record[3:9]
                    STP = record[79]
                    fromDate = pd.to_datetime(record[9:15], format='%y%m%d')
                    toDate = pd.to_datetime(record[15:21], format='%y%m%d')
                    dayStr = record[21:28]
                    serviceID = record[2]+TrainUID+record[9:15]+record[15:21]+STP

                    match STP:
                        case 'P':
                            permanent.append([serviceID, TrainUID, fromDate, toDate, dayStr])
                        case 'S':
                            shortTerm.append([serviceID, TrainUID, fromDate, toDate, dayStr])
                        case 'O':
                            overlay.append([serviceID, TrainUID, fromDate, toDate, dayStr])
                        case 'C':
                            cancellation.append([serviceID, TrainUID, fromDate, toDate, dayStr])
                                                
                case _:
                    logging.debug("Record ignored - "+record.rstrip('\n'))

   
    dfPermanent = pd.DataFrame(permanent, columns=['ServiceID', 'TrainUID', 'FromDate', 'ToDate', 'dayStr'])
    dfShortTerm = pd.DataFrame(shortTerm, columns=['ServiceID', 'TrainUID', 'FromDate', 'ToDate', 'dayStr'])
    dfOverlay = pd.DataFrame(overlay, columns=['ServiceID', 'TrainUID', 'FromDate', 'ToDate', 'dayStr'])
    dfCancellation = pd.DataFrame(cancellation, columns=['ServiceID', 'TrainUID', 'FromDate', 'ToDate', 'dayStr'])

    return dfPermanent, dfShortTerm, dfOverlay, dfCancellation



def main(cifPath):

    path = os.path.dirname(__file__)
    input_path = os.path.join(path, "input\\inputs.csv")

    importTimetable = gi.readTimetableInputs(input_path)

    cifPath = importTimetable[1]

    if bool(importTimetable[0]):
        try:
            logging.info("Processing STP services")

            folder, cifName = os.path.split(cifPath)
            
            logging.info(f"Timetable Path: {cifPath}")

            dfPermanent, dfShortTerm, dfOverlay, dfCancellation = getSTP(cifPath, folder)

            calendarFrom = Visum.Net.CalendarPeriod.AttValue('ValidFrom')
            calendarTo = Visum.Net.CalendarPeriod.AttValue('ValidUntil')

            updatedDayVectors = {}

            for i, (pServiceID, pUID, pFromDate, pToDate, pDayStr) in dfPermanent.iterrows():

                pvalidDays = [x for x in pd.date_range(pFromDate, pToDate) if pDayStr[x.weekday()]=='1']
                
                if pUID in dfCancellation.TrainUID.to_list():
                    for j, (cServiceID, cUID, cFromDate, cToDate, cDayStr) in dfCancellation.loc[dfCancellation.TrainUID==pUID].iterrows():
                        cvalidDays = [x for x in pd.date_range(cFromDate, cToDate) if cDayStr[x.weekday()]=='1']
                        pvalidDays = [x for x in pvalidDays if x not in cvalidDays]

                    newDayVector = ''.join(["1" if x in pvalidDays else "0" for x in pd.date_range(calendarFrom, calendarTo)])
                    updatedDayVectors[pServiceID+"_service"] = newDayVector
            
                if pUID in dfOverlay.TrainUID.to_list():
                    for k, (oServiceID, oUID, oFromDate, oToDate, oDayStr) in dfOverlay.loc[dfOverlay.TrainUID==pUID].iterrows():
                        ovalidDays = [x for x in pd.date_range(oFromDate, oToDate) if oDayStr[x.weekday()]=='1']
                        pvalidDays = [x for x in pvalidDays if x not in ovalidDays]

                    newDayVector = ''.join(["1" if x in pvalidDays else "0" for x in pd.date_range(calendarFrom, calendarTo)])
                    updatedDayVectors[pServiceID+"_service"] = newDayVector
            
            logging.info("Updating services")


            validDays = Visum.Net.ValidDaysCont.GetFilteredSet("[No]>1")
            validDaysDf = pd.DataFrame(validDays.GetMultipleAttributes(["No", 'Code', "DayVector"]), columns=['No', 'Code', 'DayVector'])
            validDaysDf['NewDayVector'] = validDaysDf.Code.replace(updatedDayVectors)
            validDaysDf['NewDayVector'] = np.where(validDaysDf.NewDayVector == validDaysDf.Code, validDaysDf.DayVector, validDaysDf.NewDayVector)

            newValidDays = list(zip(validDaysDf.No, validDaysDf.NewDayVector))

            validDays.SetMultipleAttributes(["No", "DayVector"], newValidDays)


            logging.info("Done")
        
        except:
            logging.error(traceback.format_exc())


if __name__ == "__main__":
        main()   
