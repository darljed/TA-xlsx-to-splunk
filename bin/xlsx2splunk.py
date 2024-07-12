#! $SPLUNK_HOME/bin/splunk cmd python
import json, os, openpyxl, base64, re, sys
from glob import glob

checkpoint_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)),"cp")
if(not os.path.exists(checkpoint_dir)):
    os.makedirs(checkpoint_dir)

def formatToDict(header, dataset):
    if(header and dataset):
        d = []
        for data in dataset:
            td = {}
            for i, field in enumerate(header):
                td[field] = data[i]
            d.append(td)
        return d
    return None

def toBase64(str):
    b = str.encode("ascii") 
    byte = base64.b64encode(b) 
    bstring = byte.decode("ascii") 
    clean = re.sub(r"[\/\\\=\+]", "", bstring)
    return [bstring,clean]

def createCheckpoint(file,state):
    dir, actFile = os.path.split(file)
    checkpointStr, checkpointFile = toBase64(actFile)
    checkpointFile = os.path.join(checkpoint_dir,checkpointFile)
    with open(checkpointFile,"w") as f:
        f.write(str(state))
        f.write(f"\n{file}")


if __name__ == "__main__":
    try:
        path = sys.argv[1]
        
        file_name = sys.argv[2]
        header_row = int(sys.argv[3])
        filepath = os.path.join(path,file_name)
        foundFiles = glob(filepath)
        fileList = []
        for ff in foundFiles:
            if ff.endswith('.xlsx'):
                fileList.append(ff)
        try:
            for file in fileList:
                if(os.path.exists(file)):
                    dir, actFile = os.path.split(file)
                    checkpointStr, checkpointFile = toBase64(actFile)

                    # check if checkpoint file exists 
                    oldState = None
                    checkpointFile = os.path.join(checkpoint_dir,checkpointFile)

                    if os.path.exists(checkpointFile):
                        with open(checkpointFile,'r') as f:
                            oldState = float(f.readline().replace('\n',""))
                    
                    state = os.path.getmtime(file)
                    proceedIndexFlag = False
                    
                    if(oldState):
                        if state > oldState:
                            proceedIndexFlag = True
                    
                    else:
                        proceedIndexFlag = True
                    
                    if(proceedIndexFlag):
                        wb = openpyxl.load_workbook(filename=file)
                        
                        for sheetname in wb.sheetnames:
                            sheet = wb[sheetname]
                            
                            
                            data = []
                            header = []
                            for row in sheet.iter_rows(values_only=True,min_row=header_row):
                                if(len(header) == 0):
                                    header = list(row)
                                else: 
                                    data.append(list(row))
                            formattedDict = formatToDict(header,data)
                            
                            c = 0
                            if(formattedDict):
                                for data in formattedDict:
                                    source = f"xlsx2splunk::{file}::{sheetname}"
                                    data['FilePath'] = source
                                    
                                    print(json.dumps(data)) # creates the event
                                    c+=1
                                
                            # set new state to checkpoint 
                        createCheckpoint(file,state)
                       
        except Exception as e:
            pass
            print(e)
    except Exception as e:
        print(e)
    



    