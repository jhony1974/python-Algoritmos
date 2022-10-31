import os
import win32com.client
import pytz
import sys
from datetime import datetime, timedelta


path = os.path.expanduser("C:\TEMP\testcorreo")

ruta_logs = "C:\TEMP"#sys.argv[5] 
outputDir = "C:\TEMP" 

def writelog(e):
    #Escribir en log en caso de fallo
    fecha = datetime.now()
    namelog = 'Log_'+fecha.strftime('%d%m%Y')+'.log'    
    filelog = open(ruta_logs+"/"+namelog,"a")
    filelog.write(fecha.strftime('%d-%m-%Y %H:%M:%S')+' Script: DownloadEmailsOutlook Message: '+str(e)+'\n')
    filelog.close()

def checkstop():   
     stoper = os.path.exists(outputDir+'/data/stop.txt')
     if(stoper):
        sys.exit()


outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
#inbox = mapi.Folders['bandeja de entrada']
inbox = mapi.GetDefaultFolder(6).Folders["0. Monitoreo"]
#inbox = mapi.GetDefaultFolder(6).Folders["your_sub_folder"]
messages = inbox.Items

f_inicio="27-10-2022 12:00:00"
received_min = datetime.strptime(f_inicio, '%d-%m-%Y %H:%M:%S')
received_min = received_min.strftime('%d/%m/%Y %H:%M %p')
print(received_min)


    #Fecha Fin + 1 segundo
f_fin="27-10-2022 15:00:00"
received_max = datetime.strptime(f_fin, '%d-%m-%Y %H:%M:%S')
received_max = received_max + timedelta(seconds=1)
received_max = received_max.strftime('%d/%m/%Y %H:%M %p')
print(received_max)

    #Busqueda
messages = inbox.Items.Restrict("[ReceivedTime] > '" + received_min + "' and [ReceivedTime] <= '" + received_max + "'")         
messages.Sort("[ReceivedTime]", False)
    
count =0

for message in messages:
        try:
            checkstop()
            #Print para guardar los id de los mensajes en el archivo list_msg.txt            
            _id = message.EntryID[-11:]       

            #print('To:'+' '.join([str(item) for item in message.Recipients]))
            
           
            print(_id)        
            #Guardar el msg
            filename = os.path.join(outputDir, _id+'.msg')        
            message.SaveAs(filename)
            if(message.UnRead):
                message.UnRead = False
            count = count + 1
        except Exception as ex:
            print("Error leyendo mensajes: " + str(ex))
            writelog("Error al iterar mensajes, "+str(ex))

print("PYTHON: Mensajes encontrados = "+ str(count))
