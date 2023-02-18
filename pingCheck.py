import xlsxwriter
import re
import sys
import time
import socket
from subprocess import run

import ipaddress

#This file is where all output is logged
timestamp = time.ctime().replace(':', '.')

#This class allows for life and terminal logging
class Logger(object):
    def __init__(self):
        self.terminal = sys.stdout
        timestamp = time.ctime().replace(':', '.')
        self.logfile = open('log-' + timestamp + '.txt', 'w')

    def write(self, message):
        self.terminal.write(message)
        self.logfile.write(message)  

class CreatePingReport(object):
    def __init__(self):
        self.hostlist = open("hosts.txt").read().split('\n')
    

    def validateIpAddress(self, ip_string):
        try:
            ip_object = ipaddress.ip_address(ip_string)
            return True
        except ValueError:
            return False


    def createDictGoodPingData(self, hostPinged, shellData):

        if self.validateIpAddress(hostPinged):
            pingedIp = hostPinged
            dnsName = self.revrsDnsLookp(hostPinged)
        else:
            pingedIp = re.search("\[(.*)\]", shellData)
            pingedIp = pingedIp.group(1)
            dnsName = hostPinged

        latencyRegex = re.search("Average.=.(.*)ms", shellData)
        latency = latencyRegex.group(1)

        retrnDict = {"pingedIp": pingedIp, "dnsName": dnsName,
        "latency": int(latency), }

        return retrnDict

    def createDictBadPingData(self, hostPinged, shellData):
        
        #DNS failed
        if 'could not find host' in shellData:
            return {"pingedIp": "DNS could not resolve", 
            "dnsName": hostPinged, "latency": -1}

        if self.validateIpAddress(hostPinged):
            pingedIp = hostPinged
            dnsName = self.revrsDnsLookp(hostPinged)
        else:
            pingedIp = re.search("\[(.*)\]", shellData)
            pingedIp = pingedIp.group(1)
            dnsName = hostPinged

        retrnDict = {"pingedIp": pingedIp, "dnsName": dnsName,
        "latency": -1}

        return retrnDict


    def ping(self, host):

        shellOutput = run("ping " +host+ " -n 1", shell=True, text=True, capture_output=True )
        shellOutput = shellOutput.stdout

        if 'Reply from' in shellOutput:
            return self.createDictGoodPingData(host, shellOutput)
        else:
            return self.createDictBadPingData(host, shellOutput)


        
    def validateIpAddress(self, ip_string):
        try:
            ip_object = ipaddress.ip_address(ip_string)
            return True
        except ValueError:
            return False
        

    #Reverse DNS lookup
    def revrsDnsLookp(self, ipAddr):
        
        try:
            ipNslookup = socket.gethostbyaddr(ipAddr)
            if ipNslookup[0]: ipNsLookupRes = ipNslookup[0]
        except:
            ipNsLookupRes = 'Could not resolve'

        return ipNsLookupRes


    def createPingList(self):
        pingResultList = []

        for host in self.hostlist:
            pingData = self.ping(host)

            pingResultList.append(pingData)
                
        return pingResultList    

    
    def createXls(self):

        data = self.createPingList()

        #Create switch file and sheet
        xlsFile = xlsxwriter.Workbook('ping-report-' + timestamp + '.xlsx')
        xlsSheet = xlsFile.add_worksheet(timestamp)
        bold = xlsFile.add_format({'bold': True})
        boldRed = xlsFile.add_format({'bold': True, 'font_color': 'red'})
        
        #Create category roles from 'switchCols' list
        intRow = 0
        
        if data == None:
            dataError = ('No data.  Make sure the file is in the correct \
            format and the hosts.txt file is in the correct directory')
            xlsSheet.write(0, 2, dataError, bold)
            xlsFile.close() 
            exit()
        
        rowSwi = 0
        colSwi = 0
        xlsSheet.write(rowSwi, colSwi , "IP Address")
        colSwi += 1
        xlsSheet.write(rowSwi, colSwi, "DNS")
        colSwi += 1
        xlsSheet.write(rowSwi, colSwi, "Latency(ms)")
        colSwi += 1
        xlsSheet.write(rowSwi, colSwi, "Status")
        
        #Set begin row
        rowSwi = 2

        #Loop overs lists in list and put into xlsx file
        for host in data:

            if host["latency"] == -1:
                status = "failed"
            else:
                status = "success"

            #Reset column
            if colSwi != 0: colSwi = 0 
            #Add interface info before ARP.  Add then incre column
            xlsSheet.write(rowSwi, colSwi , host["pingedIp"])
            colSwi += 1
            xlsSheet.write(rowSwi, colSwi, host["dnsName"])
            colSwi += 1
            xlsSheet.write(rowSwi, colSwi, host["latency"])
            colSwi += 1
            xlsSheet.write(rowSwi, colSwi, status)
            
            
            rowSwi += 1

            
            xlsSheet.write # Not really sure I need this....
        
        xlsFile.close()  # Close file


def main():
    report = CreatePingReport()
    report.createXls()
    
    
    


if __name__ == "__main__":
    main()
