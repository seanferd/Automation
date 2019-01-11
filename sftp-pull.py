#############################################################
#                         sftp-pull.py 
#                         Sean R Ford
#                          01/11/2019
#
#   This script connects to an external SFTP server and 
#   saves the files it finds
#   into a locally accessible network share that everyone 
#   can access
#
##########################################################

import paramiko

#Connetion setup
host = "sftp.centreteksolutions.net"
port = 22
transport = paramiko.Transport((host, port))

password = "woXJA5fdU5P1"
username = "ced"
transport.connect(username = username, password = password)

sftp = paramiko.SFTPClient.from_transport(transport)

#Where we want to store the files we find
localpath = '\\\\umh.edu\\data\\Personnel_Payroll\\Salary\\CENSUS\\CEDStudents\\SFTP\\'
#Array of all the files we find on the server
afiles = sftp.listdir('/files')

#Loop through the files we found above and save them
for file in afiles:
    #add the filename to the output path
    saveName = localpath + file
    file = '/files/' + file
    sftp.get(file,saveName)

#Cleanup
sftp.close()
transport.close()