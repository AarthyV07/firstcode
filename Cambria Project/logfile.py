f = open('C:\\Users\\FC-\\Desktop\\logfile.txt',"r")
lines = f.readlines()
ipaddress = "9.37.65.139"
status = "failed"
time = "08:51:02"
message = "WARNING"

for line in lines:
    if message in line:
        linestrip = line.strip().split(' ')
        print(linestrip)



