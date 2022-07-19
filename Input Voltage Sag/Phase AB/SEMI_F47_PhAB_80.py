import socket
import serial
import os
from time import sleep, gmtime, strftime
#For California Instruments MX Series


ser = serial.Serial()
#"OPEN COM4:9600,n,8,1,BIN,TB2048,RB2048"

ser.port ="COM4"
ser.baudrate = 9600
ser.open()
ser.is_open

MESSAGE = "*CLS"
ser.write(MESSAGE + '\r\n')
print(MESSAGE)
sleep(0.200)
MESSAGE = "MODE AC"
ser.write(MESSAGE+ '\r\n')
print(MESSAGE)
sleep(0.200)
MESSAGE = "FREQ 60"
ser.write(MESSAGE+ '\r\n')
print(MESSAGE)
sleep(0.200)
MESSAGE = "CURR 20"
ser.write(MESSAGE+ '\r\n')
print(MESSAGE)
sleep(0.200)
MESSAGE = "INST:COUP ALL"  #Change ALL Phase Voltage
ser.write(MESSAGE + '\r\n')
print(MESSAGE)
sleep(1)
MESSAGE = "VOLT 277"
ser.write(MESSAGE + '\r\n')
print(MESSAGE)
sleep(1)
MESSAGE = "OUTP ON"
ser.write(MESSAGE+ '\r\n')
print(MESSAGE)
sleep(20)

MESSAGE = "INST:COUP NONE"  #Change Individual Phase Voltage
ser.write(MESSAGE + '\r\n')
print(MESSAGE)
sleep(1)

MESSAGE = "INST:SEL B"
ser.write(MESSAGE + '\r\n')
print(MESSAGE)
sleep(1)
MESSAGE = "VOLT 277"        #To send Trigger
ser.write(MESSAGE + '\r\n')
print(MESSAGE)
sleep(0.001)
MESSAGE = "PHAS 273"
ser.write(MESSAGE+ '\r\n')
print(MESSAGE)
sleep(1)
MESSAGE = "PHAS 240"
ser.write(MESSAGE+ '\r\n')
print(MESSAGE)
sleep(10)
#MESSAGE = "OUTP OFF"
#ser.write(MESSAGE+ '\r\n')
#print(MESSAGE)
#sleep(0.5)

ser.close()
