# -*- coding: utf-8 -*-
"""
Created on Tue Feb 23 10:06:58 2021

@author: Mervin Kall
"""
import math
# Vth_initialisation

#

#approach:
    #calculate Vth for given P and Q and initial V_poc --> only to determine phase, becasue Vth is already known
    
    #for given P and Q and Vth:
        # increase Vpoc_target to targeted value and re-calculate Vth. If abs (Vth) > initial Vth 
        # --> increase Q

    #once final Q is determined --> calculate C to reach Q at given voltage
def calc_capacity(SCC, X_R, Ppoc, Qpoc, Vpoc, Vbase, Vtarget):    
    
    Vth_new=99999999
    Qoffset=0
    SCC=SCC*1000000#scale to VA
    
    Zgrid=(Vbase*Vbase)/SCC
    R=Zgrid/math.sqrt((1+X_R*X_R))
    X=R*X_R*1j
    ZgridC=R+X
    
    Q=Qpoc*1000000
    P=Ppoc*1000000
    
    Upoc_target=Vpoc*Vbase
    
    S=Q*1j+P  
    Ipoc=S/( (math.sqrt(3)) * Vbase*Vpoc ) 
    Ipoc=Ipoc.conjugate()    
    
    Vth=abs(Vpoc*Vbase-Ipoc*ZgridC*math.sqrt(3))/Vbase
    
    while( (abs(Vth_new)/Vbase)>Vth+0.005):
        Q=(Qpoc+Qoffset)*1000000
        P=Ppoc*1000000
        #.calc_Vth_pu(X_R=1.13,SCR=1.42733,Sbase=161.84, Vbase=220000, Qpoc=0, Ppoc=140, Vpoc=1.02)
         
        
        Upoc_target=Vtarget*Vbase
            
        # Ip=P/ Upoc_target_rms /3
        # Iq=Q/ Upoc_target_rms /3
        # Ipoc=Ip+1j*Iq
        # Ipoc=Ipoc.conjugate()
        
        S=Q*1j+P  
        Ipoc=S/( (math.sqrt(3)) * Vbase*Vtarget ) 
        Ipoc=Ipoc.conjugate()
        
        
        Vth_new=Vtarget*Vbase-Ipoc*ZgridC*math.sqrt(3)
        #print(abs(Vth_new)/Vbase)
        
        Qoffset+=1  

    capacity=Qoffset*1000000/(2*3.14159265358*50*Vtarget*Vtarget*Vbase*Vbase) #calculate capacitor value
    print("Vth: "+str(abs(Vth_new)/Vbase))
    print("Q: "+str(Qoffset))
    print("Capacity: " +str(capacity))
    return  Qoffset, capacity





def main():
    calc_capacity(SCC=768, X_R=2.31, Ppoc=140, Qpoc=0.0, Vpoc=1.02, Vbase=220000, Vtarget=1.15)
    
if __name__ == '__main__':
    main()

